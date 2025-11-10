"""
Google Sheets MCP Tools

This module provides MCP tools for interacting with Google Sheets API.
"""

import logging
import asyncio
import json
from typing import List, Optional, Union, Dict, Any


from auth.service_decorator import require_google_service, require_multiple_services
from core.server import server
from core.utils import handle_http_errors
from core.comments import create_comment_tools

# Configure module logger
logger = logging.getLogger(__name__)


@server.tool()
@handle_http_errors("list_spreadsheets", is_read_only=True, service_type="sheets")
@require_google_service("drive", "drive_read")
async def list_spreadsheets(
    service,
    user_google_email: str,
    max_results: int = 25,
) -> str:
    """
    Lists spreadsheets from Google Drive that the user has access to.

    Args:
        user_google_email (str): The user's Google email address. Required.
        max_results (int): Maximum number of spreadsheets to return. Defaults to 25.

    Returns:
        str: A formatted list of spreadsheet files (name, ID, modified time).
    """
    logger.info(f"[list_spreadsheets] Invoked. Email: '{user_google_email}'")

    files_response = await asyncio.to_thread(
        service.files()
        .list(
            q="mimeType='application/vnd.google-apps.spreadsheet'",
            pageSize=max_results,
            fields="files(id,name,modifiedTime,webViewLink)",
            orderBy="modifiedTime desc",
        )
        .execute
    )

    files = files_response.get("files", [])
    if not files:
        return f"No spreadsheets found for {user_google_email}."

    spreadsheets_list = [
        f"- \"{file['name']}\" (ID: {file['id']}) | Modified: {file.get('modifiedTime', 'Unknown')} | Link: {file.get('webViewLink', 'No link')}"
        for file in files
    ]

    text_output = (
        f"Successfully listed {len(files)} spreadsheets for {user_google_email}:\n"
        + "\n".join(spreadsheets_list)
    )

    logger.info(f"Successfully listed {len(files)} spreadsheets for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("get_spreadsheet_info", is_read_only=True, service_type="sheets")
@require_google_service("sheets", "sheets_read")
async def get_spreadsheet_info(
    service,
    user_google_email: str,
    spreadsheet_id: str,
) -> str:
    """
    Gets information about a specific spreadsheet including its sheets.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet to get info for. Required.

    Returns:
        str: Formatted spreadsheet information including title and sheets list.
    """
    logger.info(f"[get_spreadsheet_info] Invoked. Email: '{user_google_email}', Spreadsheet ID: {spreadsheet_id}")

    spreadsheet = await asyncio.to_thread(
        service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
    )

    title = spreadsheet.get("properties", {}).get("title", "Unknown")
    sheets = spreadsheet.get("sheets", [])

    sheets_info = []
    for sheet in sheets:
        sheet_props = sheet.get("properties", {})
        sheet_name = sheet_props.get("title", "Unknown")
        sheet_id = sheet_props.get("sheetId", "Unknown")
        grid_props = sheet_props.get("gridProperties", {})
        rows = grid_props.get("rowCount", "Unknown")
        cols = grid_props.get("columnCount", "Unknown")

        sheets_info.append(
            f"  - \"{sheet_name}\" (ID: {sheet_id}) | Size: {rows}x{cols}"
        )

    text_output = (
        f"Spreadsheet: \"{title}\" (ID: {spreadsheet_id})\n"
        f"Sheets ({len(sheets)}):\n"
        + "\n".join(sheets_info) if sheets_info else "  No sheets found"
    )

    logger.info(f"Successfully retrieved info for spreadsheet {spreadsheet_id} for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("read_sheet_values", is_read_only=True, service_type="sheets")
@require_google_service("sheets", "sheets_read")
async def read_sheet_values(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str = "A1:Z1000",
) -> str:
    """
    Reads values from a specific range in a Google Sheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to read (e.g., "Sheet1!A1:D10", "A1:D10"). Defaults to "A1:Z1000".

    Returns:
        str: The formatted values from the specified range.
    """
    logger.info(f"[read_sheet_values] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}")

    result = await asyncio.to_thread(
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=range_name)
        .execute
    )

    values = result.get("values", [])
    if not values:
        return f"No data found in range '{range_name}' for {user_google_email}."

    # Format the output as a readable table
    formatted_rows = []
    for i, row in enumerate(values, 1):
        # Pad row with empty strings to show structure
        padded_row = row + [""] * max(0, len(values[0]) - len(row)) if values else row
        formatted_rows.append(f"Row {i:2d}: {padded_row}")

    text_output = (
        f"Successfully read {len(values)} rows from range '{range_name}' in spreadsheet {spreadsheet_id} for {user_google_email}:\n"
        + "\n".join(formatted_rows[:50])  # Limit to first 50 rows for readability
        + (f"\n... and {len(values) - 50} more rows" if len(values) > 50 else "")
    )

    logger.info(f"Successfully read {len(values)} rows for {user_google_email}.")
    return text_output


@server.tool()
@handle_http_errors("modify_sheet_values", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def modify_sheet_values(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str,
    values: Optional[Union[str, List[List[str]]]] = None,
    value_input_option: str = "USER_ENTERED",
    clear_values: bool = False,
) -> str:
    """
    Modifies values in a specific range of a Google Sheet - can write, update, or clear values.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The range to modify (e.g., "Sheet1!A1:D10", "A1:D10"). Required.
        values (Optional[Union[str, List[List[str]]]]): 2D array of values to write/update. Can be a JSON string or Python list. Required unless clear_values=True.
        value_input_option (str): How to interpret input values ("RAW" or "USER_ENTERED"). Defaults to "USER_ENTERED".
        clear_values (bool): If True, clears the range instead of writing values. Defaults to False.

    Returns:
        str: Confirmation message of the successful modification operation.
    """
    operation = "clear" if clear_values else "write"
    logger.info(f"[modify_sheet_values] Invoked. Operation: {operation}, Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}")

    # Parse values if it's a JSON string (MCP passes parameters as JSON strings)
    if values is not None and isinstance(values, str):
        try:
            parsed_values = json.loads(values)
            if not isinstance(parsed_values, list):
                raise ValueError(f"Values must be a list, got {type(parsed_values).__name__}")
            # Validate it's a list of lists
            for i, row in enumerate(parsed_values):
                if not isinstance(row, list):
                    raise ValueError(f"Row {i} must be a list, got {type(row).__name__}")
            values = parsed_values
            logger.info(f"[modify_sheet_values] Parsed JSON string to Python list with {len(values)} rows")
        except json.JSONDecodeError as e:
            raise Exception(f"Invalid JSON format for values: {e}")
        except ValueError as e:
            raise Exception(f"Invalid values structure: {e}")

    if not clear_values and not values:
        raise Exception("Either 'values' must be provided or 'clear_values' must be True.")

    if clear_values:
        result = await asyncio.to_thread(
            service.spreadsheets()
            .values()
            .clear(spreadsheetId=spreadsheet_id, range=range_name)
            .execute
        )

        cleared_range = result.get("clearedRange", range_name)
        text_output = f"Successfully cleared range '{cleared_range}' in spreadsheet {spreadsheet_id} for {user_google_email}."
        logger.info(f"Successfully cleared range '{cleared_range}' for {user_google_email}.")
    else:
        body = {"values": values}

        result = await asyncio.to_thread(
            service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=range_name,
                valueInputOption=value_input_option,
                body=body,
            )
            .execute
        )

        updated_cells = result.get("updatedCells", 0)
        updated_rows = result.get("updatedRows", 0)
        updated_columns = result.get("updatedColumns", 0)

        text_output = (
            f"Successfully updated range '{range_name}' in spreadsheet {spreadsheet_id} for {user_google_email}. "
            f"Updated: {updated_cells} cells, {updated_rows} rows, {updated_columns} columns."
        )
        logger.info(f"Successfully updated {updated_cells} cells for {user_google_email}.")

    return text_output


@server.tool()
@handle_http_errors("append_sheet_values", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def append_sheet_values(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    range_name: str,
    values: Optional[Union[str, List[List[str]]]] = None,
    value_input_option: str = "USER_ENTERED",
    insert_data_option: str = "INSERT_ROWS",
) -> str:
    """
    Appends rows to a Google Sheet without overwriting existing data.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        range_name (str): The target range or sheet (e.g., "Sheet1", "Sheet1!A1"). Required.
        values (Optional[Union[str, List[List[str]]]]): 2D array of values to append. Can be a JSON string or Python list.
        value_input_option (str): How to interpret input values ("RAW" or "USER_ENTERED"). Defaults to "USER_ENTERED".
        insert_data_option (str): How to insert the data ("OVERWRITE" or "INSERT_ROWS"). Defaults to "INSERT_ROWS".

    Returns:
        str: Confirmation message including how many rows/cells were appended.
    """
    logger.info(
        f"[append_sheet_values] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Range: {range_name}"
    )

    # Parse values if it's a JSON string (MCP passes parameters as JSON strings)
    if values is not None and isinstance(values, str):
        try:
            parsed_values = json.loads(values)
            if not isinstance(parsed_values, list):
                raise ValueError(f"Values must be a list, got {type(parsed_values).__name__}")
            for i, row in enumerate(parsed_values):
                if not isinstance(row, list):
                    raise ValueError(f"Row {i} must be a list, got {type(row).__name__}")
            values = parsed_values
            logger.info(
                f"[append_sheet_values] Parsed JSON string to Python list with {len(values)} rows"
            )
        except json.JSONDecodeError as e:
            raise Exception(f"Invalid JSON format for values: {e}")
        except ValueError as e:
            raise Exception(f"Invalid values structure: {e}")

    if not values:
        raise Exception("'values' must be provided and be a non-empty 2D array.")

    body = {"values": values}

    result = await asyncio.to_thread(
        service.spreadsheets()
        .values()
        .append(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption=value_input_option,
            insertDataOption=insert_data_option,
            body=body,
        )
        .execute
    )

    updates = result.get("updates", {})
    updated_range = updates.get("updatedRange", range_name)
    updated_rows = updates.get("updatedRows", 0)
    updated_cells = updates.get("updatedCells", 0)

    text_output = (
        f"Successfully appended to '{updated_range}' in spreadsheet {spreadsheet_id} for {user_google_email}. "
        f"Appended: {updated_rows} rows, {updated_cells} cells."
    )

    logger.info(
        f"Successfully appended {updated_rows} rows and {updated_cells} cells for {user_google_email}."
    )
    return text_output


@server.tool()
@handle_http_errors("append_rows_by_headers", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def append_rows_by_headers(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_name: str,
    rows: Optional[Union[str, List[Dict[str, Any]]]] = None,
    value_input_option: str = "USER_ENTERED",
    write_headers_if_missing: bool = True,
) -> str:
    """
    Appends rows mapped by header names. Ensures appends happen at the end, without
    overwriting existing data. If the sheet has no headers, optionally creates them
    from the union of provided row keys. If new keys appear, extends the header row
    (creating new columns) and maps values accordingly.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        sheet_name (str): Target sheet name. Required.
        rows (Optional[Union[str, List[Dict[str, Any]]]]): List of row objects keyed by header name.
            Can be provided as JSON string or Python list. Required.
        value_input_option (str): "RAW" or "USER_ENTERED". Defaults to "USER_ENTERED".
        write_headers_if_missing (bool): If True, write headers when sheet is empty. Defaults to True.

    Returns:
        str: Summary of headers and rows appended.
    """
    logger.info(
        f"[append_rows_by_headers] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Sheet: {sheet_name}"
    )

    # Parse rows if provided as JSON string
    if rows is not None and isinstance(rows, str):
        try:
            parsed_rows = json.loads(rows)
            rows = parsed_rows
            logger.info(
                f"[append_rows_by_headers] Parsed JSON string to Python list with {len(rows) if isinstance(rows, list) else 0} items"
            )
        except json.JSONDecodeError as e:
            raise Exception(f"Invalid JSON format for rows: {e}")

    if not rows or not isinstance(rows, list):
        raise Exception("'rows' must be a non-empty list of objects keyed by headers.")

    # Validate list elements are dict-like
    for i, item in enumerate(rows):
        if not isinstance(item, dict):
            raise Exception(f"Row {i} must be an object keyed by header names.")

    # 1) Read existing header row
    header_result = await asyncio.to_thread(
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=f"{sheet_name}!1:1")
        .execute
    )

    existing_header_values = header_result.get("values", [])
    existing_headers: List[str] = existing_header_values[0] if existing_header_values else []

    # 2) Compute union of headers
    provided_keys: List[str] = []
    seen = set()
    for item in rows:
        for k in item.keys():
            if k not in seen:
                seen.add(k)
                provided_keys.append(k)

    all_headers: List[str] = list(existing_headers) if existing_headers else []
    for k in provided_keys:
        if k not in all_headers:
            all_headers.append(k)

    # 3) If no headers and permitted, or if new headers present, update header row
    need_write_headers = (not existing_headers and write_headers_if_missing) or (
        existing_headers and len(all_headers) > len(existing_headers)
    )

    if need_write_headers:
        await asyncio.to_thread(
            service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!1:1",
                valueInputOption=value_input_option,
                body={"values": [all_headers]},
            )
            .execute
        )
        logger.info(
            f"[append_rows_by_headers] Header row set/extended to {len(all_headers)} columns."
        )

    if not all_headers:
        raise Exception(
            "No headers exist and write_headers_if_missing is False; cannot map rows."
        )

    # 4) Map input objects to row lists aligned with all_headers
    values_to_append: List[List[Any]] = []
    for item in rows:
        mapped_row = [item.get(h, "") for h in all_headers]
        values_to_append.append(mapped_row)

    # 5) Append rows at the end of the sheet
    # Determine next available row by scanning column A
    col_a_values = await asyncio.to_thread(
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=f"{sheet_name}!A:A", majorDimension="ROWS")
        .execute
    )
    existing_rows = len(col_a_values.get("values", [])) if col_a_values.get("values") else 0
    # If sheet has headers, existing_rows >= 1; next_row is existing_rows + 1
    next_row = max(1, existing_rows + 1)

    CHUNK_SIZE = 5000  # rows per request to avoid large payload timeouts
    total_rows_appended = 0
    total_cells_appended = 0
    last_updated_range = f"{sheet_name}!A{next_row}"

    for start in range(0, len(values_to_append), CHUNK_SIZE):
        chunk = values_to_append[start : start + CHUNK_SIZE]
        # Compute the A1 range for this chunk starting row
        start_row_for_chunk = next_row + total_rows_appended
        target_range = f"{sheet_name}!A{start_row_for_chunk}"

        update_result = await asyncio.to_thread(
            service.spreadsheets()
            .values()
            .update(
                spreadsheetId=spreadsheet_id,
                range=target_range,
                valueInputOption=value_input_option,
                body={"values": chunk},
            )
            .execute
        )

        updated_rows = update_result.get("updatedRows", len(chunk))
        updated_cells = update_result.get("updatedCells", len(chunk) * len(all_headers))
        total_rows_appended += updated_rows
        total_cells_appended += updated_cells
        last_updated_range = update_result.get("updatedRange", target_range)

    return (
        f"Headers: {len(all_headers)} columns. Appended {total_rows_appended} rows / {total_cells_appended} cells to '{last_updated_range}'."
    )

@server.tool()
@handle_http_errors("create_spreadsheet", service_type="sheets")
@require_multiple_services([
    {"service_type": "sheets", "scopes": "sheets_write", "param_name": "sheets_service"},
    {"service_type": "drive", "scopes": "drive_file", "param_name": "drive_service"}
])
async def create_spreadsheet(
    sheets_service,
    drive_service,
    user_google_email: str,
    title: str,
    sheet_names: Optional[List[str]] = None,
    folder_id: Optional[str] = None,
    folder_name_contains: Optional[str] = None,
    search_within_folder_id: Optional[str] = None,
    folder_path: Optional[List[str]] = None,
    create_folders_if_missing: bool = True,
) -> str:
    """
    Creates a new Google Spreadsheet, optionally in a specific folder.

    Args:
        user_google_email (str): The user's Google email address. Required.
        title (str): The title of the new spreadsheet. Required.
        sheet_names (Optional[List[str]]): List of sheet names to create. If not provided, creates one sheet with default name.
        folder_id (Optional[str]): Specific folder ID to place the spreadsheet in.
        folder_name_contains (Optional[str]): Search for folder name containing this string. Uses the most recently modified match.
        search_within_folder_id (Optional[str]): When using folder_name_contains, limit search to within this parent folder. If not specified, searches all Drive.
        folder_path (Optional[List[str]]): Navigate through nested folders by name patterns (e.g., ["CLIENTS", "xxx.fr", "SEO"]). Searches for each folder in order, creating missing ones if create_folders_if_missing is True.
        create_folders_if_missing (bool): When using folder_path, create folders that don't exist. Defaults to True.

    Returns:
        str: Information about the newly created spreadsheet including ID and URL.
    """
    logger.info(f"[create_spreadsheet] Invoked. Email: '{user_google_email}', Title: {title}")

    spreadsheet_body = {
        "properties": {
            "title": title
        }
    }

    if sheet_names:
        spreadsheet_body["sheets"] = [
            {"properties": {"title": sheet_name}} for sheet_name in sheet_names
        ]

    spreadsheet = await asyncio.to_thread(
        sheets_service.spreadsheets().create(body=spreadsheet_body).execute
    )

    spreadsheet_id = spreadsheet.get("spreadsheetId")
    spreadsheet_url = spreadsheet.get("spreadsheetUrl")
    
    # Handle folder placement
    folder_info = ""
    target_folder_id = folder_id
    
    # Priority 1: folder_path (navigate through nested folders)
    if folder_path and not folder_id:
        from gdrive.drive_helpers import find_or_create_folder_path
        folder_result = await find_or_create_folder_path(
            drive_service,
            folder_path,
            root_folder_id=search_within_folder_id,
            create_missing=create_folders_if_missing
        )
        if folder_result:
            target_folder_id = folder_result['id']
            folder_info = f" | Path: {folder_result['path_summary']}"
        else:
            folder_info = f" | Warning: Could not navigate folder path {' > '.join(folder_path)}, created in My Drive"
    
    # Priority 2: folder_name_contains (simple search)
    elif folder_name_contains and not folder_id:
        from gdrive.drive_helpers import find_folder_by_name_pattern
        folder = await find_folder_by_name_pattern(
            drive_service,
            folder_name_contains,
            exact_match=False,
            user_email=user_google_email,
            parent_folder_id=search_within_folder_id
        )
        if folder:
            target_folder_id = folder['id']
            search_scope = f" within folder {search_within_folder_id}" if search_within_folder_id else ""
            folder_info = f" | Folder: '{folder['name']}' ({folder['id']}){search_scope}"
        else:
            search_scope = f" within folder {search_within_folder_id}" if search_within_folder_id else " in all Drive"
            folder_info = f" | Warning: No folder found matching '{folder_name_contains}'{search_scope}, created in My Drive"
    
    if target_folder_id:
        from gdrive.drive_helpers import move_file_to_folder
        move_success = await move_file_to_folder(
            drive_service,
            spreadsheet_id,
            target_folder_id,
            file_name=title
        )
        if move_success and not folder_info:
            folder_info = f" | Moved to folder: {target_folder_id}"

    text_output = (
        f"Successfully created spreadsheet '{title}' for {user_google_email}. "
        f"ID: {spreadsheet_id} | URL: {spreadsheet_url}{folder_info}"
    )

    logger.info(f"Successfully created spreadsheet for {user_google_email}. ID: {spreadsheet_id}")
    return text_output


@server.tool()
@handle_http_errors("create_sheet", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def create_sheet(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_name: str,
) -> str:
    """
    Creates a new sheet within an existing spreadsheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        spreadsheet_id (str): The ID of the spreadsheet. Required.
        sheet_name (str): The name of the new sheet. Required.

    Returns:
        str: Confirmation message of the successful sheet creation.
    """
    logger.info(f"[create_sheet] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, Sheet: {sheet_name}")

    request_body = {
        "requests": [
            {
                "addSheet": {
                    "properties": {
                        "title": sheet_name
                    }
                }
            }
        ]
    }

    response = await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body=request_body)
        .execute
    )

    sheet_id = response["replies"][0]["addSheet"]["properties"]["sheetId"]

    text_output = (
        f"Successfully created sheet '{sheet_name}' (ID: {sheet_id}) in spreadsheet {spreadsheet_id} for {user_google_email}."
    )

    logger.info(f"Successfully created sheet for {user_google_email}. Sheet ID: {sheet_id}")
    return text_output


# NEW TOOL: Deduplicate rows by key headers while keeping max/min on another column
@server.tool()
@handle_http_errors("deduplicate_rows_by_headers", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def deduplicate_rows_by_headers(
    service,
    user_google_email: str,
    spreadsheet_id: str,
    sheet_name: Optional[str] = None,
    key_headers: Union[str, List[str]] = None,
    sort_header: str = None,
    keep: str = "max",
    work_on_copy: bool = False,
    destination_sheet_name: Optional[str] = None,
    sheet_id: Optional[int] = None,
) -> str:
    """
    Deduplicates rows in a sheet by one or more key headers, keeping the row with
    the max (or min) value in the specified sort column.

    Sheet selection precedence:
      1) sheet_id if provided
      2) sheet_name (exact, then case/space-normalized)
      3) if exactly one sheet exists, use it
      4) otherwise error listing available sheets

    Fast path based on Sheets server-side operations:
      1) Sort rows (excluding header) by the sort column (DESC for max, ASC for min)
      2) Delete duplicates comparing only the key columns (keeps first occurrence)

    Args:
        user_google_email: The user's Google email address. Required.
        spreadsheet_id: Spreadsheet ID. Required.
        sheet_name: Target sheet title. Optional if sheet_id provided or only one sheet.
        key_headers: Header name or list of header names used as the dedupe key. Required.
        sort_header: Header name used to choose which row to keep. Required.
        keep: "max" or "min". Defaults to "max".
        work_on_copy: If True, duplicates the sheet first and operates on the copy.
        destination_sheet_name: Optional name for the copied sheet when work_on_copy=True.
        sheet_id: Optional numeric sheetId (preferred for robustness).

    Returns:
        Summary string indicating the target sheet, operation mode, and columns used.
    """
    logger.info(
        f"[deduplicate_rows_by_headers] Invoked. Email: '{user_google_email}', Spreadsheet: {spreadsheet_id}, SheetName: {sheet_name}, SheetId: {sheet_id}, keep: {keep}, work_on_copy: {work_on_copy}"
    )

    if key_headers is None or sort_header is None:
        raise Exception("'key_headers' and 'sort_header' are required.")

    # Normalize key headers
    key_headers_list: List[str] = [key_headers] if isinstance(key_headers, str) else list(key_headers)
    if not key_headers_list:
        raise Exception("'key_headers' must be a non-empty string or list of strings.")

    keep_normalized = keep.strip().lower()
    if keep_normalized not in ("max", "min"):
        raise Exception("'keep' must be either 'max' or 'min'.")

    # 1) Resolve sheetId and title
    spreadsheet = await asyncio.to_thread(
        service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute
    )

    sheets = spreadsheet.get("sheets", [])
    if not sheets:
        raise Exception(f"Spreadsheet {spreadsheet_id} has no sheets.")

    def _norm(t: str) -> str:
        return " ".join(t.split()).lower()

    target_sheet = None

    # Prefer sheet_id if provided
    if sheet_id is not None:
        for s in sheets:
            if s.get("properties", {}).get("sheetId") == sheet_id:
                target_sheet = s
                break
        if target_sheet is None:
            available = [(s.get("properties", {}).get("title", ""), s.get("properties", {}).get("sheetId")) for s in sheets]
            raise Exception(f"sheet_id {sheet_id} not found. Available: {available}")
    else:
        # Try by sheet_name if provided
        if sheet_name:
            for s in sheets:
                if s.get("properties", {}).get("title") == sheet_name:
                    target_sheet = s
                    break
            if target_sheet is None:
                normalized_input = _norm(sheet_name)
                candidates = [s for s in sheets if _norm(s.get("properties", {}).get("title", "")) == normalized_input]
                if len(candidates) == 1:
                    target_sheet = candidates[0]
        # If still none, use only sheet if there is exactly one
        if target_sheet is None:
            if len(sheets) == 1:
                target_sheet = sheets[0]
            else:
                available = [(s.get("properties", {}).get("title", ""), s.get("properties", {}).get("sheetId")) for s in sheets]
                raise Exception(
                    f"No sheet selector resolved. Provide 'sheet_id' or 'sheet_name'. Available: {available}"
                )

    effective_sheet_title = target_sheet.get("properties", {}).get("title")
    source_sheet_id = target_sheet.get("properties", {}).get("sheetId")
    grid_props = target_sheet.get("properties", {}).get("gridProperties", {})
    total_rows = grid_props.get("rowCount", 1000000)
    total_cols = grid_props.get("columnCount", 26)

    # 2) Read header row to map headers -> column indices
    header_result = await asyncio.to_thread(
        service.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=f"{effective_sheet_title}!1:1")
        .execute
    )
    header_values = header_result.get("values", [])
    headers: List[str] = header_values[0] if header_values else []
    if not headers:
        raise Exception("Header row (row 1) is empty; cannot map header names to columns.")

    def header_to_col_index_or_raise(header_name: str) -> int:
        try:
            return headers.index(header_name)
        except ValueError:
            raise Exception(f"Header '{header_name}' not found in sheet '{effective_sheet_title}'.")

    sort_col_index = header_to_col_index_or_raise(sort_header)
    key_col_indices = [header_to_col_index_or_raise(h) for h in key_headers_list]

    # 3) Prepare optional duplicate sheet step
    effective_sheet_id = source_sheet_id
    requests: List[Dict[str, Any]] = []

    if work_on_copy:
        requests.append(
            {
                "duplicateSheet": {
                    "sourceSheetId": source_sheet_id,
                    "insertSheetIndex": 0,
                    "newSheetName": destination_sheet_name or f"{effective_sheet_title} (dedup)"
                }
            }
        )

        # Execute duplication first to obtain new sheet id
        dup_response = await asyncio.to_thread(
            service.spreadsheets()
            .batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={"requests": requests}
            )
            .execute
        )
        requests = []

        replies = dup_response.get("replies", [])
        if not replies or "duplicateSheet" not in replies[0]:
            raise Exception("Failed to duplicate sheet before deduplication.")
        effective_sheet_id = replies[0]["duplicateSheet"]["properties"]["sheetId"]

    # 4) Build sort then delete-duplicates requests applied to data rows (exclude header)
    sort_order = "DESCENDING" if keep_normalized == "max" else "ASCENDING"

    data_range = {
        "sheetId": effective_sheet_id,
        "startRowIndex": 1,  # exclude header row
        "startColumnIndex": 0,
        "endRowIndex": total_rows,
        "endColumnIndex": total_cols,
    }

    requests.append(
        {
            "sortRange": {
                "range": data_range,
                "sortSpecs": [
                    {"dimensionIndex": sort_col_index, "sortOrder": sort_order}
                ],
            }
        }
    )

    comparison_columns = [
        {
            "sheetId": effective_sheet_id,
            "dimension": "COLUMNS",
            "startIndex": idx,
            "endIndex": idx + 1,
        }
        for idx in key_col_indices
    ]

    requests.append(
        {
            "deleteDuplicates": {
                "range": data_range,
                "comparisonColumns": comparison_columns,
            }
        }
    )

    _ = await asyncio.to_thread(
        service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests})
        .execute
    )

    target_title = effective_sheet_title if not work_on_copy else (destination_sheet_name or f"{effective_sheet_title} (dedup)")
    return (
        f"Deduplicated sheet '{target_title}' by keys {key_headers_list}, keeping {keep_normalized} of '{sort_header}'."
    )

# Create comment management tools for sheets
_comment_tools = create_comment_tools("spreadsheet", "spreadsheet_id")

# Extract and register the functions
read_sheet_comments = _comment_tools['read_comments']
create_sheet_comment = _comment_tools['create_comment']
reply_to_sheet_comment = _comment_tools['reply_to_comment']
resolve_sheet_comment = _comment_tools['resolve_comment']


