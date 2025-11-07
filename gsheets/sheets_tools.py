"""
Google Sheets MCP Tools

This module provides MCP tools for interacting with Google Sheets API.
"""

import logging
import asyncio
import json
from typing import List, Optional, Union, Dict, Any


from auth.service_decorator import require_google_service
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
    append_result = await asyncio.to_thread(
        service.spreadsheets()
        .values()
        .append(
            spreadsheetId=spreadsheet_id,
            range=sheet_name,  # appends after the last non-empty row
            valueInputOption=value_input_option,
            insertDataOption="INSERT_ROWS",
            body={"values": values_to_append},
        )
        .execute
    )

    updates = append_result.get("updates", {})
    updated_range = updates.get("updatedRange", sheet_name)
    updated_rows = updates.get("updatedRows", 0)
    updated_cells = updates.get("updatedCells", 0)

    return (
        f"Headers: {len(all_headers)} columns. Appended {updated_rows} rows / {updated_cells} cells to '{updated_range}'."
    )

@server.tool()
@handle_http_errors("create_spreadsheet", service_type="sheets")
@require_google_service("sheets", "sheets_write")
async def create_spreadsheet(
    service,
    user_google_email: str,
    title: str,
    sheet_names: Optional[List[str]] = None,
) -> str:
    """
    Creates a new Google Spreadsheet.

    Args:
        user_google_email (str): The user's Google email address. Required.
        title (str): The title of the new spreadsheet. Required.
        sheet_names (Optional[List[str]]): List of sheet names to create. If not provided, creates one sheet with default name.

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
        service.spreadsheets().create(body=spreadsheet_body).execute
    )

    spreadsheet_id = spreadsheet.get("spreadsheetId")
    spreadsheet_url = spreadsheet.get("spreadsheetUrl")

    text_output = (
        f"Successfully created spreadsheet '{title}' for {user_google_email}. "
        f"ID: {spreadsheet_id} | URL: {spreadsheet_url}"
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


# Create comment management tools for sheets
_comment_tools = create_comment_tools("spreadsheet", "spreadsheet_id")

# Extract and register the functions
read_sheet_comments = _comment_tools['read_comments']
create_sheet_comment = _comment_tools['create_comment']
reply_to_sheet_comment = _comment_tools['reply_to_comment']
resolve_sheet_comment = _comment_tools['resolve_comment']


