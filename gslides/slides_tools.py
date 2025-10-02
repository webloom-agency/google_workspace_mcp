"""
Google Slides MCP Tools

This module provides MCP tools for interacting with Google Slides API.
"""

import logging
import asyncio
from typing import List, Dict, Any


from auth.service_decorator import require_google_service
from core.server import server
from core.utils import handle_http_errors
from core.comments import create_comment_tools

logger = logging.getLogger(__name__)


@server.tool()
@handle_http_errors("create_presentation", service_type="slides")
@require_google_service("slides", "slides")
async def create_presentation(
    service,
    user_google_email: str,
    title: str = "Untitled Presentation"
) -> str:
    """
    Create a new Google Slides presentation.

    Args:
        user_google_email (str): The user's Google email address. Required.
        title (str): The title for the new presentation. Defaults to "Untitled Presentation".

    Returns:
        str: Details about the created presentation including ID and URL.
    """
    logger.info(f"[create_presentation] Invoked. Email: '{user_google_email}', Title: '{title}'")

    body = {
        'title': title
    }

    result = await asyncio.to_thread(
        service.presentations().create(body=body).execute
    )

    presentation_id = result.get('presentationId')
    presentation_url = f"https://docs.google.com/presentation/d/{presentation_id}/edit"

    confirmation_message = f"""Presentation Created Successfully for {user_google_email}:
- Title: {title}
- Presentation ID: {presentation_id}
- URL: {presentation_url}
- Slides: {len(result.get('slides', []))} slide(s) created"""

    logger.info(f"Presentation created successfully for {user_google_email}")
    return confirmation_message


@server.tool()
@handle_http_errors("get_presentation", is_read_only=True, service_type="slides")
@require_google_service("slides", "slides_read")
async def get_presentation(
    service,
    user_google_email: str,
    presentation_id: str,
    include_text: bool = True,
    include_notes: bool = True,
    max_chars_per_slide: int = 4000,
    max_slides: int | None = None
) -> str:
    """
    Get details and content of a Google Slides presentation.

    Args:
        user_google_email (str): The user's Google email address. Required.
        presentation_id (str): The ID of the presentation to retrieve.
        include_text (bool): Whether to extract visible text from shapes/tables. Default True.
        include_notes (bool): Whether to extract speaker notes text. Default True.
        max_chars_per_slide (int): Safety cap to avoid overly long outputs per slide. Default 4000.
        max_slides (Optional[int]): If set, limit processing to the first N slides.

    Returns:
        str: Details plus per-slide extracted content.
    """
    logger.info(f"[get_presentation] Invoked. Email: '{user_google_email}', ID: '{presentation_id}'")

    presentation = await asyncio.to_thread(
        service.presentations().get(presentationId=presentation_id).execute
    )

    title = presentation.get('title', 'Untitled')
    slides = presentation.get('slides', [])
    page_size = presentation.get('pageSize', {})

    def extract_text_from_text_elements(text_elements: list[dict]) -> str:
        lines = []
        for te in text_elements or []:
            text_run = te.get('textRun')
            if text_run:
                content = text_run.get('content', '')
                if content:
                    lines.append(content)
        return ''.join(lines)

    def extract_shape_text(shape: dict) -> str:
        text_content = []
        text = (shape or {}).get('text', {})
        for pe in text.get('textElements', []) or []:
            tr = pe.get('textRun')
            if tr and 'content' in tr:
                text_content.append(tr['content'])
        return ''.join(text_content).strip()

    def extract_table_text(table: dict) -> str:
        cell_text_parts = []
        rows = table.get('rows', 0)
        cols = table.get('columns', 0)
        table_cells = table.get('tableRows', []) or []
        # Some APIs expose rows via tableRows, others via a grid-like structure
        for row in table_cells:
            cells = row.get('tableCells', []) or []
            row_parts = []
            for cell in cells:
                cell_text = []
                for ce in cell.get('text', {}).get('textElements', []) or []:
                    tr = ce.get('textRun')
                    if tr and 'content' in tr:
                        cell_text.append(tr['content'])
                row_parts.append(''.join(cell_text).strip())
            if row_parts:
                cell_text_parts.append(' | '.join(row_parts))
        # Fallback if above structure not present
        if not cell_text_parts and rows and cols:
            cell_text_parts.append(f"[{rows}x{cols} table content not parsed]")
        return '\n'.join(cell_text_parts).strip()

    def truncate(text: str, limit: int) -> str:
        if limit and len(text) > limit:
            return text[:limit] + "\n[...truncated...]"
        return text

    slide_outputs = []
    total_slides = len(slides)
    process_count = min(total_slides, max_slides) if max_slides else total_slides

    for index, slide in enumerate(slides[:process_count], start=1):
        slide_id = slide.get('objectId', 'Unknown')
        elements = slide.get('pageElements', []) or []

        visible_text_parts = []
        if include_text:
            for el in elements:
                if 'shape' in el:
                    text_value = extract_shape_text(el.get('shape', {}))
                    if text_value:
                        visible_text_parts.append(text_value)
                elif 'table' in el:
                    table_text = extract_table_text(el.get('table', {}))
                    if table_text:
                        visible_text_parts.append(table_text)

        notes_text = ''
        if include_notes:
            notes_page = slide.get('slideProperties', {}).get('notesPage') or slide.get('notesPage')
            if not notes_page:
                # Some API responses nest notes at top-level of slide
                notes_page = slide.get('notesPage')
            if notes_page:
                # The notes page contains a shape with the notes content
                notes_elements = (notes_page.get('notesProperties', {}) or {}).get('speakerNotesObjectId')
                # More robustly iterate all page elements of notesPage
                for npe in (notes_page.get('pageElements') or []):
                    shape = npe.get('shape')
                    if shape:
                        candidate = extract_shape_text(shape)
                        if candidate:
                            notes_text += candidate
                notes_text = notes_text.strip()

        slide_text = ''
        if visible_text_parts:
            slide_text += ("\n".join(visible_text_parts)).strip()
        if notes_text:
            slide_text += ("\n\n--- SPEAKER NOTES ---\n" + notes_text)

        slide_text = truncate(slide_text, max_chars_per_slide) if slide_text else ''

        # Build section output for this slide
        header = f"Slide {index}/{total_slides} (ID: {slide_id})"
        if slide_text:
            slide_outputs.append(f"{header}\n{slide_text}")
        else:
            slide_outputs.append(f"{header}\n[No extractable text]")

    summary_header = (
        f"Presentation Details for {user_google_email}:\n"
        f"- Title: {title}\n"
        f"- Presentation ID: {presentation_id}\n"
        f"- URL: https://docs.google.com/presentation/d/{presentation_id}/edit\n"
        f"- Total Slides: {len(slides)}\n"
        f"- Page Size: {page_size.get('width', {}).get('magnitude', 'Unknown')} x {page_size.get('height', {}).get('magnitude', 'Unknown')} {page_size.get('width', {}).get('unit', '')}\n"
    )

    content_intro = "\n--- CONTENT (first {n} slides) ---\n".format(n=process_count)
    full_output = summary_header + content_intro + ("\n\n".join(slide_outputs) if slide_outputs else "No slides found")

    logger.info(f"Presentation retrieved successfully with content for {user_google_email}")
    return full_output


@server.tool()
@handle_http_errors("batch_update_presentation", service_type="slides")
@require_google_service("slides", "slides")
async def batch_update_presentation(
    service,
    user_google_email: str,
    presentation_id: str,
    requests: List[Dict[str, Any]]
) -> str:
    """
    Apply batch updates to a Google Slides presentation.

    Args:
        user_google_email (str): The user's Google email address. Required.
        presentation_id (str): The ID of the presentation to update.
        requests (List[Dict[str, Any]]): List of update requests to apply.

    Returns:
        str: Details about the batch update operation results.
    """
    logger.info(f"[batch_update_presentation] Invoked. Email: '{user_google_email}', ID: '{presentation_id}', Requests: {len(requests)}")

    body = {
        'requests': requests
    }

    result = await asyncio.to_thread(
        service.presentations().batchUpdate(
            presentationId=presentation_id,
            body=body
        ).execute
    )

    replies = result.get('replies', [])

    confirmation_message = f"""Batch Update Completed for {user_google_email}:
- Presentation ID: {presentation_id}
- URL: https://docs.google.com/presentation/d/{presentation_id}/edit
- Requests Applied: {len(requests)}
- Replies Received: {len(replies)}"""

    if replies:
        confirmation_message += "\n\nUpdate Results:"
        for i, reply in enumerate(replies, 1):
            if 'createSlide' in reply:
                slide_id = reply['createSlide'].get('objectId', 'Unknown')
                confirmation_message += f"\n  Request {i}: Created slide with ID {slide_id}"
            elif 'createShape' in reply:
                shape_id = reply['createShape'].get('objectId', 'Unknown')
                confirmation_message += f"\n  Request {i}: Created shape with ID {shape_id}"
            else:
                confirmation_message += f"\n  Request {i}: Operation completed"

    logger.info(f"Batch update completed successfully for {user_google_email}")
    return confirmation_message


@server.tool()
@handle_http_errors("get_page", is_read_only=True, service_type="slides")
@require_google_service("slides", "slides_read")
async def get_page(
    service,
    user_google_email: str,
    presentation_id: str,
    page_object_id: str
) -> str:
    """
    Get details about a specific page (slide) in a presentation.

    Args:
        user_google_email (str): The user's Google email address. Required.
        presentation_id (str): The ID of the presentation.
        page_object_id (str): The object ID of the page/slide to retrieve.

    Returns:
        str: Details about the specific page including elements and layout.
    """
    logger.info(f"[get_page] Invoked. Email: '{user_google_email}', Presentation: '{presentation_id}', Page: '{page_object_id}'")

    result = await asyncio.to_thread(
        service.presentations().pages().get(
            presentationId=presentation_id,
            pageObjectId=page_object_id
        ).execute
    )

    page_type = result.get('pageType', 'Unknown')
    page_elements = result.get('pageElements', [])

    elements_info = []
    for element in page_elements:
        element_id = element.get('objectId', 'Unknown')
        if 'shape' in element:
            shape_type = element['shape'].get('shapeType', 'Unknown')
            elements_info.append(f"  Shape: ID {element_id}, Type: {shape_type}")
        elif 'table' in element:
            table = element['table']
            rows = table.get('rows', 0)
            cols = table.get('columns', 0)
            elements_info.append(f"  Table: ID {element_id}, Size: {rows}x{cols}")
        elif 'line' in element:
            line_type = element['line'].get('lineType', 'Unknown')
            elements_info.append(f"  Line: ID {element_id}, Type: {line_type}")
        else:
            elements_info.append(f"  Element: ID {element_id}, Type: Unknown")

    confirmation_message = f"""Page Details for {user_google_email}:
- Presentation ID: {presentation_id}
- Page ID: {page_object_id}
- Page Type: {page_type}
- Total Elements: {len(page_elements)}

Page Elements:
{chr(10).join(elements_info) if elements_info else '  No elements found'}"""

    logger.info(f"Page retrieved successfully for {user_google_email}")
    return confirmation_message


@server.tool()
@handle_http_errors("get_page_thumbnail", is_read_only=True, service_type="slides")
@require_google_service("slides", "slides_read")
async def get_page_thumbnail(
    service,
    user_google_email: str,
    presentation_id: str,
    page_object_id: str,
    thumbnail_size: str = "MEDIUM"
) -> str:
    """
    Generate a thumbnail URL for a specific page (slide) in a presentation.

    Args:
        user_google_email (str): The user's Google email address. Required.
        presentation_id (str): The ID of the presentation.
        page_object_id (str): The object ID of the page/slide.
        thumbnail_size (str): Size of thumbnail ("LARGE", "MEDIUM", "SMALL"). Defaults to "MEDIUM".

    Returns:
        str: URL to the generated thumbnail image.
    """
    logger.info(f"[get_page_thumbnail] Invoked. Email: '{user_google_email}', Presentation: '{presentation_id}', Page: '{page_object_id}', Size: '{thumbnail_size}'")

    result = await asyncio.to_thread(
        service.presentations().pages().getThumbnail(
            presentationId=presentation_id,
            pageObjectId=page_object_id,
            thumbnailProperties_thumbnailSize=thumbnail_size,
            thumbnailProperties_mimeType='PNG'
        ).execute
    )

    thumbnail_url = result.get('contentUrl', '')

    confirmation_message = f"""Thumbnail Generated for {user_google_email}:
- Presentation ID: {presentation_id}
- Page ID: {page_object_id}
- Thumbnail Size: {thumbnail_size}
- Thumbnail URL: {thumbnail_url}

You can view or download the thumbnail using the provided URL."""

    logger.info(f"Thumbnail generated successfully for {user_google_email}")
    return confirmation_message


# Create comment management tools for slides
_comment_tools = create_comment_tools("presentation", "presentation_id")
read_presentation_comments = _comment_tools['read_comments']
create_presentation_comment = _comment_tools['create_comment']
reply_to_presentation_comment = _comment_tools['reply_to_comment']
resolve_presentation_comment = _comment_tools['resolve_comment']

# Aliases for backwards compatibility and intuitive naming
read_slide_comments = read_presentation_comments
create_slide_comment = create_presentation_comment
reply_to_slide_comment = reply_to_presentation_comment
resolve_slide_comment = resolve_presentation_comment