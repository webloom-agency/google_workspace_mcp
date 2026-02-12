"""
Google Drive Helper Functions

Shared utilities for Google Drive operations including permission checking.
"""
import re
import asyncio
from typing import List, Dict, Any, Optional


def check_public_link_permission(permissions: List[Dict[str, Any]]) -> bool:
    """
    Check if file has 'anyone with the link' permission.
    
    Args:
        permissions: List of permission objects from Google Drive API
        
    Returns:
        bool: True if file has public link sharing enabled
    """
    return any(
        p.get('type') == 'anyone' and p.get('role') in ['reader', 'writer', 'commenter']
        for p in permissions
    )


def format_public_sharing_error(file_name: str, file_id: str) -> str:
    """
    Format error message for files without public sharing.
    
    Args:
        file_name: Name of the file
        file_id: Google Drive file ID
        
    Returns:
        str: Formatted error message
    """
    return (
        f"❌ Permission Error: '{file_name}' not shared publicly. "
        f"Set 'Anyone with the link' → 'Viewer' in Google Drive sharing. "
        f"File: https://drive.google.com/file/d/{file_id}/view"
    )


def get_drive_image_url(file_id: str) -> str:
    """
    Get the correct Drive URL format for publicly shared images.
    
    Args:
        file_id: Google Drive file ID
        
    Returns:
        str: URL for embedding Drive images
    """
    return f"https://drive.google.com/uc?export=view&id={file_id}"


# Precompiled regex patterns for Drive query detection
DRIVE_QUERY_PATTERNS = [
    re.compile(r'\b\w+\s*(=|!=|>|<)\s*[\'"].*?[\'"]', re.IGNORECASE),  # field = 'value'
    re.compile(r'\b\w+\s*(=|!=|>|<)\s*\d+', re.IGNORECASE),            # field = number
    re.compile(r'\bcontains\b', re.IGNORECASE),                         # contains operator
    re.compile(r'\bin\s+parents\b', re.IGNORECASE),                     # in parents
    re.compile(r'\bhas\s*\{', re.IGNORECASE),                          # has {properties}
    re.compile(r'\btrashed\s*=\s*(true|false)\b', re.IGNORECASE),      # trashed=true/false
    re.compile(r'\bstarred\s*=\s*(true|false)\b', re.IGNORECASE),      # starred=true/false
    re.compile(r'[\'"][^\'"]+[\'"]\s+in\s+parents', re.IGNORECASE),    # 'parentId' in parents
    re.compile(r'\bfullText\s+contains\b', re.IGNORECASE),             # fullText contains
    re.compile(r'\bname\s*(=|contains)\b', re.IGNORECASE),             # name = or name contains
    re.compile(r'\bmimeType\s*(=|!=)\b', re.IGNORECASE),               # mimeType operators
]


def build_drive_list_params(
    query: str,
    page_size: int,
    drive_id: Optional[str] = None,
    include_items_from_all_drives: bool = True,
    corpora: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Helper function to build common list parameters for Drive API calls.

    Args:
        query: The search query string
        page_size: Maximum number of items to return
        drive_id: Optional shared drive ID
        include_items_from_all_drives: Whether to include items from all drives
        corpora: Optional corpus specification

    Returns:
        Dictionary of parameters for Drive API list calls
    """
    list_params = {
        "q": query,
        "pageSize": page_size,
        "fields": "nextPageToken, files(id, name, mimeType, webViewLink, iconLink, modifiedTime, size)",
        "supportsAllDrives": True,
        "includeItemsFromAllDrives": include_items_from_all_drives,
        "orderBy": "modifiedTime desc",
    }

    if drive_id:
        list_params["driveId"] = drive_id
        if corpora:
            list_params["corpora"] = corpora
        else:
            list_params["corpora"] = "drive"
    elif corpora:
        list_params["corpora"] = corpora

    return list_params


async def find_folder_by_name_pattern(
    service,
    name_pattern: str,
    exact_match: bool = False,
    user_email: Optional[str] = None,
    parent_folder_id: Optional[str] = None,
) -> Optional[Dict[str, str]]:
    """
    Search for a folder by name pattern (case-insensitive contains by default).
    Returns the most recently modified matching folder.
    
    Args:
        service: Google Drive service instance
        name_pattern: String to search for in folder names
        exact_match: If True, requires exact name match (case-insensitive)
        user_email: Optional user email for logging
        parent_folder_id: Optional parent folder ID to search within (recursive). If None, searches all Drive.
        
    Returns:
        Dict with 'id', 'name', and 'webViewLink' of the folder, or None if not found
    """
    import logging
    logger = logging.getLogger(__name__)
    
    # Escape single quotes in the pattern
    escaped_pattern = name_pattern.replace("'", "\\'")
    
    # Build the base query for folder name matching
    if exact_match:
        query = f"mimeType='application/vnd.google-apps.folder' and name='{escaped_pattern}' and trashed=false"
    else:
        query = f"mimeType='application/vnd.google-apps.folder' and name contains '{escaped_pattern}' and trashed=false"
    
    # Add parent folder constraint if specified (this searches recursively)
    if parent_folder_id:
        escaped_parent_id = parent_folder_id.replace("'", "\\'")
        query += f" and '{escaped_parent_id}' in parents"
        logger.info(f"[find_folder_by_name_pattern] Searching for folder with pattern: '{name_pattern}' within parent folder: {parent_folder_id} (exact={exact_match})")
    else:
        logger.info(f"[find_folder_by_name_pattern] Searching for folder with pattern: '{name_pattern}' across all Drive (exact={exact_match})")
    
    try:
        results = await asyncio.to_thread(
            service.files().list(
                q=query,
                pageSize=10,
                fields="files(id, name, webViewLink, modifiedTime, parents)",
                orderBy="modifiedTime desc",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            ).execute
        )
        
        folders = results.get('files', [])
        
        if not folders:
            search_scope = f"within parent {parent_folder_id}" if parent_folder_id else "in all Drive"
            logger.warning(f"[find_folder_by_name_pattern] No folders found matching '{name_pattern}' {search_scope}")
            return None
        
        # Return the most recently modified folder (first in list due to orderBy)
        best_match = folders[0]
        logger.info(f"[find_folder_by_name_pattern] Found folder: '{best_match['name']}' (ID: {best_match['id']})")
        
        if len(folders) > 1:
            logger.info(f"[find_folder_by_name_pattern] Found {len(folders)} matching folders, using most recent")
        
        return {
            'id': best_match['id'],
            'name': best_match['name'],
            'webViewLink': best_match.get('webViewLink', '')
        }
        
    except Exception as e:
        logger.error(f"[find_folder_by_name_pattern] Error searching for folder: {e}")
        return None


async def move_file_to_folder(
    service,
    file_id: str,
    folder_id: str,
    file_name: Optional[str] = None,
) -> bool:
    """
    Move a file to a specific folder in Google Drive.
    
    Args:
        service: Google Drive service instance
        file_id: ID of the file to move
        folder_id: ID of the destination folder
        file_name: Optional file name for logging
        
    Returns:
        bool: True if successful, False otherwise
    """
    import logging
    logger = logging.getLogger(__name__)
    
    logger.info(f"[move_file_to_folder] Moving file {file_id} to folder {folder_id}")
    
    try:
        # Get current parents
        file = await asyncio.to_thread(
            service.files().get(
                fileId=file_id,
                fields='parents',
                supportsAllDrives=True
            ).execute
        )
        
        previous_parents = ",".join(file.get('parents', []))
        
        # Move the file to the new folder
        await asyncio.to_thread(
            service.files().update(
                fileId=file_id,
                addParents=folder_id,
                removeParents=previous_parents,
                fields='id, parents',
                supportsAllDrives=True
            ).execute
        )
        
        logger.info(f"[move_file_to_folder] Successfully moved file to folder {folder_id}")
        return True
        
    except Exception as e:
        logger.error(f"[move_file_to_folder] Error moving file: {e}")
        return False


async def create_folder(
    service,
    folder_name: str,
    parent_folder_id: Optional[str] = None,
) -> Optional[Dict[str, str]]:
    """
    Create a new folder in Google Drive.
    
    Args:
        service: Google Drive service instance
        folder_name: Name of the folder to create
        parent_folder_id: Optional parent folder ID. If None, creates in My Drive root.
        
    Returns:
        Dict with 'id', 'name', and 'webViewLink' of the created folder, or None if failed
    """
    import logging
    logger = logging.getLogger(__name__)
    
    logger.info(f"[create_folder] Creating folder '{folder_name}' in parent {parent_folder_id or 'root'}")
    
    try:
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        
        if parent_folder_id:
            file_metadata['parents'] = [parent_folder_id]
        
        folder = await asyncio.to_thread(
            service.files().create(
                body=file_metadata,
                fields='id, name, webViewLink',
                supportsAllDrives=True
            ).execute
        )
        
        logger.info(f"[create_folder] Successfully created folder '{folder_name}' (ID: {folder['id']})")
        
        return {
            'id': folder['id'],
            'name': folder['name'],
            'webViewLink': folder.get('webViewLink', '')
        }
        
    except Exception as e:
        logger.error(f"[create_folder] Error creating folder: {e}")
        return None


async def find_or_create_folder_path(
    service,
    folder_path: List[str],
    root_folder_id: Optional[str] = None,
    create_missing: bool = True,
) -> Optional[Dict[str, Any]]:
    """
    Navigate through a folder path (e.g., ["CLIENTS", "xxx.fr", "SEO"]) by searching for each folder.
    Creates missing folders if create_missing is True.
    
    Args:
        service: Google Drive service instance
        folder_path: List of folder names to navigate through (in order). Uses exact name matching.
        root_folder_id: Optional starting folder ID. If None, starts from My Drive root.
        create_missing: If True, creates folders that don't exist. If False, returns None if path doesn't exist.
        
    Returns:
        Dict with 'id', 'name', 'webViewLink', and 'path_summary' of the final folder, or None if failed
    """
    import logging
    logger = logging.getLogger(__name__)
    
    if not folder_path:
        logger.warning("[find_or_create_folder_path] Empty folder path provided")
        return None
    
    # Replace empty/whitespace-only folder names with "Untitled" to avoid creating unnamed folders
    sanitized_path = []
    for i, name in enumerate(folder_path):
        if not name or not name.strip():
            logger.warning(
                f"[find_or_create_folder_path] Empty folder name at position {i + 1}, defaulting to 'Untitled'"
            )
            sanitized_path.append("Untitled")
        else:
            sanitized_path.append(name.strip())
    folder_path = sanitized_path
    
    logger.info(f"[find_or_create_folder_path] Navigating path: {' > '.join(folder_path)} (create_missing={create_missing})")
    
    # Use 'root' as the starting parent if no root_folder_id is provided
    # This ensures the first folder is searched in My Drive root, not anywhere in Drive
    current_parent_id = root_folder_id if root_folder_id else 'root'
    path_summary = []
    
    for i, folder_name in enumerate(folder_path):
        # Try to find existing folder with exact name match
        folder = await find_folder_by_name_pattern(
            service,
            folder_name,
            exact_match=True,  # Use exact match for folder path navigation
            parent_folder_id=current_parent_id
        )
        
        if folder:
            logger.info(f"[find_or_create_folder_path] Found existing folder: '{folder['name']}' (ID: {folder['id']})")
            path_summary.append(f"{folder['name']}")
            current_parent_id = folder['id']
        elif create_missing:
            # Create the folder with the exact name specified
            logger.info(f"[find_or_create_folder_path] Folder '{folder_name}' not found, creating it...")
            new_folder = await create_folder(
                service,
                folder_name,
                parent_folder_id=current_parent_id if current_parent_id != 'root' else None
            )
            
            if not new_folder:
                logger.error(f"[find_or_create_folder_path] Failed to create folder '{folder_name}'")
                return None
            
            logger.info(f"[find_or_create_folder_path] Created folder: '{new_folder['name']}' (ID: {new_folder['id']})")
            path_summary.append(f"{new_folder['name']} (created)")
            current_parent_id = new_folder['id']
            folder = new_folder
        else:
            logger.warning(f"[find_or_create_folder_path] Folder '{folder_name}' not found and create_missing=False")
            return None
    
    # Return the final folder in the path
    final_folder_info = {
        'id': current_parent_id,
        'name': folder['name'],
        'webViewLink': folder.get('webViewLink', ''),
        'path_summary': ' > '.join(path_summary)
    }
    
    logger.info(f"[find_or_create_folder_path] Successfully navigated to: {final_folder_info['path_summary']}")
    return final_folder_info