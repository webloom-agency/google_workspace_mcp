"""
Tool Registry for Conditional Tool Registration

This module provides a registry system that allows tools to be conditionally registered
based on tier configuration, replacing direct @server.tool() decorators.
"""

import logging
import functools
import traceback
from typing import Set, Optional, Callable

logger = logging.getLogger(__name__)

# Global registry of enabled tools
_enabled_tools: Optional[Set[str]] = None

def set_enabled_tools(tool_names: Optional[Set[str]]):
    """Set the globally enabled tools."""
    global _enabled_tools
    _enabled_tools = tool_names

def get_enabled_tools() -> Optional[Set[str]]:
    """Get the set of enabled tools, or None if all tools are enabled."""
    return _enabled_tools

def is_tool_enabled(tool_name: str) -> bool:
    """Check if a specific tool is enabled."""
    if _enabled_tools is None:
        return True  # All tools enabled by default
    return tool_name in _enabled_tools

def conditional_tool(server, tool_name: str):
    """
    Decorator that conditionally registers a tool based on the enabled tools set.
    
    Args:
        server: The FastMCP server instance
        tool_name: The name of the tool to register
    
    Returns:
        Either the registered tool decorator or a no-op decorator
    """
    def decorator(func: Callable) -> Callable:
        if is_tool_enabled(tool_name):
            logger.debug(f"Registering tool: {tool_name}")
            return server.tool()(func)
        else:
            logger.debug(f"Skipping tool registration: {tool_name}")
            return func
    
    return decorator

def global_exception_handler(func: Callable) -> Callable:
    """
    Global exception handler that prevents any uncaught exceptions from crashing the MCP server.
    This ensures that all exceptions are properly logged and returned as error messages.
    
    This is a safety net - most exceptions should be caught by the tool's own error handling
    (@handle_http_errors, etc.). This catches anything that escapes.
    """
    @functools.wraps(func)
    async def wrapper(*args, **kwargs):
        tool_name = func.__name__
        try:
            return await func(*args, **kwargs)
        except KeyboardInterrupt:
            # Don't catch keyboard interrupts - allow graceful shutdown
            raise
        except Exception as e:
            # Log the full exception with traceback
            error_str = str(e)
            error_type = type(e).__name__
            
            logger.error(
                f"[GLOBAL EXCEPTION HANDLER] Uncaught exception in tool '{tool_name}': {error_type}: {error_str}",
                exc_info=True
            )
            
            # Log full traceback at debug level for troubleshooting
            tb_str = traceback.format_exc()
            logger.debug(f"Full traceback for {tool_name}:\n{tb_str}")
            
            # If the error message already looks user-friendly (starts with ** or has newlines),
            # just return it as-is
            if error_str.startswith("**") or error_str.startswith("Error:") or error_str.startswith("API error"):
                logger.info(f"[GLOBAL EXCEPTION HANDLER] Returning formatted error from {tool_name}")
                return error_str
            
            # Otherwise, create a formatted error message
            error_msg = (
                f"**Unexpected Error in {tool_name}**\n\n"
                f"An error occurred while executing this tool:\n\n"
                f"```\n{error_type}: {error_str}\n```\n\n"
                f"This error has been logged for investigation. "
                f"If this persists, please check the server logs for more details."
            )
            
            return error_msg
    
    return wrapper


def wrap_server_tool_method(server):
    """
    Track tool registrations, add global exception handling, and filter them post-registration.
    This prevents uncaught exceptions from crashing the MCP server.
    
    The global exception handler is applied as the OUTERMOST layer, ensuring that exceptions
    from any decorator in the stack (including @require_google_service, @handle_http_errors, etc.)
    are caught and don't crash the server.
    """
    original_tool = server.tool
    server._tracked_tools = []
    
    def tracking_tool(*args, **kwargs):
        original_decorator = original_tool(*args, **kwargs)
        
        def wrapper_decorator(func: Callable) -> Callable:
            tool_name = func.__name__
            server._tracked_tools.append(tool_name)
            
            # First, wrap the user's function with global exception handler
            # This goes at the bottom of the decorator stack
            safe_func = global_exception_handler(func)
            
            # Then apply the FastMCP decorator to register it
            # Any additional decorators (@handle_http_errors, @require_google_service)
            # are already applied to `func` before this point
            return original_decorator(safe_func)
        
        return wrapper_decorator
    
    server.tool = tracking_tool

def filter_server_tools(server):
    """Remove disabled tools from the server after registration."""
    enabled_tools = get_enabled_tools()
    if enabled_tools is None:
        return
    
    tools_removed = 0
    
    # Access FastMCP's tool registry via _tool_manager._tools
    if hasattr(server, '_tool_manager'):
        tool_manager = server._tool_manager
        if hasattr(tool_manager, '_tools'):
            tool_registry = tool_manager._tools
            
            tools_to_remove = []
            for tool_name in list(tool_registry.keys()):
                if not is_tool_enabled(tool_name):
                    tools_to_remove.append(tool_name)
            
            for tool_name in tools_to_remove:
                del tool_registry[tool_name]
                tools_removed += 1
    
    if tools_removed > 0:
        logger.info(f"🔧 Tool tier filtering: removed {tools_removed} tools, {len(enabled_tools)} enabled")