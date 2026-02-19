"""Structured logging for Solid Edge MCP server.

Provides consistent logging throughout the codebase with appropriate
log levels and context for debugging COM automation issues.
"""

import logging
import sys
from typing import Any

# Default log level can be overridden via environment variable
LOG_LEVEL = logging.INFO

# Create a module-level logger
logger = logging.getLogger("solidedge_mcp")

# Create handler with formatter if not already configured
if not logger.handlers:
    handler = logging.StreamHandler(sys.stdout)
    handler.setLevel(LOG_LEVEL)
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)

logger.setLevel(LOG_LEVEL)


def get_logger(name: str) -> logging.Logger:
    """Get a logger instance with the given name.

    Args:
        name: The name for the logger (typically __name__)

    Returns:
        A configured logging.Logger instance
    """
    return logging.getLogger(f"solidedge_mcp.{name}")


def log_com_call(
    logger_instance: logging.Logger,
    method_name: str,
    args: list,
    result: Any = None,
    error: str = None,
) -> None:
    """Log a COM method call and its result.

    Args:
        logger_instance: The logger to use
        method_name: Name of the COM method called
        args: List of arguments passed to the method
        result: The result returned by the method (if successful)
        error: Error message if the call failed
    """
    if error:
        logger_instance.debug(f"COM call {method_name} failed: {error}")
    else:
        logger_instance.debug(f"COM call {method_name}(args={args}) -> {result}")


def log_sketch_operation(
    logger_instance: logging.Logger,
    operation: str,
    profile_state: dict,
) -> None:
    """Log sketch operations for debugging.

    Args:
        logger_instance: The logger to use
        operation: Name of the sketch operation
        profile_state: Current state of the profile (has_active_profile, element counts, etc.)
    """
    logger_instance.info(f"Sketch operation: {operation}, state: {profile_state}")


def log_document_event(logger_instance: logging.Logger, event: str, doc_info: dict) -> None:
    """Log document lifecycle events.

    Args:
        logger_instance: The logger to use
        event: Type of event (created, opened, saved, closed, activated)
        doc_info: Document metadata (name, type, path)
    """
    logger_instance.info(f"Document event: {event}, info: {doc_info}")
