"""Diagnostic tools for Solid Edge MCP."""

from typing import Any

from solidedge_mcp.backends.errors import error_result
from solidedge_mcp.managers import diagnose_document, diagnose_feature, doc_manager
from solidedge_mcp.tools._registry import register_tool


def diagnose_api() -> dict[str, Any]:
    """Inspect the COM connection and active document (type, collections, methods)."""
    doc = doc_manager.get_active_document()
    return diagnose_document(doc)


def diagnose_feature_tool(feature_index: int = 0) -> dict[str, Any]:
    """Inspect a Models entry (0-based feature_index): type, properties, methods."""

    try:
        doc = doc_manager.get_active_document()
        model = doc.Models.Item(feature_index + 1)
        return diagnose_feature(model)
    except Exception as e:
        return error_result(e)


def register(mcp: Any) -> None:
    """Register diagnostic tools with the MCP server."""
    tags = {"diagnostics"}
    register_tool(mcp, diagnose_api, tags=tags, read_only=True)
    register_tool(mcp, diagnose_feature_tool, tags=tags, read_only=True)
