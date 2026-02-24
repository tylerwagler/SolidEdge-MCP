"""Diagnostic tools for Solid Edge MCP."""

from typing import Any

from solidedge_mcp.managers import diagnose_document, diagnose_feature, doc_manager


def diagnose_api() -> dict[str, Any]:
    """Run diagnostic checks on the Solid Edge API connection and active document."""
    doc = doc_manager.get_active_document()
    return diagnose_document(doc)


def diagnose_feature_tool(feature_index: int = 0) -> dict[str, Any]:
    """Inspect a feature/model object - shows type, properties, available methods.

    Args:
        feature_index: 0-based index into the Models collection (default: first model)
    """
    import traceback

    try:
        doc = doc_manager.get_active_document()
        model = doc.Models.Item(feature_index + 1)
        return diagnose_feature(model)
    except Exception as e:
        return {"error": str(e), "traceback": traceback.format_exc()}


def register(mcp: Any) -> None:
    """Register diagnostic tools with the MCP server."""
    mcp.tool()(diagnose_api)
    mcp.tool()(diagnose_feature_tool)
