"""Diagnostic tools for Solid Edge MCP."""

from solidedge_mcp.managers import diagnose_document, diagnose_feature, doc_manager

def diagnose_api() -> dict:
    """Run diagnostic checks on the Solid Edge API connection and active document."""
    # Assuming diagnose_document takes the active document object
    doc = doc_manager.get_active_document()
    return diagnose_document(doc)

def register(mcp):
    """Register diagnostic tools with the MCP server."""
    mcp.tool()(diagnose_api)
