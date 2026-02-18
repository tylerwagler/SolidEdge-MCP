from . import (
    assembly,
    connection,
    diagnostics,
    documents,
    export,
    features,
    query,
    resources,
    sheet_metal,
    sketching,
)


def register_tools(mcp):
    """Register all tools and resources with the MCP server instance."""
    # Resources (read-only data endpoints)
    resources.register(mcp)
    # Tools (actions that modify state)
    connection.register(mcp)
    documents.register(mcp)
    sketching.register(mcp)
    features.register(mcp)
    assembly.register(mcp)
    sheet_metal.register(mcp)
    query.register(mcp)
    export.register(mcp)
    diagnostics.register(mcp)
