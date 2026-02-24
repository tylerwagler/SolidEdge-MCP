from typing import Any

from . import (
    assembly,
    connection,
    diagnostics,
    documents,
    export,
    features,
    query,
    resources,
    sketching,
)


def register_tools(mcp: Any) -> None:
    """Register all tools and resources with the MCP server instance."""
    # Resources (read-only data endpoints)
    resources.register(mcp)
    # Tools (actions that modify state)
    connection.register(mcp)
    documents.register(mcp)
    sketching.register(mcp)
    features.register(mcp)
    assembly.register(mcp)
    query.register(mcp)
    export.register(mcp)
    diagnostics.register(mcp)
