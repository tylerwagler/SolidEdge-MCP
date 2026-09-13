"""Tool, resource, and prompt registration.

Each module exposes ``register(mcp)``. Tools and resources are registered
through :mod:`solidedge_mcp.tools._registry` so they run on the COM worker
thread and carry MCP annotations and tags.
"""

from typing import Any

from solidedge_mcp.prompts import register_prompts

from . import (
    assembly,
    connection,
    diagnostics,
    documents,
    export,
    features,
    guide,
    query,
    resources,
    sketching,
)


def register_tools(mcp: Any) -> None:
    """Register all tools, resources, and prompts with the MCP server instance."""
    # Guidance (static text)
    guide.register(mcp)
    register_prompts(mcp)
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
