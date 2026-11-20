"""Solid Edge MCP Server entry point."""

from __future__ import annotations

from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
from importlib.metadata import PackageNotFoundError, version

from fastmcp import FastMCP

from solidedge_mcp.backends.logging import configure_logging
from solidedge_mcp.prompts import SERVER_INSTRUCTIONS
from solidedge_mcp.tools import register_tools


def _package_version() -> str:
    try:
        return version("solidedge-mcp")
    except PackageNotFoundError:  # running from a source tree without install
        return "0.0.0"


@asynccontextmanager
async def _no_docket() -> AsyncIterator[None]:
    """Stand-in for FastMCP's docket lifespan: start nothing, just yield."""
    yield


def create_server() -> FastMCP:
    """Build a fully registered server instance (used by main() and tests)."""
    configure_logging()
    server = FastMCP(
        "Solid Edge MCP Server",
        instructions=SERVER_INSTRUCTIONS,
        version=_package_version(),
        website_url="https://github.com/tylerwagler/SolidEdge-MCP",
    )
    register_tools(server)
    # FastMCP starts an in-memory docket worker for background tasks. Its
    # fakeredis backend busy-polls, burning a full core while idle on Windows,
    # and this server uses no background tasks, so skip it. Guarded so a
    # FastMCP upgrade that renames the hook does not break startup.
    if hasattr(server, "_docket_lifespan"):
        server._docket_lifespan = _no_docket  # type: ignore[method-assign]
    return server


mcp = create_server()


def main() -> None:
    """Entry point for the MCP server (stdio transport)."""
    mcp.run()


if __name__ == "__main__":
    main()
