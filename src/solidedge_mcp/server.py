"""Solid Edge MCP Server"""

from fastmcp import FastMCP

# Import managers to ensure they are initialized (though tools import them too)
from solidedge_mcp.tools import register_tools

# Create the FastMCP server
mcp = FastMCP("Solid Edge MCP Server")

# Register all tools
register_tools(mcp)


def main() -> None:
    """Entry point for the MCP server."""
    mcp.run()


if __name__ == "__main__":
    main()
