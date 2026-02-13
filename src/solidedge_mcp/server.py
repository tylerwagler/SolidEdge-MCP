"""Solid Edge MCP Server"""

from fastmcp import FastMCP
from solidedge_mcp.tools import register_tools
# Import managers to ensure they are initialized (though tools import them too)
from solidedge_mcp import managers

# Create the FastMCP server
mcp = FastMCP("Solid Edge MCP Server")

# Register all tools
register_tools(mcp)

if __name__ == "__main__":
    # Verification check (optional, but requested by user)
    # Access _tool_manager._tools if available or assume correct registration
    try:
        if hasattr(mcp, '_tool_manager'):
            count = len(mcp._tool_manager._tools)
            print(f"Registered {count} tools")
    except Exception:
        pass

    mcp.run()
