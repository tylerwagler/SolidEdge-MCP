# Claude Code CLI Configuration

## Setup Instructions

To use the SolidEdge MCP server with Claude Code CLI, add it to your MCP configuration file.

### Configuration File Location

**Windows**: `C:\Users\<username>\.claude\mcp_config.json`

### Configuration

Add this to your `mcp_config.json`:

```json
{
  "mcpServers": {
    "solidedge": {
      "command": "uv",
      "args": [
        "--directory",
        "C:/Users/tyler/Dev/repos/SolidEdge_MCP",
        "run",
        "solidedge-mcp"
      ],
      "env": {}
    }
  }
}
```

### Quick Setup Command

Run this command to add the configuration:

```bash
mkdir -p ~/.claude
cat > ~/.claude/mcp_config.json << 'EOF'
{
  "mcpServers": {
    "solidedge": {
      "command": "uv",
      "args": [
        "--directory",
        "C:/Users/tyler/Dev/repos/SolidEdge_MCP",
        "run",
        "solidedge-mcp"
      ],
      "env": {}
    }
  }
}
EOF
```

### Verify Setup

1. Restart Claude Code CLI if it's running
2. The SolidEdge MCP server will auto-start when needed
3. Ask Claude Code: "Connect to Solid Edge" to test

### Available Tools

Once configured, you'll have access to 28+ tools for:
- Connecting to Solid Edge
- Creating and managing documents
- Sketching 2D geometry
- Creating 3D features
- Querying mass properties
- Exporting to various formats
- Assembly operations
- API diagnostics

### Troubleshooting

If the server doesn't load:
1. Check that `uv` is in your PATH
2. Verify the directory path is correct
3. Test manually: `uv --directory C:/Users/tyler/Dev/repos/SolidEdge_MCP run solidedge-mcp`
4. Check Claude Code logs for errors
