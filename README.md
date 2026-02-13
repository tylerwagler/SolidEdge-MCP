# Solid Edge MCP Server

AI-assisted CAD design through the [Model Context Protocol](https://modelcontextprotocol.io). Create, analyze, modify, and export Solid Edge models — all from your AI assistant.

**252 MCP tools** | **404 unit tests** | Windows-only (COM automation)

## What It Does

This MCP server gives AI assistants (Claude, etc.) full access to Solid Edge CAD workflows:

- **Connect to Solid Edge** application via COM automation
- **Create and manage** parts, assemblies, sheet metal, and drafts
- **Sketch 2D geometry** - lines, circles, arcs, rectangles, polygons, splines, constraints
- **Create 3D features** - extrude, revolve, sweep, loft, helix, cutouts, rounds, chamfers, holes
- **Query and analyze** models - dimensions, mass properties, feature trees, materials
- **Assemblies** - place components, move/rotate, BOM, interference checks
- **Draft/Drawing** - create views, annotations, parts lists
- **Export** models to STEP, STL, IGES, PDF, DXF, Parasolid, JT
- **View control** - set viewpoints, zoom, camera, display modes

## Quick Start

### Install

```bash
# Requires Python 3.11+ and Windows (Solid Edge is Windows-only)
uv sync --all-extras
```

### Configure Claude Code CLI

Add to your MCP configuration file at `~/.claude/mcp_config.json`:

```json
{
  "mcpServers": {
    "solidedge": {
      "command": "uv",
      "args": [
        "--directory",
        "C:/path/to/SolidEdge_MCP",
        "run",
        "solidedge-mcp"
      ],
      "env": {}
    }
  }
}
```

### Configure Claude Desktop

Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "solidedge-mcp": {
      "command": "uv",
      "args": ["--directory", "C:/path/to/SolidEdge_MCP", "run", "solidedge-mcp"]
    }
  }
}
```

### Run Standalone

```bash
uv run solidedge-mcp
```

### Verify Setup

1. Restart your AI client if it's running
2. The SolidEdge MCP server will auto-start when needed
3. Ask: "Connect to Solid Edge" to test

### Troubleshooting

If the server doesn't load:
1. Check that `uv` is in your PATH
2. Verify the directory path is correct
3. Test manually: `uv --directory C:/path/to/SolidEdge_MCP run solidedge-mcp`
4. Check your client's logs for errors

## Architecture

### COM Automation Backend

The server communicates with Solid Edge through Windows COM automation (pywin32):

- **Connection management** - Connect to running instance or start new one
- **Document handling** - Create, open, save, close parts and assemblies
- **Sketching** - 2D profile creation on reference planes
- **Features** - 3D modeling operations (extrude, revolve, etc.)
- **Assembly** - Component placement, constraints, patterns
- **Query** - Extract geometry, dimensions, properties, materials
- **Export** - Convert models to standard CAD formats

### Package Layout

```
src/solidedge_mcp/
├── server.py           # FastMCP server entry point (252 @mcp.tool() wrappers)
├── backends/           # COM automation implementations
│   ├── connection.py   # Application connection management
│   ├── documents.py    # Document operations
│   ├── sketching.py    # 2D sketch creation
│   ├── features.py     # 3D feature operations
│   ├── assembly.py     # Assembly operations
│   ├── query.py        # Model interrogation
│   ├── export.py       # Export and view operations
│   └── constants.py    # Solid Edge API constants
```

## Requirements

- **Python 3.11+**
- **Windows** (Solid Edge is Windows-only)
- **Solid Edge** installed and licensed
- **pywin32** for COM automation

## Development

```bash
# Install with dev dependencies
uv sync --all-extras

# Run tests
uv run pytest

# Lint and format
uv run ruff check .
uv run ruff format .

# Type check
uv run mypy src/
```

## License

MIT
