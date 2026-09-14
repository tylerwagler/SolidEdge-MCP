# Solid Edge MCP Server

AI-assisted CAD through the [Model Context Protocol](https://modelcontextprotocol.io). Create, analyze, modify, and export Solid Edge models from Claude or any MCP client.

**117 tools** · **54 resources** · **4 prompts** · Windows only (COM automation) · MIT

## What it does

- **Connect** to a running Solid Edge, or start one
- **Documents**: create, open, save, close parts, sheet metal, assemblies, drafts
- **Sketch** 2D profiles: lines, circles, arcs, rectangles, polygons, splines, constraints
- **Features**: extrude, revolve, sweep, loft, helix, cutouts, holes, rounds, chamfers, patterns, ref planes, surfaces, sheet metal
- **Assemblies**: place components, relations, transforms, BOM, interference
- **Drafts**: drawing views, annotations, dimensions, parts lists
- **Query**: geometry, mass properties, feature tree, variables, materials
- **Export**: STEP, STL, IGES, JT, Parasolid, PDF, DXF, images

## Requirements

- Windows 10/11
- Solid Edge installed, licensed, and launched at least once (registers the COM server). Developed against Solid Edge 2025/2026.
- Python 3.11+ and [uv](https://docs.astral.sh/uv/)

## Install

```bash
git clone https://github.com/tylerwagler/SolidEdge-MCP
cd SolidEdge-MCP
uv sync --all-extras
```

## Configure your client

**Claude Code**: copy [`examples/claude_code.mcp.json`](examples/claude_code.mcp.json) to `.mcp.json` in your project (or run `claude mcp add solidedge -- uv --directory C:/path/to/SolidEdge-MCP run solidedge-mcp`).

**Claude Desktop**: merge [`examples/claude_desktop_config.json`](examples/claude_desktop_config.json) into your `claude_desktop_config.json`.

Both examples use:

```json
{
  "mcpServers": {
    "solidedge": {
      "command": "uv",
      "args": ["--directory", "C:/path/to/SolidEdge-MCP", "run", "solidedge-mcp"]
    }
  }
}
```

Restart the client, then ask it to connect to Solid Edge.

### Environment variables

| Variable | Effect |
|---|---|
| `SOLIDEDGE_MCP_LOG_LEVEL` | `DEBUG`, `INFO`, `WARNING` (default), `ERROR`. Logs go to stderr. |
| `SOLIDEDGE_MCP_DEBUG` | `1` forces DEBUG logging and includes Python tracebacks in error results. |

### Troubleshooting

- **Server does not start**: run `uv --directory C:/path/to/SolidEdge-MCP run solidedge-mcp` in a terminal and read stderr.
- **"Not connected"**: call `manage_connection(action="connect")` first. If Solid Edge was restarted, call it again; the server drops the dead connection and reattaches.
- **COM error 0x80070005 (E_ACCESSDENIED)** on assembly pattern/mirror: those APIs are blocked for automation in Solid Edge 2025/2026. Pattern at part level instead.
- **Feature "created" but nothing visible**: the server verifies geometry after every material-changing feature and reports an error when nothing changed. Check the sketch is closed and on the plane you expect.

## Conventions the LLM sees

The server publishes these in its MCP `instructions` and in two resources, `solidedge://guide/workflows` and `solidedge://guide/conventions`:

- Lengths in **meters**, angles in **degrees**
- Reference planes are **1-based**: 1=Top (XY), 2=Right (YZ), 3=Front (XZ)
- Faces, edges, features, components are **0-based**
- Every result is a dict; `{"error": ...}` means failure. COM failures carry `hresult`.
- Solids need a closed sketch: `create_sketch → draw_* → close_sketch → create_extrude`

Tools carry MCP annotations (`readOnlyHint`, `destructiveHint`) and tags (`part`, `assembly`, `draft`, `sheet_metal`, `query`, `export`, `app`, `document`, `sketch`, `diagnostics`) so clients can filter the manifest.

## Architecture

```
src/solidedge_mcp/
├── server.py          # FastMCP instance: instructions, registration
├── managers.py        # Global backend manager instances
├── prompts/           # Server instructions, guides, prompt templates
├── tools/             # MCP surface: one module per area, register() each
│   ├── _registry.py   #   register_tool/register_resource: COM thread + annotations
│   ├── features/      #   58 feature tools split by family
│   ├── resources.py   #   52 read-only solidedge:// resources
│   └── guide.py       #   solidedge://guide/* resources
└── backends/          # pywin32 COM automation
    ├── connection.py  # attach/start Solid Edge, liveness + reconnect
    ├── com_thread.py  # single STA worker thread all COM calls run on
    ├── errors.py      # com_error -> readable error dicts
    ├── documents.py   # create/open/save; tracks document switches
    ├── sketching.py   # 2D profiles
    ├── features/      # 3D features (mixin package)
    ├── assembly/      # assembly operations (mixin package)
    ├── query/         # interrogation (mixin package)
    ├── export/        # export, drafting, views (mixin package)
    └── constants.py   # Solid Edge enum values
```

Every tool and resource is marshalled onto one COM worker thread (`CoInitializeEx` STA), because FastMCP otherwise runs synchronous tools on a thread pool and COM proxies are apartment-bound.

## Development

```bash
uv sync --all-extras
uv run pytest              # unit tests (mocked COM); integration tests deselected
uv run pytest -m integration   # needs a running, licensed Solid Edge
uv run pytest --cov        # coverage
uv run ruff check . && uv run ruff format --check .
uv run mypy src/
```

CI runs lint, format, type check, and unit tests on `windows-latest` for every push and pull request.

### COM conformance

Unit tests mock COM with objects that answer to any attribute, so a misspelled member or a wrong argument count passes them and only fails against real Solid Edge. Five checks close that gap using the scraped type libraries:

```bash
uv run python scripts/audit_com_signatures.py --by-file
uv run python scripts/audit_com_signatures.py --filter backends/features/_holes.py
uv run python scripts/audit_com_receivers.py
uv run python scripts/audit_com_writes.py
uv run python scripts/audit_com_hasattr.py
uv run python scripts/audit_reported_writes.py
uv run python scripts/audit_dead_params.py
```

`tests/unit/test_constants_typelib.py` verifies every COM enum value in `constants.py`. `tests/unit/test_com_members.py` fails on any COM member name absent from the type libraries. `tests/unit/test_com_receivers.py` goes further: it infers what each receiver is by following declared types and fails when a member is read off an interface that does not have it, which is invisible to a name check when the name is real somewhere else. `tests/unit/test_com_writes.py` covers the other direction: a write to a member that is a method, or to a property the type library marks read-only. Solid Edge answers "Property 'Item.X' can not be set." and a try/except around it turns the dead feature into a reported success. All skip when the dump is missing, so a fresh clone and CI stay green.

Four checks need no type library. `tests/unit/test_reported_writes.py` fails when a COM write is swallowed by a suppress and the value is then reported as applied -- the shape behind most of the silent no-ops found in this codebase, from text heights that never changed to a variable rename that reported "not found". `tests/unit/test_com_hasattr.py` fails on `hasattr` used to test what a COM proxy can do: the probe reads False both for a member that is absent and for one whose getter raised, which is how every layer call came to refuse draft documents, whose layers live on the sheet rather than the document. `tests/unit/test_manager_mro.py` fails when two mixins define the same method on a manager, which leaves one of them unreachable.

`tests/unit/test_dead_params.py` fails when a parameter is declared and never read, in a tool or a backend method. Nothing downstream can catch that: Solid Edge is never told the value, so it has nothing to reject, and the call returns a cheerful success. Three shipped that way -- a component position that always placed at the origin, a table position Solid Edge chose for itself, and a revolve axis setting nothing consulted. A parameter that genuinely has nowhere to go says so with `del <param>`.

`scripts/scrape_typelibs.py` regenerates `reference/typelib_dump.json` (gitignored, ~20 MB) from the installed Solid Edge type libraries. `reference/typelib_summary.md` is the committed digest. `scripts/manual/` holds hand-run COM experiments; they are not tests.

## License

MIT
