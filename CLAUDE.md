# CLAUDE.md

Guidance for Claude Code when working in this repository.

## What this is

A Solid Edge MCP server: FastMCP 2.x over pywin32 COM automation. Windows only. MIT.
Surface: 117 tools, 52 data resources + 2 guide resources, 4 prompts.

## Commands

```bash
uv sync --all-extras          # install (Python 3.11+, Windows)
uv run solidedge-mcp          # run the server (stdio)
uv run pytest                 # unit tests; integration tests are deselected by default
uv run pytest -m integration  # needs a running, licensed Solid Edge
uv run pytest tests/unit/test_features_extrude.py::TestCreateExtrude::test_success
uv run pytest --cov           # coverage report
uv run ruff check . && uv run ruff format .
uv run mypy src/

# COM conformance against the scraped type libraries
uv run python scripts/audit_com_signatures.py --by-file
uv run python scripts/audit_com_signatures.py --filter backends/features/_holes.py
uv run python scripts/scrape_typelibs.py    # regenerate the dump (needs Solid Edge)
```

CI (`.github/workflows/ci.yml`, windows-latest) runs ruff check, ruff format --check, mypy, pytest. Keep all four green.

## Layout

```
src/solidedge_mcp/
├── server.py          create_server(): FastMCP(instructions=...), register_tools()
├── managers.py        global manager instances (connection, doc_manager, sketch_manager, ...)
├── prompts/           SERVER_INSTRUCTIONS, WORKFLOWS_GUIDE, CONVENTIONS_GUIDE, register_prompts()
├── tools/             MCP surface. One module per area; each has register(mcp)
│   ├── _registry.py   register_tool()/register_resource(): COM-thread wrap + annotations + tags
│   ├── features/      feature tools split by family (_extrude.py, _cutout.py, ...)
│   ├── resources.py   52 solidedge:// resources (read-only JSON)
│   └── guide.py       solidedge://guide/workflows, solidedge://guide/conventions
└── backends/          COM automation
    ├── connection.py  SolidEdgeConnection: attach/start, liveness probe, reconnect
    ├── com_thread.py  ComThread: single STA worker; on_com_thread() decorator
    ├── errors.py      error_result(e): decodes com_error; tracebacks only if SOLIDEDGE_MCP_DEBUG
    ├── logging.py     stderr logger; level from SOLIDEDGE_MCP_LOG_LEVEL
    ├── documents.py   DocumentManager: tracks doc switches, clears sketch state
    ├── sketching.py   SketchManager: active_profile, accumulated_profiles
    ├── features/      FeatureManager mixins (_base.py has @verifies_geometry)
    ├── assembly/      AssemblyManager mixins
    ├── query/         QueryManager mixins
    ├── export/        ExportManager + ViewModel mixins
    ├── constants.py   Solid Edge enum values (cite the source enum in a comment)
    └── validation.py  validate_numerics(), validate_path()
tests/unit/            mocked-COM tests (conftest-free; fixtures per file)
tests/integration/     @pytest.mark.integration, real Solid Edge
scripts/               scrape_typelibs.py (load-bearing); manual/ = hand-run COM experiments
reference/             typelib_summary.md (committed), typelib_dump.json (gitignored, regenerate)
```

## How a tool is built

1. **Backend method** on the right manager mixin. Wrap COM in `try/except Exception as e: return error_result(e)`. Return `dict[str, Any]` with a `status` key on success. Never return a bare traceback.
2. **Tool function** in `tools/<area>.py` (or `tools/features/_<family>.py`). Composite tools dispatch on a `Literal[...]` discriminator (`method`/`type`/`action`) with `match/case`; keep `case _:` returning `{"error": "Unknown ..."}`. Docstrings are LLM-facing schema text: terse, state units and index base, say which params apply to which method.
3. **Register** in that module's `register(mcp)` via `register_tool(mcp, fn, tags={...}, read_only=..., destructive=...)`. Never call `mcp.tool()` directly; the registry wraps the call onto the COM thread and sets annotations.
4. **Tests** in `tests/unit/test_tools_<area>.py` (dispatch, every case label) and `tests/unit/test_<backend>.py` (COM call arguments with `assert_called_once_with`, not just "returns status").
5. Update `reference/TYPELIB_IMPLEMENTATION_MAP.md` if you add COM coverage.

Count tools with `grep -rc "register_tool(" src/solidedge_mcp/tools | awk -F: '{s+=$2} END {print s}'`.

## Solid Edge / COM rules

- **Units**: meters and degrees at the tool boundary. Convert to radians (`math.radians`) before COM.
- **Plane indices are 1-based**: 1=Top/XY, 2=Right/YZ, 3=Front/XZ (`RefPlanes.Item(n)`). Face/edge/feature/component indices are 0-based in tools and converted at the COM boundary.
- **Sketch then feature**: `create_sketch → draw_* → close_sketch → create_<feature>`. `close_sketch` calls `Profile.End(igProfileClosed)` and queues the profile in `sketch_manager.accumulated_profiles`; the feature consumes it.
- **Threading**: all COM runs on `com_thread`. Tool functions are plain sync `def`. Backend code may call other backend code freely (nested calls run inline).
- **Never `hasattr()` a COM proxy to test capability**; check `doc.Type` against `DocumentTypeConstants` or `try/except` the actual call.
- **Never compare COM proxies with `==`**; compare `FullName`/`Name`.
- **Collections are 1-based** in COM (`Item(1)`).
- **Pass SAFEARRAYs explicitly** where the type library says so: `VARIANT(VT_ARRAY | VT_DISPATCH, [...])`; see `features/_base.py`.
- **Cutouts** use collection-level APIs (`model.ExtrudedCutouts.AddFiniteMulti`), not `Models.AddExtrudedCutout`.
- **Known unsupported via COM (SE 2025/2026)**: `AssemblyFeaturesPatterns.Add`, `AssemblyFeaturesMirrors.Add` (E_ACCESSDENIED); shell/thin-wall (needs interactive face pick); multiple disjoint profiles in one cutout sketch.
- **Front plane quirk**: COM "Normal" on the Front plane points to world −Y. Cutout tools do not auto-swap; prefer `direction="Symmetric"` there.
- **Exports** use `SaveCopyAs`, never `SaveAs` (which repoints the live document).

## Type library reference

`reference/typelib_dump.json` is the source of truth for COM signatures and enum values but is **gitignored and absent from a fresh clone**. Regenerate it with `uv run python scripts/scrape_typelibs.py` (requires Solid Edge installed). Until then use `reference/typelib_summary.md` (truncated to the first values of each enum) and `reference/TYPELIB_IMPLEMENTATION_MAP.md`.

Rules: never guess a constant or signature. Look it up, copy the exact value into `constants.py` with a comment naming the enum, and prefer collection-level `Add*` methods.

Three checks enforce this, and all three skip when the dump is absent:

| Check | What it catches |
|---|---|
| `tests/unit/test_constants_typelib.py` | A constant whose value disagrees with its enum. Classes that are our own vocabulary go in `LOCAL_GROUPINGS` with a note. |
| `tests/unit/test_com_members.py` | A COM member name that exists in no type library. A ratchet: new names fail, and fixing one fails until you delete it from `UNVERIFIED`. |
| `scripts/audit_com_signatures.py` | A call with the wrong number of arguments. Run `--filter <path>` to see the full parameter list for each finding, `--by-file` for counts. |

The signature audit resolves the receiver, so `cutouts = model.ExtrudedCutouts` followed by `cutouts.AddFiniteMulti(...)` is checked against `ExtrudedCutouts` specifically rather than against every interface with that method name.

### When a COM call cannot be formed

Some methods require an object this server cannot obtain: a `KeyPoint`, a specific `Face`, a tangent face, a user selection. Do not pass `None` or guess. Return an error without touching COM, and keep the method signature so tool dispatch still works:

```python
return {
    "error": (
        "Normal cutout to a keypoint needs a KeyPoint or tangent face object, "
        "which this server cannot select. Use create_normal_cutout(distance) "
        "or the Solid Edge UI."
    ),
    "unsupported": True,
}
```

The same applies to APIs Solid Edge blocks for automation, such as `AssemblyFeaturesPatterns.Add` and `AssemblyFeaturesMirrors.Add` (both `E_ACCESSDENIED` on 2025/2026). An honest error beats a call that always fails.

## Testing notes

- Unit tests mock COM with `MagicMock`; a wrong COM member name still passes unless a test asserts it. Assert call arguments (`assert_called_once_with`) for any new COM call.
- `@verifies_geometry` is bypassed when `doc` is a `unittest.mock` object; `tests/unit/test_features_base.py` tests it with hand-written fakes.
- `tests/unit/test_server.py` builds the real server and checks every tool has annotations/tags and runs on the COM thread.
- Integration tests need Solid Edge and are deselected by default (`addopts = -m 'not integration'`).
