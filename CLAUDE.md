# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

A Solid Edge MCP (Model Context Protocol) server for AI-assisted CAD design. Windows-only, built on FastMCP and pywin32 COM automation. Licensed MIT.

The goal is to provide AI assistants with full access to Solid Edge CAD workflows: **connect → create → sketch → feature → query → export** with session management and undo/rollback support.

## Commands

```bash
# Install all dependencies (including dev)
uv sync --all-extras

# Run the MCP server (stdio transport)
uv run solidedge-mcp

# Run tests
uv run pytest
uv run pytest tests/unit/test_foo.py::test_bar  # single test

# Lint and format
uv run ruff check .
uv run ruff format .

# Type check
uv run mypy src/
```

## Architecture

### COM Automation Backend

Unlike KiCad (which uses file parsing), Solid Edge automation requires Windows COM through pywin32. The server communicates with a running Solid Edge instance via COM interfaces:

- **Connection layer** (`backends/connection.py`): Manages GetActiveObject/Dispatch, early/late binding
- **Document layer** (`backends/documents.py`): Create/open/save parts, assemblies, drafts
- **Sketching layer** (`backends/sketching.py`): 2D profile creation (lines, circles, arcs, rectangles, polygons)
- **Feature layer** (`backends/features.py`): 3D operations (extrude, revolve, sweep, loft, holes, fillets)
- **Assembly layer** (`backends/assembly.py`): Component placement, constraints, patterns
- **Query layer** (`backends/query.py`): Extract geometry, mass properties, feature trees
- **Export layer** (`backends/export.py`): Convert to STEP, STL, IGES, PDF, DXF

### Package Layout

```
src/solidedge_mcp/
├── server.py              # FastMCP server entry point
├── backends/              # COM automation implementations
│   ├── connection.py      # Application connection (GetActiveObject/Dispatch)
│   ├── documents.py       # Document create/open/save/close
│   ├── sketching.py       # 2D sketch profiles
│   ├── features.py        # 3D feature operations
│   ├── assembly.py        # Assembly operations
│   ├── query.py           # Model interrogation
│   ├── export.py          # Export to standard formats
│   └── constants.py       # Solid Edge API constants
├── tools/                 # MCP tool wrappers (pending)
├── resources/             # MCP Resources (read-only state) (pending)
├── prompts/               # MCP Prompt templates (pending)
└── session/               # Session/undo management (pending)
```

### Current State

**✅ FULLY IMPLEMENTED**: All 299 MCP tools are registered and operational!

- **Backend layer**: Complete COM automation using pywin32 (connection, documents, sketching, features, assembly, query, export, diagnostics)
- **MCP tools**: All 241 tools registered in `server.py` using `@mcp.tool()` decorator
- **Tool categories**: Connection (6), Documents (13), Sketching (22), Primitives (8), Extrusions (4), Revolves (5), Cutouts (9), Ref Planes (5), Rounds/Chamfers/Holes (9), Mirror (1), Loft (2), Sweep (2), Helix (4), Surfaces (3), Sheet Metal (8), Body Operations (9), Simplification (4), View (7), Variables (5), Custom Properties (3), Body Topology (3), Performance (2), Query/Analysis (20), Feature Management (5), Draft/Drawing (10), Export (10), Assembly (22), Part Features (10), Select Set (3), Diagnostics (2)
- **Coverage**: 97% of core Solid Edge Part COM API feature collections implemented
- **Test suite**: 504 unit tests across 5 test files

**Pending**: Resource providers (read-only state), prompt templates, session management/undo

### Three-Pillar MCP Design

Following the MCP spec, the server exposes:

- **Tools** ✅ (299 implemented): Actions that create/modify models (connect, create_sketch, extrude, cutout, round, chamfer, hole, mirror, place_component, export)
- **Resources** ⏳ (pending): Read-only model data (feature list, component tree, mass properties, document info)
- **Prompts** ⏳ (pending): Conversation templates (design review, manufacturability check, modeling guidance)

### Tool Categories (299 total)

See `reference/TYPELIB_IMPLEMENTATION_MAP.md` for the complete list. High-level categories:

1. **Connection (3)**: `connect_to_solidedge`, `get_application_info`, `quit_application`
2. **Documents (11)**: Create (part, assembly, sheet metal), open, save, close, list, activate, undo, redo
3. **Sketching (11)**: Lines, circles, arcs, rectangles, polygons, ellipses, splines, constraints, sketch on any plane
4. **Primitives (5)**: Box (3 variants), cylinder, sphere
5. **Extrusions (3)**: Finite, infinite, thin-wall
6. **Revolves (5)**: Basic, finite, sync variants, thin-wall
7. **Cutouts (5)**: Extruded (finite + through-all), revolved, normal, lofted
8. **Reference Planes (2)**: Offset parallel plane, list planes
9. **Rounds/Chamfers/Holes (3)**: Round (fillet), chamfer, hole
10. **Mirror (1)**: Mirror copy across plane
11. **Loft (2)**: Basic and thin-wall
12. **Sweep (2)**: Basic and thin-wall
13. **Helix/Spiral (4)**: Various helix creation methods
14. **Sheet Metal (8)**: Base flange/tab, lofted flange, web network
15. **Body Operations (7)**: Add body, thicken, mesh, tag-based, construction
16. **Simplification (4)**: Auto-simplify, enclosure, duplicate
17. **View/Display (8)**: Set view orientation, zoom, display mode, rotate/pan/zoom camera, refresh
18. **Variables (3)**: Get all, get by name, set value
19. **Custom Properties (3)**: Get all, set/create, delete
20. **Body Topology (3)**: Body faces, body edges, face info
21. **Performance (2)**: Set performance mode, recompute
22. **Query/Analysis (12)**: Mass properties, bounding box, features, measurements, facets, solid bodies, modeling mode, select set
23. **Feature Management (2)**: Suppress, unsuppress features
24. **Draft/Drawing (2)**: Add sheet, assembly drawing view
25. **Export (10)**: STEP, STL, IGES, PDF, DXF, flat DXF, Parasolid, JT, screenshot, drawing
26. **Assembly (16)**: Place, list, constraints, patterns, suppress, BOM, interference, bbox, relations, doc tree
27. **Select Set (2)**: Get selection, clear selection
28. **Diagnostics (2)**: API discovery tools

### Tool Registration Pattern

All backend operations are wrapped as MCP tools using the `@mcp.tool()` decorator. Example:

```python
@mcp.tool()
def connect_to_solidedge(start_if_needed: bool = True) -> dict:
    """Connect to Solid Edge application (start if needed)"""
    return connection_manager.connect(start_if_needed)
```

**Key patterns:**
- Tools call backend manager methods (e.g., `connection_manager`, `feature_manager`, `sketch_manager`)
- All tools return `Dict[str, Any]` with consistent structure: `{"status": "...", ...}` or `{"error": "...", "traceback": "..."}`
- Backend managers are initialized globally at module level in `server.py`
- Type hints on all parameters for better IDE support and validation
- Docstrings describe purpose, parameters, and return values

### Manager Pattern

The codebase uses a manager pattern to organize backend operations:

```python
# Global manager instances (initialized in server.py)
connection_manager = ConnectionManager()
doc_manager = DocumentManager(connection_manager)
sketch_manager = SketchManager(doc_manager)
feature_manager = FeatureManager(doc_manager, sketch_manager)
assembly_manager = AssemblyManager(doc_manager)
query_manager = QueryManager(doc_manager)
export_manager = ExportManager(doc_manager)
view_manager = ViewModel(doc_manager)
```

Each manager encapsulates related COM operations and maintains necessary state (e.g., `sketch_manager` tracks the active sketch).

## Solid Edge-Specific Notes

- **Windows-only**: Solid Edge COM automation requires Windows. pywin32 does not work on Linux/macOS.
- **COM binding**: Use `gencache.EnsureDispatch()` for early binding (type hints, IntelliSense) or `Dispatch()` for late binding (more compatible but slower). We use `Dispatch()` for broader compatibility.
- **Active document pattern**: Most operations require an active document. The DocumentManager tracks the active document via COM.
- **Sketch-then-feature workflow**: 3D features (extrude, revolve) require a closed 2D sketch profile. The typical flow is:
  ```
  create_sketch() → draw_line/circle/etc() → close_sketch() → create_extrude()
  ```
  The `SketchManager` maintains `self.active_profile` to track the current sketch.
- **COM exception handling**: COM operations can raise `pywintypes.com_error`. Always wrap in try/except with traceback for debugging:
  ```python
  try:
      # COM operation
  except Exception as e:
      return {"error": str(e), "traceback": traceback.format_exc()}
  ```
- **Reference planes**: Solid Edge has 3 default planes (Top/XZ, Front/XY, Right/YZ). Sketches are created on these planes using `RefPlanes.Item(index)`.
- **Units**: Solid Edge internal units are **meters**. Convert mm to meters by dividing by 1000. All tool parameters use meters.
- **Feature tree**: Features are stored in `Document.Models` collection (1-indexed in COM). Each feature has properties like Name, Type, Status.
- **Cutout operations**: Available via collection-level APIs (`model.ExtrudedCutouts.AddFiniteMulti`, `model.RevolvedCutouts.AddFiniteMulti`), NOT via `models.AddExtrudedCutout` (which doesn't exist).
- **Collections are 1-indexed**: COM collections use 1-based indexing (`collection.Item(1)` is first item), but our tools use 0-based indexing for Python consistency.
- **Profile validation**: After drawing geometry, profiles need to be validated/closed before using them for features. The `close_sketch()` tool calls `profile.End(0)`.
- **Angle units**: Most angle parameters in the API expect **radians**, so convert degrees to radians using `math.radians(angle)`.

## Type Library Reference (MANDATORY)

**All MCP server development MUST use the scraped type library data as the source of truth.**

We have a complete dump of every Solid Edge COM type library (41 .tlb files, 2,240 interfaces,
21,237 methods, 14,575 enum values) in structured JSON. This eliminates guessing at constant
values, method signatures, and parameter types.

### Reference Files

| File | Purpose |
|------|---------|
| `reference/typelib_dump.json` | **Primary reference** - Full structured dump of all type libraries. Search this for exact method signatures, parameter names/types, enum values, and interface hierarchies. |
| `reference/typelib_summary.md` | Human-readable overview of all 40 type libraries with interface/enum counts. |
| `reference/TYPELIB_IMPLEMENTATION_MAP.md` | Maps every COM API surface against current MCP tool coverage. Identifies gaps and prioritizes what to implement. Check this before starting any new tool work. |
| `scripts/scrape_typelibs.py` | Scraper script to regenerate the dump (run if SE version changes). |

### Rules for COM API Development

1. **NEVER guess constant values.** Look up the exact enum name and value in `typelib_dump.json` under `Program/constant.tlb > enums`. Example: search for `FeaturePropertyConstants` to find `igLeft=1, igRight=2, igSymmetric=3`, etc.

2. **NEVER guess method signatures.** Look up the interface in the appropriate .tlb entry (usually `Program/Part.tlb`, `Program/assembly.tlb`, `Program/draft.tlb`, or `Program/framewrk.tlb`). The dump includes exact parameter names, types (VT_I4, VT_R8, VT_DISPATCH, SAFEARRAY, etc.), flags (in/out/optional), and return types.

3. **Check the implementation map first.** Before implementing a new tool, consult `reference/TYPELIB_IMPLEMENTATION_MAP.md` to see:
   - Whether the API method is already implemented
   - Which collection/interface the method belongs to
   - The exact method signature from the type library
   - Priority tier and any known issues (e.g., SAFEARRAY marshaling problems)

4. **Use correct collection-level APIs.** The type library shows which methods exist on collection interfaces (e.g., `ExtrudedCutouts.AddFiniteMulti`) vs. the `Models` interface (e.g., `Models.AddFiniteExtrudedProtrusion`). Prefer collection-level methods as they are proven to work.

5. **Cross-reference user-defined types.** When a parameter type shows a name like `FeaturePropertyConstants` or `RefAxis*`, look up that type in the same .tlb or in `constant.tlb` to find the actual values/interface.

6. **Update constants.py from type library data.** When adding new constants to `backends/constants.py`, copy the exact values from the scraped data. Add a comment noting which enum they came from.

### Quick Lookup Examples

```python
# To find a method signature:
# Search typelib_dump.json for: "AddAngularByAngle" in Program/Part.tlb > interfaces > RefPlanes > methods

# To find a constant value:
# Search typelib_dump.json for: "FeaturePropertyConstants" in Program/constant.tlb > enums

# To check what's implemented vs available:
# Read reference/TYPELIB_IMPLEMENTATION_MAP.md
```

## Development Workflow

When adding new capabilities:

1. **Check the implementation map**: Read `reference/TYPELIB_IMPLEMENTATION_MAP.md` to understand what's available and what's already done
2. **Look up the method signature**: Find the exact COM method in `reference/typelib_dump.json` - get parameter names, types, and order
3. **Verify constants**: Look up any enum values in the constants type library data, add to `backends/constants.py` if missing
4. **Backend first**: Implement the raw COM operation in the appropriate `backends/` module (e.g., `features.py` for new feature types)
5. **Test manually**: Use `python -i` to import and test the backend function directly, or use the diagnostic tools
6. **Wrap as tool**: Add `@mcp.tool()` decorator wrapper in `server.py` that calls the backend manager method
7. **Update tracking**: Update `reference/TYPELIB_IMPLEMENTATION_MAP.md` to mark the tool as implemented
8. **Add tests**: Write pytest tests in `tests/unit/` or `tests/integration/`
9. **Update docs**: Add to README.md if it's a major user-facing feature

### Common Development Tasks

**Testing a specific COM method:**
```python
# Use the diagnostic tools to inspect available methods
from src.solidedge_mcp.backends.diagnostics import diagnose_document
doc = doc_manager.get_active_document()
print(diagnose_document(doc))
```

**Adding a new feature type:**
1. Look up the method signature in `reference/typelib_dump.json` (search for the method name in the relevant .tlb)
2. Look up any required constants in `Program/constant.tlb > enums` and add to `backends/constants.py`
3. Add backend method to `backends/features.py` in the `FeatureManager` class, using exact parameter names/types from the type library
4. Add MCP tool wrapper in `server.py` in the appropriate section (marked with comments)
5. Follow the existing pattern: try/except, return dict with status or error
6. Update `reference/TYPELIB_IMPLEMENTATION_MAP.md` - check off the tool and update counts

**Checking tool count:**
```bash
grep -c "@mcp.tool()" src/solidedge_mcp/server.py
```

## Common Workflows

### Creating a Simple Extruded Part
```
1. connect_to_solidedge()
2. create_part_document()
3. create_sketch(plane="Top")
4. draw_rectangle(x1=0, y1=0, x2=0.1, y2=0.1)  # 100mm square (units in meters)
5. close_sketch()
6. create_extrude(distance=0.05, operation="Add")  # 50mm tall
7. save_document(file_path="C:/temp/box.par")
8. export_step(file_path="C:/temp/box.step")
```

### Creating a Revolved Part
```
1. connect_to_solidedge()
2. create_part_document()
3. create_sketch(plane="Front")
4. draw_line(x1=0, y1=0, x2=0.05, y2=0)  # Profile line
5. draw_line(x1=0.05, y1=0, x2=0.05, y2=0.1)
6. draw_line(x1=0.05, y1=0.1, x2=0, y2=0.1)
7. draw_line(x1=0, y1=0.1, x2=0, y2=0)  # Close profile
8. close_sketch()
9. create_revolve(angle=360)  # Full revolution
10. save_document(file_path="C:/temp/revolved.par")
```

### Assembly Workflow
```
1. connect_to_solidedge()
2. create_assembly_document()
3. place_component(component_path="C:/parts/base.par", x=0, y=0, z=0)
4. place_component(component_path="C:/parts/top.par", x=0, y=0, z=0.1)
5. list_assembly_components()  # Get component indices
6. create_mate(mate_type="Planar", component1_index=0, component2_index=1)
7. save_document(file_path="C:/assemblies/assembly.asm")
```

## Testing Notes

- **Unit tests**: Mock COM objects to test logic without Solid Edge installed
- **Integration tests**: Require Solid Edge running on Windows. Mark with `@pytest.mark.integration`
- **CI limitations**: GitHub Actions runners do not have Solid Edge. Integration tests run locally only.
- **Manual testing**: The easiest way to test is to run the MCP server and use it through Claude Code or another MCP client
- **Diagnostic tools**: Use `diagnose_api()` and `diagnose_feature()` to explore the COM API interactively

## Comparison to KiCad-MCP

| Aspect | KiCad-MCP | Solid Edge MCP |
|---|---|---|
| Platform | Cross-platform (Python, file I/O) | Windows-only (COM automation) |
| Read operations | S-expr parser, no KiCad needed | COM, requires Solid Edge running |
| Write operations | File mutation + kicad-cli | COM API calls |
| Session model | File-based undo/rollback | COM undo stack (pending) |
| Tool routing | 2-tier (8 direct, 67 routed) | TBD (likely simpler, fewer tools) |
| Primary use case | PCB design (board, schematic) | CAD modeling (parts, assemblies) |
