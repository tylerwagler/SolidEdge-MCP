# CLAUDE.md

Guidance for Claude Code when working in this repository.

## What this is

A Solid Edge MCP server: FastMCP 2.x over pywin32 COM automation. Windows only. MIT.
Surface: 119 tools, 53 data resources + 2 guide resources, 4 prompts.

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
uv run python scripts/audit_com_receivers.py
uv run python scripts/audit_com_writes.py
uv run python scripts/audit_com_hasattr.py
uv run python scripts/audit_reported_writes.py
uv run python scripts/audit_com_enum_args.py
uv run python scripts/audit_dead_params.py  # parameters declared and never read
uv run python scripts/count_verified_creators.py  # which create_* are checked, and how
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
│   ├── resources.py   53 solidedge:// resources (read-only JSON)
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

`reference/VERIFICATION_STATUS.md` tracks the other question — not "is this API wired up" but "does it actually do what it says". Read it before deciding what to work on, and re-measure it after landing anything that changes creator coverage or the audits. It carries the commands that regenerate every number in it.

Count tools with `grep -rc "register_tool(" src/solidedge_mcp/tools | awk -F: '{s+=$2} END {print s}'`.

## Solid Edge / COM rules

- **Units**: meters and degrees at the tool boundary. Convert to radians (`math.radians`) before COM.
- **Plane indices are 1-based**: 1=Top/XY, 2=Right/YZ, 3=Front/XZ (`RefPlanes.Item(n)`). Face/edge/feature/component indices are 0-based in tools and converted at the COM boundary.
- **Sketch then feature**: `create_sketch → draw_* → close_sketch → create_<feature>`. `close_sketch` calls `Profile.End(igProfileClosed)` and queues the profile in `sketch_manager.accumulated_profiles`; the feature consumes it.
- **Threading**: all COM runs on `com_thread`. Tool functions are plain sync `def`. Backend code may call other backend code freely (nested calls run inline).
- **Never `hasattr()` a COM proxy to test capability.** Use `com_get(obj, "Member")` and test for `None`, or check `doc.Type` against `DocumentTypeConstants`. The probe is a separate `GetIDsOfNames` round trip that reads False both for a member that is absent and for one whose getter raised, so a real error becomes a silent "unsupported" -- that is how every layer call came to refuse drafts, whose layers live on `Sheet`. `tests/unit/test_com_hasattr.py` (`scripts/audit_com_hasattr.py`) pins this at zero and needs no type library.
- **Never compare COM proxies with `==`**; compare `FullName`/`Name`.
- **Collections are 1-based** in COM (`Item(1)`).
- **An assembly's base planes are `AsmRefPlanes`, not `RefPlanes`.** An AssemblyDocument has no `RefPlanes` member at all; its three are `AsmRefPlanes`, same 1=Top/XY, 2=Right/YZ, 3=Front/XZ order, and `ProfileSets` works there exactly as in a part. `sketching.py: ref_planes_of(doc)` picks the right one. Reaching only for `RefPlanes` made every sketch in an assembly fail, which left all six assembly-level creators unreachable -- each consumes an accumulated profile and there was no way to make one.
- **A property's name is not its interface's name.** `AssemblyFeatures.ExtrudedProtrusions` returns an `AssemblyFeaturesExtrudedProtrusions`, and the code called the property by the interface name. `tests/unit/test_com_members.py` cannot see this -- the name is real, just not as a property -- so it is the receiver audit that catches it, and only once the receiver resolves. Helpers returning a tuple of COM objects need an entry in `TUPLE_SEEDS` for that to happen.
- **Assembly feature sides are `FeaturePropertyConstants`.** There is no `AssemblyFeaturePropertyConstants` enum in any library. Every side argument on `AssemblyFeatures*.Add` takes igLeft=1, igRight=2, igSymmetric=3; 0 is `igNullConstant`. Solid Edge accepts all of them without complaint and records the feature either way, so a wrong value shows up only as geometry that never changed -- the assembly cutout was passed 0 and never cut. `verifies_assembly_geometry` in `assembly/_base.py` counts occurrence body faces, because an AssemblyDocument has no `Models` and the part-level check is silently inert there.
- **An assembly hole needs a `HoleData`.** `AssemblyFeaturesHoles.Add` with `None` for `pHoledata` records nothing and raises nothing -- the collection's Count stays 0. Build one from the assembly's own `HoleDataCollection.Add(HoleType=igRegularHole, HoleDiameter=...)`; with it the same call cuts the placed part.
- **Assembly protrusions take no scope parts**, so there is no occurrence body for them to change; `occurrence_face_count` can only ever say "nothing happened". They are recorded -- `AssemblyFeatures.ExtrudedProtrusions.Count` grows -- and that collection is their verification signal. Cutouts and holes do take scope parts and are verified by occurrence faces.
- **A construction surface needs no solid.** `Constructions.*Surfaces.Add` works on an empty part; a "requires a base feature" guard is a precondition Solid Edge does not have, and here it masked two calls that had always failed.
- **`Application.ConvertByFilePath` returns without saying whether it did anything**, and can return before the writer finishes. The output file is what says a conversion happened; poll for it and report its size, or report an error.
- **`DesignEdgebarFeatures` is not only features.** It holds the three base reference planes as well, and `RefPlane` has no `Suppress` member at all, so a part with one suppressed extrusion reads as four entries of which three cannot be suppressed. Anything that walks the tree asking a feature-only question must skip the entries that cannot answer it rather than treating a missing member as "no".
- **SAFEARRAY marshalling is method-specific.** All of the following were verified against Solid Edge 2026, so change them only with new evidence:
  - `[in,out] SAFEARRAY(VT_R8)*` output buffers (`Body.GetRange`, `Occurrence.GetMatrix`, the mass-property buffers): pass a **plain Python list** and read the filled values from the **return value**. A `VARIANT` wrapper, with or without `VT_BYREF`, raises `Objects for SAFEARRAYS must be sequences`. Helpers: `query/_base.py: r8_array/i4_array/bool_array`.
  - Profile and edge arrays: `Rounds.Add` accepts `VARIANT(VT_ARRAY | VT_DISPATCH, [...])` and is integration-tested that way, but the helix APIs reject it and need a plain `[profile]`. When in doubt, a plain sequence is the safer default.
  - A parameter the type library declares `[in] VT_R8*` that Solid Edge actually fills (`DraftDocument.GetSymbolFileOrigin`): the late-bound call returns `None` and the values are lost. Invoke it through `doc._oleobj_.InvokeTypes` with the parameters declared `(VT_BYREF | VT_R8, 3)` and read the returned tuple; see `export/_draft.py`.
  - Enumerating `Variables.Item(i)` shows a dimension as a nameless entry; `Variables.Item("<display name>")` returns it, and `Variables.Query` finds it only with `NamedBy=seVariableNameByBoth` (ByUser sees user variables only, BySystem only dimensions and PhysicalProperties_*). Angular variables hold radians; `UnitsType` (not `Units`, which is a member of nothing) says which.
  - pywin32 gives every parameter a positional slot, `[out]` ones included. When an out-parameter sits between ones you must supply, pass the later arguments **by keyword** using the type library's parameter names.
- **Cutouts** use collection-level APIs (`model.ExtrudedCutouts.AddFiniteMulti`), not `Models.AddExtrudedCutout`.
- **A loft or sweep pairs its cross-sections through the `Origins` array**, and each entry must be a point that lies on its own section. `comutil.profile_origin` takes a start point from lines and arcs and a centre from circles. A hardcoded `(0, 0)` makes Solid Edge build nothing at all -- no error, no geometry -- unless every profile happens to cross the sketch origin.
- **3D sketch lines are drawable**: `Sketches3D.Add()` then `Lines3D.Add(x1, y1, z1, x2, y2, z2)`, in parts and assemblies (`SketchManager.draw_line_3d`, the `draw_3d_line` tool). `StructuralFrames.Add(part, n, [Line3D...], VT_EMPTY x3)` runs a frame along them. **The first frame in a session raises an informational "The Segments group of commands are replaced with the 3D Draw group..." dialog that `DisplayAlerts` does not suppress and that blocks the call until OK is clicked**; `backends/dialogs.py: dismiss_informational_dialog(text_prefix)` clicks OK on exactly that dialog while the call runs and nothing else. `Wires.Add` still answers E_FAIL along such a line.
- **Slots.Add cuts along an open line** with `igRight` and a finite or through-all extent (6 -> 10 faces); `igLeft` and a closed profile record a slot that removes nothing. Its two `KeyPointFlags` take 0, which is in no member of `KeyPointExtentConstants`; the enum audit's `ALLOWED` carries that pair.
- **Two calls that want faces or bodies this server *can* obtain**: `Splits.Add(1, [model.Body], 1, [RefPlane], 0, 0)` splits into two design bodies (the first model keeps its faces, so the check watches `Models.*.Splits`); `Models.AddThickenFeature(side, distance, n, faces)` takes `Constructions.Item(n).Body.Faces` -- the surface's own `Faces` property raises, and `Thickens.AddSync` lives on a Model a surface-only document does not have.
- **`@verifies_geometry` watches face count and body volume.** A mirror of a box across its own face is a wider box with the same six faces; only the volume moves. The snapshot is `(models, faces, volume)`, and the volume decides when the faces are unchanged and it can be read on both sides.
- **Known unsupported via COM (SE 2025/2026)**: `AssemblyFeaturesPatterns.Add`, `AssemblyFeaturesMirrors.Add` (E_ACCESSDENIED); shell/thin-wall (needs interactive face pick); multiple disjoint profiles in one cutout sketch; `AssemblyFeatures.Recompute` (E_FAIL whatever it is given -- use `AssemblyDocument.UpdateAll`); structural frames by orientation (not driven; `method='basic'` builds along `draw_3d_line` lines -- see below); the ordered `Flanges.Add*` (record a Flange that never solves: range, volume and face count unchanged after Recompute, `Status` raises) -- `Flanges.AddSync` **does build, in a document set to synchronous before its base tab** (`create_flange` routes there when the document is synchronous), and switching after an ordered tab answers E_FAIL; `ContourFlanges.AddSync` (E_INVALIDARG in that same document); `Threads.Add` on an existing cylinder (E_INVALIDARG for a boss and a hole with every HoleData; a thread has to come with its hole); `Louvers.Add` (same never-solves shape) and `Louvers.AddSync` (0x807B0086); `ContourFlanges.AddEx`/`Add` (E_FAIL on 52 open-profile placements); `Constructions.LoftedSurfaces.Add`/`Add2` and `BlueSurfs.Add` (E_INVALIDARG on every shape that builds a solid loft); `Models.AddLoftedFlange` (DISP_E_TYPEMISMATCH on four origin shapes); `Relations3d.AddPlanar` (faces from `Occurrence.Body` raise E_FAIL, faces from the occurrence's part document answer 0x80040225); `Constructions.SweptSurfaces.Add` (**crashes the Solid Edge process** on 3 of 5 calls once a session has been through a few documents -- it builds on a fresh instance with `Origins=[element]`, `OriginRefs=keypoint`; a crash is worse than a refusal, so the tool refuses).
- **`Documents.Add("SolidEdge.WeldmentDocument")` raises a modal.** The ProgID is registered and Solid Edge accepts it, then looks for its default weldment template; on an install without the weldment environment that file is absent and the answer is a modal "Path not found" that `DisplayAlerts` does not suppress, which hangs the server until someone clicks it. Only then does the call return `0x80030003`. `create_weldment` therefore never takes the ProgID route and requires an existing template path. A template a caller names that is not on disk is refused by every document creator before any COM call -- it used to fall through silently to the default document.
- **Front plane quirk**: COM "Normal" on the Front plane points to world −Y. Cutout tools do not auto-swap; prefer `direction="Symmetric"` there.
- **Exports** use `SaveCopyAs`, never `SaveAs` (which repoints the live document).
- **Never hand Solid Edge a path that already exists.** It answers with a modal "This file exists. Do you want to overwrite it?" prompt that `DisplayAlerts` does not suppress. Solid Edge has one UI thread, so the COM call never returns, every later call queues behind it, and the client eventually reports `Connection closed` as if the server had crashed. Every write goes through `backends/validation.py: guard_overwrite(file_path, overwrite)` first, which refuses by default and deletes the file when `overwrite=True`. Any new call that writes a path needs the same guard and an `overwrite` parameter.

## Type library reference

`reference/typelib_dump.json` is the source of truth for COM signatures and enum values but is **gitignored and absent from a fresh clone**. Regenerate it with `uv run python scripts/scrape_typelibs.py` (requires Solid Edge installed). Until then use `reference/typelib_summary.md` (truncated to the first values of each enum) and `reference/TYPELIB_IMPLEMENTATION_MAP.md`.

Rules: never guess a constant or signature. Look it up, copy the exact value into `constants.py` with a comment naming the enum, and prefer collection-level `Add*` methods.

Six checks enforce this, and all six skip when the dump is absent:

| Check | What it catches |
|---|---|
| `tests/unit/test_constants_typelib.py` | A constant whose value disagrees with its enum. Classes that are our own vocabulary go in `LOCAL_GROUPINGS` with a note. |
| `tests/unit/test_com_members.py` | A COM member name that exists in no type library. A ratchet: new names fail, and fixing one fails until you delete it from `UNVERIFIED`. |
| `tests/unit/test_com_receivers.py` (`scripts/audit_com_receivers.py`) | A member read off an interface that does not have it, which the name check cannot see because the name is real elsewhere. |
| `scripts/audit_com_signatures.py` | A call with the wrong number of arguments. Run `--filter <path>` to see the full parameter list for each finding, `--by-file` for counts. |
| `tests/unit/test_com_enum_args.py` (`scripts/audit_com_enum_args.py`) | A constant whose value is in no member of the enum its parameter declares -- a stray literal, or a value from the wrong constants class that falls outside the target enum. It **cannot** see the wrong *member* of the right enum, which is what most constant bugs here have been; only live geometry verification catches those. Bitmask enums accept 0 and any combination. |
| `tests/unit/test_com_writes.py` (`scripts/audit_com_writes.py`) | An assignment to a member that is a method, or to a property the type library marks read-only. Solid Edge answers "Property 'Item.X' can not be set." and a surrounding try/except turns that into a reported success. |

The signature audit resolves the receiver with the same inference the receiver audit uses, so `doc.Occurrences.Item(1)` is checked against `Occurrence` specifically rather than against every interface with that method name. That matters because its fallback is weak on purpose: an unresolved receiver only has to fit *some* interface with a method of that name, which is how `occurrence.Replace(path)` passed while `Occurrence.Replace` requires two arguments.

The receiver audit checks lowercase members too, once the receiver resolves. PascalCase is the usual spelling but 857 lowercase member names exist across the libraries, and requiring PascalCase hid two bugs of one shape: `textbox.x` and `balloon.x` are members of nothing, so the position was quietly missing from every text box and balloon reported. Our own attributes stay out of it because `self.active_profile` resolves to no interface.

The receiver audit goes further and infers what a receiver *is* by following declared types: `doc.Models.Item(1).Features` resolves PartDocument to Models to Model to Features. That is what catches the sharpest class of bug here, a real name on the wrong interface, such as `model.RevolvedSurfaces` when RevolvedSurfaces belongs to `Constructions`, or `line.StartPoint.X` when Line2d only has `GetStartPoint()`.

Four more checks need no type library:

`tests/unit/test_manager_mro.py` covers a different trap: the managers are built from a dozen mixins each, and two mixins defining the same method leaves one silently unreachable.

`tests/unit/test_reported_writes.py` (`scripts/audit_reported_writes.py`) fails when a COM write sits inside a `contextlib.suppress` or a `try` that passes, and the value written is then handed back in the result. If Solid Edge refuses the write the failure is hidden and the caller is told it applied. This shape has produced a bug every time it was checked: `TextBox.TextHeight`, `Leader.Text` and `FeatureControlFrame.Text` are on no interface, `variable.DisplayName` is read-only, `DraftPrintUtility.PaperWidth` is millimetres and was given meters, and `PMI.Show` is refused on a part with no PMI. Let the write raise, report what Solid Edge holds afterwards, or add the pair to `ALLOWED` once it has been driven live.

`tests/unit/test_dead_params.py` (`scripts/audit_dead_params.py`) fails when a parameter is declared and never read, at either layer. That is this server's quietest bug: the call succeeds, the result reports what the caller asked for, and the value never reached Solid Edge, so there is nothing for it to reject. `add_adjustable_part(x, y, z)` placed the part at the origin, `create_parts_list(x, y)` let Solid Edge choose the position, and `create_revolve(axis_type)` offered a choice nothing consulted. A parameter with genuinely nowhere to go says so with `del <param>`, which the audit reads as deliberate.

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
