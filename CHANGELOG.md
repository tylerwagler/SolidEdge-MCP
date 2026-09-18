# Changelog

Notable changes, newest first. Dates are the day the entry landed on the branch.

## 0.1.0rc1 - 2026-09-18

First release candidate. The `modernize` branch (September 2026) turned a server of
wired-up COM calls into one whose results are checked against Solid Edge 2026, and every
number below is measured by the commands in `reference/VERIFICATION_STATUS.md`.

### Changed

- FastMCP 2.x server with `instructions`, MCP annotations (`readOnlyHint`,
  `destructiveHint`) and tags on every tool. Every tool and resource runs on one STA COM
  worker thread; a lost Solid Edge is detected and reattached.
- Surface: 119 composite tools, 53 data resources, 2 guide resources, 4 prompts. Related
  calls dispatch on a `method`/`type`/`action` discriminator.
- Units at the tool boundary are meters and degrees throughout; cone half-angles, feature
  extents, the camera and angular variables no longer leak radians.
- Errors decode the COM HRESULT and say what was being attempted. Tracebacks only with
  `SOLIDEDGE_MCP_DEBUG=1`.
- Feature status, feature type and document type are reported by name, not as integers.

### Fixed

- More than 140 COM calls that could never have worked: members that exist on no
  interface, members read off the wrong interface, wrong argument counts, SAFEARRAY
  wrappers Solid Edge rejects, constants from the wrong enum, writes to read-only
  properties, and parameters declared and never read.
- Creators that reported `status: "created"` and built nothing. Every material-changing
  creator now snapshots the body's face count and volume, and every collection-backed
  creator (reference planes, surfaces, sketches, documents, etches, threads, drafts, face
  rotates, parts lists, bend tables) checks that its collection grew.
- Assembly-level cutouts, holes (with a `HoleData`) and protrusions; assembly relations and
  mate/align constraints through `AssemblyDocument.CreateReference`; sketches in assemblies
  (`AsmRefPlanes`).
- Sheet metal: synchronous flanges (with bend deduction), splits, thickens, slots, beads,
  lofted flanges. Part: face rotates, the delete-blend planar guard, mirrors
  (volume-checked), lofts and sweeps with real `Origins`, helix cutouts, primitives.
- Drafts: dimensions by the element under a point, drawing views, layers, text box height,
  balloon positions, paper size in millimetres, symbol file origin read back, select-all.
- Variables: units and angle handling, rename, search, dimensions found by display name.
- Modal dialogs that hung the server: overwrite prompts (every write goes through
  `guard_overwrite`), the missing weldment template, the 3D-draw notice raised by
  structural frames.
- New tool `draw_3d_line`; structural frames run along 3D sketch lines.

### Refused with evidence

- 49 `create_*` methods return `unsupported: True` before touching COM, each carrying the
  Solid Edge defect, crash or UI-only selection behind it: swept surfaces and auto
  simplify (crash the process), lofted and bounded surfaces, threads on an existing
  cylinder, ordered flanges and louvers, contour flanges, blends, wires, assembly patterns
  and mirrors. See `reference/VERIFICATION_STATUS.md`.

### Tooling

- Structural audits, six against the scraped type libraries and four against the source
  tree alone: COM member names, receivers, signatures, enum arguments, refused writes,
  `hasattr` probes, swallowed writes reported as applied, dead parameters, duplicate mixin
  methods, and published counts that fail when they drift.
- `scripts/live_sweep.py` drives all 119 tools against a live Solid Edge and records the
  outcome in `reference/LIVE_SWEEP.md`: 116 OK, 16 honest refusals, 2 deliberate error
  probes across 134 cases.
- 80 integration tests, each mutation-tested when written, with part, sheet metal,
  assembly and draft fixtures; 2,969 unit tests.
