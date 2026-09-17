# Verification Status

**Last measured: 2026-09-16**, against Solid Edge 2026 (226.00.01.04), branch `modernize`.

Regenerate every number here with the commands in [How to re-measure](#how-to-re-measure).
Nothing in this document is an estimate.

## Why this document exists

`TYPELIB_IMPLEMENTATION_MAP.md` tracks a different question — *is this COM API
wired up?* — and answers 96%. That measure has been wrong in the same direction
repeatedly: it counts a tool as done when the call exists, and this server's
defining bug class is **a call that exists, returns `status: "created"`, and
changes nothing**.

Six assembly-level creators were counted as implemented and had never once been
able to run: they consume a sketch profile, and no sketch could be created in an
assembly at all. Nothing failed. They reported "No profiles available", which
reads like a caller mistake.

So this document tracks the other question: **does it actually do what it says?**

## Where we are

| | count | what it means |
|---|---|---|
| Tools | 118 | the MCP action surface |
| Resources | 53 + 2 guides | read-only `solidedge://` endpoints |
| Unit tests | 2,949 | mocked COM; catch shape, not truth |
| Integration tests | 32 | drive real Solid Edge; 19 of them pin exact values against a known box |
| Structural audits | 7 | all at zero but one documented finding |

### Creator verification

Two decorator families snapshot the document before and after a creator runs
and downgrade a false success to an explicit error. `@verifies_geometry` (and
`@verifies_assembly_geometry`) watch the body's face count; for creators that
build no solid -- a plane, a surface, a sketch, a document -- that count never
moves, so `@verifies_collection_growth(...)` watches the COM collection their
`Add` lands in instead. This is the only kind of check that catches a call
which succeeds and builds nothing.

| | count | |
|---|---|---|
| Verified live | **176** | the body or the target collection must change, or the result becomes an error |
| Refuse honestly | 37 | return `unsupported: True` -- 34 outright, plus 3 that refuse one argument value and otherwise work (`create_extrude`/`create_revolve` with `operation="Intersect"`, `create_revolve_full` with a treatment) |
| Neither | **0** | |
| **Total `create_*`** | **213** | |

Of the 176, 127 are face-count checks and 49 are collection-growth checks:
15 reference planes (`RefPlanes`), 15 surfaces (the five `Constructions.*Surfaces`
summed), 9 that reshape or annotate (`Models.*.Etches`, `.Threads`, `.Drafts`,
`.FaceRotates`; `PartsLists`, `DraftBendTables`), 3 surface blends
(`Models.*.Rounds` + `.Blends`), 2 sketches (`ProfileSets`), and 6 document
creators (`Application.Documents`). Each group was driven live: the snapshot
reads a real integer on the right document type, a known-good creator moves
it and passes through, and a creator that claims success without adding is
downgraded to an error on the live path -- not only against fakes.

### Structural audits

All seven are ratchets: a new violation fails the suite.

| audit | catches | status |
|---|---|---|
| `audit_com_signatures` | wrong argument count | 0 |
| `audit_com_receivers` | a real member read off the wrong interface | 1 known (`SaveAsPLMXML`) |
| `audit_com_writes` | assignment to a method or read-only property | 0 |
| `audit_com_hasattr` | `hasattr` probes of COM members | 0 |
| `audit_reported_writes` | a swallowed write reported as applied | 0 |
| `audit_com_enum_args` | a value outside the enum its parameter declares | 0 |
| `audit_dead_params` | a parameter declared and never read | 0 |

## What is left

Ordered by how much risk each removes, not by effort.

### 1. Live verification is barely reproducible

**13 integration tests, touching 16 of 118 tools.** All 14 assembly tools, all 18
draft/export tools, all 11 sheet-metal tools and all 6 surface tools have none.
Almost every bug found so far was
caught by a throwaway script that was then thrown away. Three have been
converted (`tests/integration/test_reported_values.py`) and each was
mutation-tested by restoring the original defect. The rest are gone.

This is the largest gap, because it is the one that lets fixed bugs come back.

### 2. Numeric correctness -- now checked, one item open

`tests/integration/test_known_answer_box.py` builds a 0.08 x 0.048 x 0.03 m box
and asserts, at float-noise tolerance, every number that can be computed by
hand: volume, surface area, each face area, counts, bounding box, centre of
gravity, mass at two densities, the body diagonal, a right angle, and the six
global moments of inertia with their frame (about the model origin) and sign
(products positive) -- plus five cross-path invariants that need no reference
value. Nineteen of nineteen pass live.

Writing it found and fixed: vertices reported as one-element points (`result[0]`
of a flat `(x, y, z)`); a short physical-property tuple padded with zeros and
reported as computed; a partial surface-area sum reported as the total; moments
of inertia that silently assumed steel; `GetDirection{1,2}Extent` read with its
slots shifted so the side constant was reported as the distance and the real
value stringified into a non-existent `face_ref`; a cone `half_angle` and the
camera's perspective field of view reported in radians.

**Open:** `Variable.Value` for an *angular* variable is unconfirmed. Neither
`Variables.Add(name, "45 deg")` nor an explicit `units_type` would create one
on this install, and `query_variables("*")` returns only user-defined variables
-- `doc.Variables.Count` reads 4 while the query reports 0 -- so a revolve's own
angle dimension cannot be observed through this server either. Two things are
established: a bare-number formula is read in the *document's* units (on an
inch template `"0.785398"` became 0.0199 m), and the variable query never sees
dimension variables. Whether angular values arrive in radians is not.

### 3. Known-broken, with the evidence

`scripts/live_sweep.py` drove all 118 tools once through the real server and
recorded every outcome in `reference/LIVE_SWEEP.md`. What it left as
known-broken, each with the evidence that closes the question on this install:

| tool / method | evidence |
|---|---|
| `create_flange` (basic) | 24 face/edge combinations through the tool and 18 `FlangeSide`/`ThicknessSide` combinations through raw COM on a fresh tab: every one builds nothing or answers `E_POINTER`. `Flanges.Add` never builds from a `Face.Edges` edge here. Reports honestly via the decorator. |
| `create_lofted_surface`, `_v2`, `create_swept_surface` | `E_INVALIDARG` / `E_FAIL` with a base feature and without, while `create_extruded_surface` builds beside them. A "requires a base feature" guard masked this for years; it is gone and the COM error shows. |
| `create_assembly_hole`, `_extruded_protrusion`, `_revolved_protrusion` | record a feature and change no occurrence body; refuse honestly via `verifies_assembly_geometry`. Under investigation: a protrusion takes no scope parts and may build an assembly-level body the check does not read. |
| `create_assembly_swept_protrusion` | `E_INVALIDARG`. |
| `create_contour_flange` | `E_FAIL`; it needs an open profile against a specific edge the sketch tools cannot guarantee. |
| `select_set(action="all")` | `SelectSet.AddAll` answers `E_FAIL` whatever the selection holds; now refuses honestly. |
| `draft_config(action="get_origin")` | raw `DISP_E_BADINDEX`; not yet investigated. |
| `face_operation(rotate_by_edge)` | raw `E_INVALIDARG` on a box face; not yet investigated. |
| `create_weldment` without a template | raised a modal that hung the server; now refuses before COM. |

Fixed because the sweep found them: `convert_by_file_path` reported
"converted" with no output on disk; vertices came back as one-element points;
extent slots were shifted so the side constant was reported as the distance.

### 4. The swallow surface is only partly audited

771 `except Exception` and 210 `contextlib.suppress` in the backends. The
write-then-reported slice is pinned by `audit_reported_writes`. The read side was
measured and is clean: only 10 swallowed reads have a literal fallback and 9 of
those fall back to `None`, which a caller can tell apart from a real value.

### 5. Nothing is pushed

70 commits on `modernize`. `master` has not moved since the merge of #3. No PR.

## What "done" would mean

There is no state where this is finished, because Solid Edge keeps its own
counsel. A defensible bar:

- [ ] every tool driven live at least once, with the result recorded
- [ ] every bug fixed this far pinned by a test that fails without the fix
- [x] the 49 unverified creators given a collection-growth check
- [x] one numeric-correctness check per measurement tool (angular variables excepted, see above)
- [ ] the branch merged

## How to re-measure

```bash
uv run pytest -q                       # unit
uv run pytest -m integration -q        # needs a running, licensed Solid Edge
uv run ruff check . && uv run mypy src/
for a in signatures receivers writes hasattr enum_args dead_params; do
  uv run python scripts/audit_com_$a.py
done
uv run python scripts/audit_reported_writes.py
```

```bash
uv run python scripts/count_verified_creators.py   # the creator table above
```

Do not count creators with a grep for a decorator name. There are two
decorator families and two class-level forms that cover whole mixins, and a
grep for one name has undercounted this table twice.
