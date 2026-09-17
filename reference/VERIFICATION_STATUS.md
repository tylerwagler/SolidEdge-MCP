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
| Unit tests | 2,897 | mocked COM; catch shape, not truth |
| Integration tests | 13 | drive real Solid Edge; between them they touch **16 of 118** tools |
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

### 2. No check verifies a value is numerically right

Every check answers *did the write take* or *did the face count change*. Nothing
answers *is this number correct*. A mass, a centre of gravity or a bounding box
could be wrong by a unit factor and every gate stays green. `DraftPrintUtility.
PaperWidth` being handed metres when it wanted millimetres was exactly this, and
it was found by hand.

### 3. Four assembly creators still build nothing

`create_assembly_hole`, `create_assembly_extruded_protrusion`,
`create_assembly_revolved_protrusion` refuse honestly now rather than reporting
success; `create_assembly_swept_protrusion` answers `E_INVALIDARG`. The cutouts
were fixed and build. Whether the rest can work through COM on 2026 is open.

### 4. `create_contour_flange` does not build

It needs an open profile positioned against a specific edge, which the sketch
tools cannot guarantee. It fails honestly, which is the floor, not a fix.

### 5. The swallow surface is only partly audited

771 `except Exception` and 210 `contextlib.suppress` in the backends. The
write-then-reported slice is pinned by `audit_reported_writes`. The read side was
measured and is clean: only 10 swallowed reads have a literal fallback and 9 of
those fall back to `None`, which a caller can tell apart from a real value.

### 6. Nothing is pushed

70 commits on `modernize`. `master` has not moved since the merge of #3. No PR.

## What "done" would mean

There is no state where this is finished, because Solid Edge keeps its own
counsel. A defensible bar:

- [ ] every tool driven live at least once, with the result recorded
- [ ] every bug fixed this far pinned by a test that fails without the fix
- [x] the 49 unverified creators given a collection-growth check
- [ ] one numeric-correctness check per measurement tool
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
