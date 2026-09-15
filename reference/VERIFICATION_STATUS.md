# Verification Status

**Last measured: 2026-09-15**, against Solid Edge 2026 (226.00.01.04), branch `modernize`.

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
| Tools | 119 | the MCP action surface |
| Resources | 53 + 2 guides | read-only `solidedge://` endpoints |
| Unit tests | 2,897 | mocked COM; catch shape, not truth |
| Integration tests | 13 | drive real Solid Edge |
| Structural audits | 7 | all at zero but one documented finding |

### Creator verification

`@verifies_geometry` (and `@verifies_assembly_geometry`) snapshot the body before
and after, and downgrade a false success to an explicit error. This is the only
check that catches a call which succeeds and builds nothing.

| | count | |
|---|---|---|
| Verified live | **127** | the body must change or the result becomes an error |
| Refuse honestly | 36 | return `unsupported: True` without touching COM |
| Neither | **50** | see below |
| **Total `create_*`** | **213** | |

The 50 unverified break down as:

| count | file | why |
|---|---|---|
| 15 | `features/_ref_planes.py` | a plane is not a solid; no face count to change |
| 15 | `features/_surfaces.py` | builds constructions, not body material |
| 5 | `documents.py` | creates documents |
| 3 | `features/_misc.py` | face rotations and draft angles reshape faces without changing the count |
| 3 | `features/_rounds_chamfers.py` | surface blends build constructions |
| 3 | `features/_sheet_metal.py` | cosmetic threads and etches change no material |
| 2 | `export/_drawing.py` | drawings and parts lists are not geometry |
| 2 | `sketching.py` | profiles, not solids |
| 1 | `assembly/_relations.py` | `create_mate` constrains, it does not build |
| 1 | `export/_draft.py` | a bend table is a table |

**Roughly 40 of the 50 genuinely cannot be face-counted.** They are not a backlog.
The honest gap is the remaining ~10, which need a different kind of check
(does the plane exist afterwards, did the surface get added to `Constructions`)
rather than the one we have.

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

**13 integration tests against 119 tools.** Almost every bug found so far was
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
- [ ] the ~10 genuinely-unverified creators given a check that suits them
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

Creator counts come from walking `backends/` for `create_*` and checking for
`@verifies_geometry`, `@verifies_assembly_geometry`, the class-level
`@verify_geometry_on_creators`, or an `"unsupported": True` return. A grep for
the decorator alone undercounts badly — the class decorator covers whole mixins.
