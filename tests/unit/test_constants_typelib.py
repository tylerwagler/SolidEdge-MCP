"""Verify backends/constants.py against the scraped Solid Edge type libraries.

CLAUDE.md makes the type library the source of truth for every COM enum value,
but nothing enforced it: a wrong value silently produces wrong geometry, and an
audit found 15 of them. This test closes that loop.

It needs ``reference/typelib_dump.json``, which is gitignored because it is
~20 MB. Regenerate it on a machine with Solid Edge installed:

    uv run python scripts/scrape_typelibs.py

Without the dump the test skips, so a fresh clone and CI stay green.
"""

from __future__ import annotations

import ast
import json
import pathlib
from collections import defaultdict

import pytest

REPO_ROOT = pathlib.Path(__file__).resolve().parents[2]
DUMP_PATH = REPO_ROOT / "reference" / "typelib_dump.json"
CONSTANTS_PATH = REPO_ROOT / "src" / "solidedge_mcp" / "backends" / "constants.py"

# Classes in constants.py that group values under a name Solid Edge does not
# use for an enum. They are our own vocabulary, so there is nothing to compare
# against; each entry says where the values actually come from.
LOCAL_GROUPINGS: dict[str, str] = {
    # Convenience aliases for FeaturePropertyConstants members (igLeft/igRight/...)
    "DirectionConstants": "aliases of FeaturePropertyConstants",
    "ExtentTypeConstants": "aliases of FeaturePropertyConstants (igFinite=13, igThroughAll=16)",
    "FaceQueryConstants": "aliases of FeatureTopologyQueryTypeConstants",
    "LoftSweepConstants": "aliases of FeaturePropertyConstants",
    "NormalCutoutMethodConstants": "aliases of FeaturePropertyConstants (igSM*Cutout)",
    "ProfileValidationConstants": "aliases of ProfileValidationType, plus composites",
    "ReferenceElementConstants": "aliases of ReferenceElementConstants members",
    "RefPlaneConstants": "our 1-based RefPlanes.Item() ordering, not a COM enum",
    "HoleTypeConstants": "aliases of FeaturePropertyConstants",
    "FoldTypeConstants": "aliases of FeaturePropertyConstants",
    "KeyPointTypeConstants": "aliases of KeyPointType",
    # Solid Edge exposes no enum for these; values were determined empirically.
    "FeatureOperationConstants": "not in any type library; unused by COM calls",
    "MateTypeConstants": "assembly relation ordinals; no matching COM enum",
    "AssemblyFeaturePropertyConstants": "assembly feature ordinals; no matching COM enum",
    "AssemblyRelationConstants": "Relation3d object-type ids, not an enum",
    "DrawingViewOrientationConstants": "empirically verified draft view ordinals",
    "DrawingViewTypeConstants": "draft view ordinals",
    "PatternTypeConstants": "pattern ordinals; no matching COM enum",
    "PatternOffsetTypeConstants": "pattern ordinals; no matching COM enum",
    "PatternTransformTypeConstants": "pattern ordinals; no matching COM enum",
    "PatternTransformRotateTypeConstants": "pattern ordinals; no matching COM enum",
    "PatternCurveAnchorSideConstants": "pattern ordinals; no matching COM enum",
    "ViewOrientationConstants": "view ordinals; no matching COM enum",
    "RenderModeConstants": "render mode ordinals; no matching COM enum",
    "ModelingModeConstants": "modeling mode ordinals; no matching COM enum",
    "SaveAsConstants": "our own flag, not a COM enum",
    "DocumentTypeConstants": "aliases of DocumentTypeConstants members (verified below)",
}

# Members we knowingly keep even though the type library has no such value.
KNOWN_EXTRA: set[tuple[str, str]] = {
    # Solid Edge has no combined crown+draft value; kept so callers resolve.
    ("TreatmentTypeConstants", "seTreatmentCrownAndDraft"),
}


def _load_enums() -> dict[str, dict[str, int]]:
    data = json.loads(DUMP_PATH.read_text(encoding="utf-8"))
    enums: dict[str, dict[str, int]] = defaultdict(dict)
    for payload in data["typelibs"].values():
        for enum_name, values in (payload.get("enums") or {}).items():
            if isinstance(values, dict):
                enums[enum_name].update(
                    {
                        k: v
                        for k, v in values.items()
                        if isinstance(v, int) and not isinstance(v, bool)
                    }
                )
    return dict(enums)


def _load_constants() -> dict[str, dict[str, tuple[int, int]]]:
    """class name -> {member: (value, lineno)}."""
    tree = ast.parse(CONSTANTS_PATH.read_text(encoding="utf-8"))
    out: dict[str, dict[str, tuple[int, int]]] = {}
    for node in tree.body:
        if not isinstance(node, ast.ClassDef):
            continue
        members: dict[str, tuple[int, int]] = {}
        for stmt in node.body:
            if isinstance(stmt, ast.Assign) and isinstance(stmt.value, ast.Constant):
                value = stmt.value.value
                if isinstance(value, int) and not isinstance(value, bool):
                    for target in stmt.targets:
                        if isinstance(target, ast.Name):
                            members[target.id] = (value, stmt.lineno)
        if members:
            out[node.name] = members
    return out


pytestmark = pytest.mark.skipif(
    not DUMP_PATH.exists(),
    reason="reference/typelib_dump.json absent; run scripts/scrape_typelibs.py",
)


@pytest.fixture(scope="module")
def enums() -> dict[str, dict[str, int]]:
    return _load_enums()


@pytest.fixture(scope="module")
def constants() -> dict[str, dict[str, tuple[int, int]]]:
    return _load_constants()


def test_dump_looks_complete(enums):
    assert len(enums) > 500, f"only {len(enums)} enums parsed; dump may be truncated"
    # A value every feature call depends on.
    assert enums["FeaturePropertyConstants"]["igRight"] == 2


def test_every_directly_named_enum_matches(enums, constants):
    """A constants class named after a real enum must match it value for value."""
    problems: list[str] = []
    compared = 0
    for cls_name, members in constants.items():
        if cls_name in LOCAL_GROUPINGS:
            continue
        enum = enums.get(cls_name)
        assert enum is not None, (
            f"{cls_name} is neither a type library enum nor listed in LOCAL_GROUPINGS. "
            "Add it to LOCAL_GROUPINGS with a note on where its values come from."
        )
        for member, (value, lineno) in members.items():
            if member not in enum:
                if (cls_name, member) not in KNOWN_EXTRA:
                    problems.append(
                        f"constants.py:{lineno} {cls_name}.{member} is not a member of {cls_name}"
                    )
                continue
            compared += 1
            if enum[member] != value:
                problems.append(
                    f"constants.py:{lineno} {cls_name}.{member} = {value}, "
                    f"type library says {enum[member]}"
                )
    assert compared >= 20, f"only {compared} constants compared; the audit is not doing its job"
    assert not problems, "constant values disagree with the type library:\n  " + "\n  ".join(
        problems
    )


def test_alias_groupings_resolve_somewhere(enums, constants):
    """Every alias grouping member must exist with that value in some enum.

    Catches a typo'd alias without asserting which enum it came from.
    """
    by_member: dict[str, set[int]] = defaultdict(set)
    for enum in enums.values():
        for name, value in enum.items():
            by_member[name].add(value)

    alias_classes = {
        "DirectionConstants",
        "ExtentTypeConstants",
        "FaceQueryConstants",
        "LoftSweepConstants",
        "NormalCutoutMethodConstants",
        "HoleTypeConstants",
        "FoldTypeConstants",
        "KeyPointTypeConstants",
    }
    problems: list[str] = []
    for cls_name in alias_classes:
        for member, (value, lineno) in constants.get(cls_name, {}).items():
            seen = by_member.get(member)
            if seen is None:
                problems.append(f"constants.py:{lineno} {cls_name}.{member} exists in no enum")
            elif value not in seen:
                problems.append(
                    f"constants.py:{lineno} {cls_name}.{member} = {value}, "
                    f"type library values for that name: {sorted(seen)}"
                )
    assert not problems, "alias constants disagree with the type library:\n  " + "\n  ".join(
        problems
    )


@pytest.mark.parametrize(
    ("enum_name", "member", "expected"),
    [
        # Load-bearing values the review flagged as unverifiable.
        ("FeaturePropertyConstants", "igFinite", 13),
        ("FeaturePropertyConstants", "igThroughAll", 16),
        ("FeaturePropertyConstants", "igProfileBasedCrossSection", 48),
        ("FeaturePropertyConstants", "igLeft", 1),
        ("FeaturePropertyConstants", "igRight", 2),
        ("FeaturePropertyConstants", "igSymmetric", 3),
        ("DocumentTypeConstants", "igWeldmentDocument", 6),
        ("DocumentTypeConstants", "igWeldmentAssemblyDocument", 7),
        ("ProfileValidationType", "igProfileClosed", 1),
        ("ProfileValidationType", "igProfileAllowNested", 8192),
    ],
)
def test_flagged_values(enums, enum_name, member, expected):
    assert enums[enum_name][member] == expected
