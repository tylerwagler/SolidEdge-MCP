"""Check the COM member names the backends call against the type libraries.

Every mock in this suite is a ``MagicMock``, which answers to any attribute. A
call to a misspelled COM member therefore passes every unit test and only fails
against real Solid Edge. This test closes that hole statically: it collects
every PascalCase attribute the backend modules touch and checks it against the
set of interface methods, properties, enums, and enum members in the scraped
type libraries.

It is a ratchet. ``UNVERIFIED`` pins the names that are currently unmatched;
anything new fails, and fixing one fails too, prompting its removal from the
list. Both directions keep the list honest.

Needs ``reference/typelib_dump.json`` (gitignored, ~20 MB). Regenerate with:

    uv run python scripts/scrape_typelibs.py
"""

from __future__ import annotations

import ast
import json
import pathlib
import re
from collections import defaultdict

import pytest

REPO_ROOT = pathlib.Path(__file__).resolve().parents[2]
DUMP_PATH = REPO_ROOT / "reference" / "typelib_dump.json"
BACKENDS = REPO_ROOT / "src" / "solidedge_mcp" / "backends"

PASCAL_CASE = re.compile(r"^[A-Z][A-Za-z0-9]*$")

# Names that look like COM members but are not, or are COM members the scrape
# does not cover. Grouped by why they are here. Shrink this list, never grow it
# without a reason.
UNVERIFIED: frozenset[str] = frozenset(
    {
        # --- not COM at all: Python stdlib and pywin32 -----------------------
        "CoInitializeEx",
        "DEBUG",
        "Dispatch",
        "EnsureDispatch",
        "GetIDsOfNames",  # IDispatch itself, via doc._oleobj_ (export/_draft.py)
        "InvokeTypes",  # same: declared-type invoke for [in, out] parameters
        "Formatter",
        "GetActiveObject",
        "Logger",
        "StreamHandler",
        "WARNING",
        # --- plain Python attributes on our own value objects ----------------
        "X",
        "Y",
        "Z",
        # --- Solid Edge members absent from the scraped libraries ------------
        # Removed after being fixed against Solid Edge 2026: AddRadial,
        # AddDiameter, AddOrdinate and AddRadialDimension became
        # Dimensions.AddRadius / AddCircularDiameter / AddRadialDiameter /
        # AddCoordinate; StartX, StartY, EndX, EndY, CenterX and CenterY
        # became GetStartPoint / GetEndPoint / GetCenterPoint; EndAngle
        # became StartAngle + SweepAngle; and FCFs became
        # FeatureControlFrames. See tests/unit/test_export_annotations.py.
        # Also removed: AlignToView and RemoveAlignment became the
        # collection's DrawingViews.Align / Unalign, driven through the
        # document select set; XPosition, YPosition, OriginX and OriginY on
        # a drawing view became SetOrigin / GetOrigin; DisplayMode became
        # Shading and Defaults_ShowHiddenEdges;
        # AddAssemblyViewWithConfiguration became a seventh argument to
        # AddAssemblyView; AddPartViewWithConfiguration became
        # AddPartViewByConfiguration; and DraftPrintUtility.PrintAllSheets
        # became AddDocument / AddSheet.
        # And: MoveAfter/MoveBefore became Feature.Reorder(target,
        # InsertBefore); ConvertToType became ConvertToCutout /
        # ConvertToProtrusion; Body.SurfaceArea became the sum of
        # Face.Area; Style.SetForegroundColor / ForegroundColor became
        # FaceStyle.SetDiffuse / GetDiffuse; Document.ApplyStyle and the
        # Material object's YoungsModulus / PoissonsRatio became
        # MatTable.ApplyMaterialToDoc and GetMaterialPropValueFromDoc; and
        # Models.AddFiniteRevolvedSurface became
        # doc.Constructions.RevolvedSurfaces.AddFinite.
        # Last round: Occurrence.Bodies became Body plus
        # GetSimplifiedBodies; SetColor / UseOccurrenceColor /
        # OccurrenceColor became Occurrence.FaceStyle;
        # Profile.OffsetProfile became Offset2d, which acts only on
        # selected geometry; Feature.Parents has no COM equivalent at all;
        # and SEInstallData is not registered on Solid Edge 2026, so
        # GetInstalledLanguage / GetInstalledVersion became
        # Application.Version / Name / AppDataFolder / RegistryPath.
        # CenterPoint, StartPoint and EndPoint were read as properties of a
        # 2D element throughout sketching.py. They are pure-[out] methods:
        # GetCenterPoint, GetStartPoint, GetEndPoint. sketch_rotate and
        # sketch_scale deleted the sketch and rebuilt nothing because of it.
        # Each still needs checking against a live install; they are reachable
        # only through interfaces the type libraries do not describe, or they
        # are wrong. Treat a failure here as a real bug until proven otherwise.
    }
)

pytestmark = pytest.mark.skipif(
    not DUMP_PATH.exists(),
    reason="reference/typelib_dump.json absent; run scripts/scrape_typelibs.py",
)


@pytest.fixture(scope="module")
def typelib_members() -> set[str]:
    """Every interface, method, property, enum, and enum member name."""
    data = json.loads(DUMP_PATH.read_text(encoding="utf-8"))
    names: set[str] = set()
    for payload in data["typelibs"].values():
        for enum_name, values in (payload.get("enums") or {}).items():
            names.add(enum_name)
            if isinstance(values, dict):
                names.update(values)
        for iface_name, iface in (payload.get("interfaces") or {}).items():
            names.add(iface_name)
            for key in ("methods", "properties"):
                names.update(iface.get(key) or {})
    return names


@pytest.fixture(scope="module")
def referenced_members() -> dict[str, list[str]]:
    """PascalCase attribute -> source locations that use it."""
    used: dict[str, list[str]] = defaultdict(list)
    for path in sorted(BACKENDS.rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if isinstance(node, ast.Attribute) and PASCAL_CASE.match(node.attr):
                used[node.attr].append(f"{path.relative_to(REPO_ROOT)}:{node.lineno}")
    return dict(used)


def test_scan_found_the_backends(typelib_members, referenced_members):
    assert len(typelib_members) > 15000, "type library index looks truncated"
    assert len(referenced_members) > 400, "attribute scan found suspiciously little"
    # Sanity: a member we know is real and one we know is not.
    assert "AddFiniteExtrudedProtrusion" in typelib_members
    assert "DefinitelyNotASolidEdgeMember" not in typelib_members


def test_no_new_unknown_com_members(typelib_members, referenced_members):
    """Any PascalCase attribute must exist in a type library, or be pinned."""
    unknown = {
        name: locations
        for name, locations in referenced_members.items()
        if name not in typelib_members and name not in UNVERIFIED
    }
    assert not unknown, "COM member names not found in any type library:\n  " + "\n  ".join(
        f"{name} at {locations[0]} ({len(locations)} use(s))"
        for name, locations in sorted(unknown.items())
    )


def test_unverified_list_has_no_stale_entries(typelib_members, referenced_members):
    """Remove names from UNVERIFIED once they are fixed or no longer used."""
    now_valid = sorted(n for n in UNVERIFIED if n in typelib_members)
    unused = sorted(n for n in UNVERIFIED if n not in referenced_members)
    problems = []
    if now_valid:
        problems.append(f"now present in a type library: {now_valid}")
    if unused:
        problems.append(f"no longer referenced by any backend: {unused}")
    assert not problems, "UNVERIFIED is stale; delete these entries:\n  " + "\n  ".join(problems)


@pytest.mark.parametrize(
    ("member", "reason"),
    [
        ("UnSuppress", "capital S; Unsuppress does not exist"),
        ("CenterLines", "capital L; Centerlines does not exist"),
        ("TextureFileName", "FaceStyle property; TextureName does not exist"),
        ("Occurrence1", "relation property; OccurrencePart1 does not exist"),
        ("Occurrence2", "relation property; OccurrencePart2 does not exist"),
        ("PutMatrix", "Occurrence method; SetMatrix does not exist"),
        ("SetSuppressComponent", "AssemblyDocument owns component suppression"),
        ("Dirty", "document modified flag; Document.Saved does not exist"),
        ("AddByStartAlongEnd", "3-point arc; AddByStartCenterEnd does not exist"),
        ("AddPartView", "part drawing view; only the assembly one was wired up"),
        ("GetActiveCommand", "a method; Application.ActiveCommand does not exist"),
        ("AddDistanceBetweenObjects", "dimensions measure between objects, not points"),
        ("ModelMembers", "ShowTangentEdges lives on the member, not the view"),
        ("AddAsFillet", "sketch fillet; AddByFillet does not exist"),
        ("AddAsChamfer", "sketch chamfer; AddByChamfer does not exist"),
        ("GetStartPoint", "line endpoints, for locating the corner to fillet"),
    ],
)
def test_corrected_names_are_real(typelib_members, member, reason):
    """Names an audit corrected, pinned so a revert is caught."""
    assert member in typelib_members, f"{member} missing ({reason})"


def test_feature_suppress_is_a_property_not_a_method():
    """Part features expose Suppress as a read/write VT_BOOL property.

    Calling ``feat.Suppress()`` invokes the getter and then calls a bool.
    """
    data = json.loads(DUMP_PATH.read_text(encoding="utf-8"))
    part = data["typelibs"]["Program/Part.tlb"]["interfaces"]
    for iface in ("ExtrudedProtrusion", "ExtrudedCutout", "Hole", "Round"):
        properties = part[iface].get("properties") or {}
        methods = part[iface].get("methods") or {}
        assert "Suppress" in properties, f"{iface}.Suppress should be a property"
        assert properties["Suppress"]["type"] == "VT_BOOL"
        assert "get/put" in properties["Suppress"]["access"]
        assert "Suppress" not in methods, f"{iface}.Suppress should not be a method"
