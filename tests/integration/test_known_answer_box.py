"""A box of known size: every measurement has a value that can be computed by hand.

Nothing else in this suite checks that a returned number is *right*. Every
other gate answers "did the write take" or "did the face count change", and a
measurement wrong by a unit factor passes all of them -- DraftPrintUtility.
PaperWidth being handed metres when it wanted millimetres was exactly that,
and it was found by hand.

The box is 0.08 x 0.048 x 0.03 m, so:

    volume        = 1.152e-4 m^3
    surface area  = 2(0.08*0.048 + 0.08*0.03 + 0.048*0.03) = 0.01536 m^2
    face areas    = {0.00384 x2, 0.0024 x2, 0.00144 x2}
    body diagonal = sqrt(0.08^2 + 0.048^2 + 0.03^2) = 0.098 m exactly

Tolerances are float noise only (rel=1e-9). If a path cannot hit an exact
volume at that tolerance, that is a finding about the path, not about the
test. Face and edge ordering is not contractual, so multisets are compared
sorted, never index-to-value.

The extrude's sign along Z is not assumed: the corners are derived from the
measured bounding box, and the centre of gravity is required to equal that
box's centre.

Run with: uv run pytest -m integration
"""

from __future__ import annotations

import math

import pytest

pytestmark = pytest.mark.integration

X, Y, Z = 0.08, 0.048, 0.03
VOLUME = X * Y * Z
AREA = 2 * (X * Y + X * Z + Y * Z)
FACE_AREAS = sorted([X * Y, X * Y, X * Z, X * Z, Y * Z, Y * Z])
DIAGONAL = math.sqrt(X * X + Y * Y + Z * Z)
EXACT = dict(rel=1e-9, abs=1e-12)


def _box(sketch_mgr, feature_mgr):
    sketch_mgr.create_sketch("Top")
    sketch_mgr.draw_rectangle(0, 0, X, Y)
    sketch_mgr.close_sketch()
    result = feature_mgr.create_extrude(Z)
    assert result["status"] == "created", f"Extrude failed: {result}"


@pytest.fixture
def query_mgr(managers):
    from solidedge_mcp.backends.query import QueryManager

    doc_mgr, _, _ = managers
    return QueryManager(doc_mgr)


@pytest.fixture
def box(managers, query_mgr, new_part):
    _, sketch_mgr, feature_mgr = managers
    _box(sketch_mgr, feature_mgr)
    return query_mgr


class TestExactValues:
    def test_volume(self, box):
        r = box.get_volume()
        assert "error" not in r, r
        assert r["volume"] == pytest.approx(VOLUME, **EXACT)
        assert r["volume_mm3"] == pytest.approx(VOLUME * 1e9, **EXACT)

    def test_surface_area(self, box):
        r = box.get_surface_area()
        assert "error" not in r, r
        assert r["surface_area"] == pytest.approx(AREA, **EXACT)
        assert r["face_count"] == 6

    def test_every_face_area_as_a_multiset(self, box):
        r = box.get_body_faces()
        assert "error" not in r, r
        areas = sorted(item["area"] for item in r["items"])
        assert len(areas) == 6
        assert areas == pytest.approx(FACE_AREAS, **EXACT)

    def test_face_area_by_index_matches_the_listing(self, box):
        listing = {item["index"]: item["area"] for item in box.get_body_faces()["items"]}
        for index, area in listing.items():
            single = box.get_face_area(index)
            assert single["area"] == pytest.approx(area, **EXACT), index
            assert single["area_mm2"] == pytest.approx(area * 1e6, **EXACT), index

    def test_counts(self, box):
        assert box.get_face_count()["face_count"] == 6
        edges = box.get_edge_count()
        assert edges["total_edge_references"] == 24, "12 edges, each on 2 faces"

    def test_bounding_box_dimensions(self, box):
        r = box.get_bounding_box()
        assert "error" not in r, r
        assert r["dimensions"]["x"] == pytest.approx(X, **EXACT)
        assert r["dimensions"]["y"] == pytest.approx(Y, **EXACT)
        assert r["dimensions"]["z"] == pytest.approx(Z, **EXACT)

    def test_the_eight_corners(self, box):
        bb = box.get_bounding_box()
        lo, hi = bb["min"], bb["max"]
        expected = sorted(
            (x, y, z) for x in (lo[0], hi[0]) for y in (lo[1], hi[1]) for z in (lo[2], hi[2])
        )
        r = box.get_body_vertices()
        assert "error" not in r, r
        points = sorted({tuple(round(c, 9) for c in item["point"]) for item in r["items"]})
        assert len(points) == 8, points
        for got, want in zip(points, expected, strict=True):
            assert got == pytest.approx(want, **EXACT)

    def test_centre_of_gravity_is_the_box_centre(self, box):
        bb = box.get_bounding_box()
        centre = [(lo + hi) / 2 for lo, hi in zip(bb["min"], bb["max"], strict=True)]
        r = box.get_center_of_gravity()
        assert "error" not in r, r
        assert r["center_of_gravity"] == pytest.approx(centre, **EXACT)
        assert r["center_of_gravity_mm"] == pytest.approx([c * 1000 for c in centre], **EXACT)
        assert abs(r["center_of_gravity"][2]) == pytest.approx(Z / 2, **EXACT)

    def test_mass_at_a_given_density(self, box):
        for density in (1000.0, 7850.0):
            r = box.get_mass_properties(density=density)
            assert "error" not in r, r
            assert r["mass"] == pytest.approx(VOLUME * density, **EXACT), density
            assert r["density"] == density

    def test_pure_python_measurements(self, box):
        d = box.measure_distance(0, 0, 0, X, Y, Z)
        assert d["distance"] == pytest.approx(DIAGONAL, **EXACT)
        assert d["distance"] == pytest.approx(0.098, **EXACT)

        a = box.measure_angle(1, 0, 0, 0, 0, 0, 0, 1, 0)
        assert a["angle_degrees"] == pytest.approx(90.0, **EXACT)
        assert a["angle_radians"] == pytest.approx(math.pi / 2, **EXACT)


class TestInvariantsThatNeedNoReferenceValue:
    """Cross-checks between paths that report the same quantity.

    These would have caught the PaperWidth class of bug on their own: two
    code paths for one number that disagree, or a scaled twin that is not
    the stated multiple of its base.
    """

    def test_volume_agrees_across_paths(self, box):
        direct = box.get_volume()["volume"]
        via_mass = box.get_mass_properties()["volume"]
        assert direct == pytest.approx(via_mass, **EXACT)

    def test_surface_area_agrees_across_paths(self, box):
        direct = box.get_surface_area()["surface_area"]
        via_mass = box.get_mass_properties()["surface_area"]
        summed = sum(item["area"] for item in box.get_body_faces()["items"])
        assert direct == pytest.approx(via_mass, **EXACT)
        assert direct == pytest.approx(summed, **EXACT)

    def test_mass_is_volume_times_density(self, box):
        r = box.set_material_density(2700.0)
        assert "error" not in r, r
        assert r["mass"] == pytest.approx(r["volume"] * 2700.0, **EXACT)
        assert r["density"] == 2700.0

    def test_centre_of_gravity_equals_centre_of_volume(self, box):
        r = box.get_mass_properties()
        assert r["center_of_gravity"] == pytest.approx(r["center_of_volume"], **EXACT)

    def test_bounding_box_agrees_with_spatial_context(self, box):
        bb = box.get_bounding_box()["dimensions"]
        sc = box.get_spatial_context()
        assert "error" not in sc, sc
        dims = sc["bounding_box"]["dimensions"]
        assert [dims["x"], dims["y"], dims["z"]] == pytest.approx(
            [bb["x"], bb["y"], bb["z"]], **EXACT
        )
