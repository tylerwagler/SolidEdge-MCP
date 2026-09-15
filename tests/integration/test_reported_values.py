"""What a tool reports must be true of the document Solid Edge is holding.

Every bug these cover was found by driving a live Solid Edge and reading the
value back, and none of them could be seen from a mocked test: the call
succeeded, the result dict looked right, and the document disagreed.

Run with: uv run pytest -m integration
"""

from __future__ import annotations

import pytest

pytestmark = pytest.mark.integration


def _box(sketch_mgr, feature_mgr, size=0.08, height=0.03):
    sketch_mgr.create_sketch("Top")
    sketch_mgr.draw_rectangle(0, 0, size, size * 0.6)
    sketch_mgr.close_sketch()
    result = feature_mgr.create_extrude(height)
    assert result["status"] == "created", f"Extrude failed: {result}"


def _cutout(sketch_mgr, feature_mgr, cx, cy, radius=0.004, depth=0.01):
    sketch_mgr.create_sketch("Top")
    sketch_mgr.draw_circle(cx, cy, radius)
    sketch_mgr.close_sketch()
    result = feature_mgr.create_extruded_cutout(depth)
    assert result["status"] == "created", f"Cutout failed: {result}"


@pytest.fixture
def query_mgr(managers):
    """A QueryManager sharing the document manager the other fixtures use."""
    from solidedge_mcp.backends.query import QueryManager

    doc_mgr, _, _ = managers
    return QueryManager(doc_mgr)


class TestTheReportedIndexAddressesTheFeatureItNames:
    """An index a caller is handed has to be the index a tool will take.

    get_feature_status numbered features by their spot in
    DesignEdgebarFeatures, which carries the three reference planes as well.
    Every index it reported was three too high: asked about the base
    extrusion it answered 3, and passing 3 to a tool hit the third cutout and
    reported success. With a delete, that removes the wrong geometry silently.
    """

    def test_status_index_round_trips_through_the_tool_lookup(self, managers, query_mgr, new_part):
        _, sketch_mgr, feature_mgr = managers
        _box(sketch_mgr, feature_mgr)
        for cx, cy in ((0.02, 0.02), (0.06, 0.02), (0.02, 0.04), (0.06, 0.04)):
            _cutout(sketch_mgr, feature_mgr, cx, cy)

        listed = feature_mgr.list_features()["features"]
        assert len(listed) >= 5, f"expected the box and four cutouts: {listed}"

        for entry in listed:
            name = entry["name"]
            status = query_mgr.get_feature_status(name)
            assert "error" not in status, status

            assert status["index"] == entry["index"], (
                f"status put {name!r} at index {status['index']}, "
                f"list_features puts it at {entry['index']}"
            )

            resolved, err = feature_mgr._get_feature_by_index(status["index"])
            assert err is None, err
            assert resolved.Name == name, (
                f"index {status['index']} reported for {name!r} actually "
                f"addresses {resolved.Name!r}"
            )

    def test_the_tree_position_is_reported_separately(self, managers, query_mgr, new_part):
        """The Pathfinder spot is still available, under a name that says so."""
        _, sketch_mgr, feature_mgr = managers
        _box(sketch_mgr, feature_mgr)

        name = feature_mgr.list_features()["features"][0]["name"]
        status = query_mgr.get_feature_status(name)

        assert status["index"] == 0
        assert status["tree_position"] > status["index"], (
            "the reference planes sit ahead of the first feature in the tree, "
            f"so its tree position should exceed 0: {status}"
        )


class TestAnEmptyBodySaysWhyItIsEmpty:
    """Suppressing every feature empties the body as a mode switch does.

    Both leave Body unreadable and they want opposite advice. Every geometry
    query used to name the ordered/synchronous mode whatever the cause, which
    sent a caller whose features were suppressed to a setting that could not
    help them.
    """

    def test_suppressing_everything_names_suppression(self, managers, query_mgr, new_part):
        _, sketch_mgr, feature_mgr = managers
        _box(sketch_mgr, feature_mgr)

        assert "error" not in query_mgr.get_surface_area()

        for entry in feature_mgr.list_features()["features"]:
            result = feature_mgr.feature_suppress(entry["index"])
            assert "error" not in result, result

        after = query_mgr.get_surface_area()
        assert "error" in after, f"a part with nothing built should not measure: {after}"
        assert "suppressed" in after["error"], after["error"]
        assert "synchronous" not in after["error"], (
            "naming the modelling mode here sends the caller nowhere useful: " + after["error"]
        )

    def test_unsuppressing_brings_the_measurement_back(self, managers, query_mgr, new_part):
        _, sketch_mgr, feature_mgr = managers
        _box(sketch_mgr, feature_mgr)
        before = query_mgr.get_surface_area()

        for entry in feature_mgr.list_features()["features"]:
            feature_mgr.feature_suppress(entry["index"])
        for entry in feature_mgr.list_features()["features"]:
            feature_mgr.feature_unsuppress(entry["index"])

        after = query_mgr.get_surface_area()
        assert "error" not in after, after
        assert after == pytest.approx(before, rel=1e-6) or set(after) == set(before)


class TestALoftPairsProfilesByAPointOnEachOfThem:
    """A loft's Origins array must hold a point lying on its own profile.

    A hardcoded (0, 0) makes Solid Edge build nothing at all -- no error, no
    geometry -- unless every profile happens to cross the sketch origin. Both
    profiles here are placed away from it, which is the case that failed.
    """

    def test_a_loft_between_off_origin_profiles_builds(self, managers, query_mgr, new_part):
        doc_mgr, sketch_mgr, feature_mgr = managers

        sketch_mgr.create_sketch("Top")
        sketch_mgr.draw_rectangle(0.02, 0.02, 0.06, 0.05)
        sketch_mgr.close_sketch()

        plane = feature_mgr.create_ref_plane_by_offset(1, 0.04)
        assert "error" not in plane, plane

        sketch_mgr.create_sketch_on_plane_index(4)
        sketch_mgr.draw_rectangle(0.03, 0.025, 0.05, 0.045)
        sketch_mgr.close_sketch()

        result = feature_mgr.create_loft()
        assert "error" not in result, f"the loft built nothing: {result}"
        assert result["status"] == "created", result

        faces = query_mgr.get_surface_area()
        assert "error" not in faces, faces
