"""
Integration tests for Tier 1 feature operations.

These require a running Solid Edge on Windows. Shared fixtures (connection,
managers, and the ``new_part`` scratch document) live in conftest.py, which
also guarantees no pre-existing document is ever closed.

Run with: uv run pytest -m integration
"""

import pytest

pytestmark = pytest.mark.integration


def _create_box(sketch_mgr, feature_mgr, size=0.1, height=0.05):
    """Helper: create a simple extruded box."""
    sketch_mgr.create_sketch("Top")
    sketch_mgr.draw_rectangle(0, 0, size, size)
    sketch_mgr.close_sketch()
    result = feature_mgr.create_extrude(height)
    assert result["status"] == "created", f"Extrude failed: {result}"


class TestRoundOnBox:
    def test_round_on_box(self, managers, new_part):
        _, sketch_mgr, feature_mgr = managers
        _create_box(sketch_mgr, feature_mgr)

        result = feature_mgr.create_round(0.002)
        assert result["status"] == "created"
        assert result["type"] == "round"
        assert result["edge_count"] > 0


class TestChamferOnBox:
    def test_chamfer_on_box(self, managers, new_part):
        _, sketch_mgr, feature_mgr = managers
        _create_box(sketch_mgr, feature_mgr)

        result = feature_mgr.create_chamfer(0.001)
        assert result["status"] == "created"
        assert result["type"] == "chamfer"
        assert result["edge_count"] > 0


class TestHoleInBox:
    def test_hole_in_box(self, managers, new_part):
        _, sketch_mgr, feature_mgr = managers
        _create_box(sketch_mgr, feature_mgr)

        result = feature_mgr.create_hole(0.05, 0.05, 0.01, 0.02)
        assert result["status"] == "created"
        assert result["type"] == "hole"


class TestExtrudedCutout:
    def test_finite_cutout(self, managers, new_part):
        _, sketch_mgr, feature_mgr = managers
        _create_box(sketch_mgr, feature_mgr, size=0.1, height=0.05)

        sketch_mgr.create_sketch("Top")
        sketch_mgr.draw_circle(0.05, 0.05, 0.01)
        sketch_mgr.close_sketch()

        result = feature_mgr.create_extruded_cutout(0.02)
        assert result["status"] == "created"
        assert result["type"] == "extruded_cutout"

    def test_through_all_cutout(self, managers, new_part):
        _, sketch_mgr, feature_mgr = managers
        _create_box(sketch_mgr, feature_mgr, size=0.1, height=0.05)

        sketch_mgr.create_sketch("Top")
        sketch_mgr.draw_circle(0.05, 0.05, 0.01)
        sketch_mgr.close_sketch()

        result = feature_mgr.create_extruded_cutout_through_all()
        assert result["status"] == "created"
        assert result["type"] == "extruded_cutout_through_all"


class TestRevolvedCutout:
    def test_revolved_cutout(self, managers, new_part):
        _, sketch_mgr, feature_mgr = managers

        # Create a cylinder to cut into
        sketch_mgr.create_sketch("Front")
        sketch_mgr.draw_rectangle(0.01, 0, 0.05, 0.1)
        sketch_mgr.set_axis_of_revolution(0, 0, 0, 0.1)
        sketch_mgr.close_sketch()
        result = feature_mgr.create_revolve(360)
        assert result["status"] == "created"

        # Create a revolved cutout (groove)
        sketch_mgr.create_sketch("Front")
        sketch_mgr.draw_rectangle(0.04, 0.04, 0.06, 0.06)
        sketch_mgr.set_axis_of_revolution(0, 0, 0, 0.1)
        sketch_mgr.close_sketch()
        result = feature_mgr.create_revolved_cutout(360)
        assert result["status"] == "created"
        assert result["type"] == "revolved_cutout"


class TestNormalCutout:
    def test_normal_cutout(self, managers, new_part):
        _, sketch_mgr, feature_mgr = managers
        _create_box(sketch_mgr, feature_mgr, size=0.1, height=0.05)

        sketch_mgr.create_sketch("Top")
        sketch_mgr.draw_circle(0.05, 0.05, 0.01)
        sketch_mgr.close_sketch()

        result = feature_mgr.create_normal_cutout(0.02)

        # A normal cutout to a finite depth is a sheet metal feature. On an
        # ordinary part the COM call is accepted and removes no material, so
        # the document type is checked before COM is touched at all. Verified
        # against SE 2026, where the same call cuts on a sheet metal document
        # and the through_all variant cuts on a part.
        assert "error" in result, f"expected a refusal on a part, got {result}"
        assert "sheet metal feature" in result["error"].lower()


class TestLoftedCutout:
    def test_lofted_cutout(self, managers, new_part):
        _, sketch_mgr, feature_mgr = managers

        # Create a large box as base feature
        _create_box(sketch_mgr, feature_mgr, size=0.2, height=0.1)

        # Create two cross-section profiles on different planes for the lofted cutout
        # Profile 1: circle on top plane
        sketch_mgr.create_sketch("Top")
        sketch_mgr.draw_circle(0.1, 0.1, 0.03)
        sketch_mgr.close_sketch()

        # Profile 2: circle on offset plane
        result = feature_mgr.create_ref_plane_by_offset(1, 0.05)
        assert result["status"] == "created"
        plane_idx = result["new_plane_index"]

        sketch_mgr.create_sketch_on_plane_index(plane_idx)
        sketch_mgr.draw_circle(0.1, 0.1, 0.01)
        sketch_mgr.close_sketch()

        result = feature_mgr.create_lofted_cutout()
        assert result["status"] == "created"
        assert result["type"] == "lofted_cutout"
        assert result["num_profiles"] == 2
