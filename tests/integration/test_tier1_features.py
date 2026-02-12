"""
Integration tests for Tier 1 feature operations.

These tests require Solid Edge to be running on Windows.
Run with: uv run pytest tests/integration/test_tier1_features.py -v -m integration

Each test creates a new part, builds geometry, applies the feature, and verifies success.
"""

import pytest

# Mark all tests in this module as integration tests
pytestmark = pytest.mark.integration


@pytest.fixture(scope="module")
def managers():
    """Initialize and connect all managers for integration tests."""
    from solidedge_mcp.backends.connection import SolidEdgeConnection
    from solidedge_mcp.backends.documents import DocumentManager
    from solidedge_mcp.backends.features import FeatureManager
    from solidedge_mcp.backends.sketching import SketchManager

    conn = SolidEdgeConnection()
    result = conn.connect(start_if_needed=True)
    if "error" in result:
        pytest.skip(f"Cannot connect to Solid Edge: {result['error']}")

    doc_mgr = DocumentManager(conn)
    sketch_mgr = SketchManager(doc_mgr)
    feature_mgr = FeatureManager(doc_mgr, sketch_mgr)

    return doc_mgr, sketch_mgr, feature_mgr


@pytest.fixture
def new_part(managers):
    """Create a new part document for each test, close it after."""
    doc_mgr, _, _ = managers
    result = doc_mgr.create_part()
    assert "error" not in result, f"Failed to create part: {result}"
    yield
    doc_mgr.close_document(save=False)


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
        assert result["status"] == "created"
        assert result["type"] == "normal_cutout"


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
