"""
Unit tests for FeatureManager backend methods.

Uses unittest.mock to simulate COM objects so tests run without Solid Edge.
"""

from unittest.mock import MagicMock

import pytest


@pytest.fixture
def managers():
    """Create mock doc_manager and sketch_manager for FeatureManager."""
    doc_mgr = MagicMock()
    sketch_mgr = MagicMock()

    # Default: active document with models collection
    doc = MagicMock()
    doc_mgr.get_active_document.return_value = doc

    # Default: model exists (models.Count >= 1)
    models = MagicMock()
    models.Count = 1
    model = MagicMock()
    models.Item.return_value = model
    doc.Models = models

    # Default: body with faces and edges
    body = MagicMock()
    model.Body = body
    face = MagicMock()
    edge1 = MagicMock()
    edge2 = MagicMock()
    edges = MagicMock()
    edges.Count = 2
    edges.Item.side_effect = lambda i: [None, edge1, edge2][i]
    face.Edges = edges
    faces = MagicMock()
    faces.Count = 1
    faces.Item.return_value = face
    body.Faces.return_value = faces

    # Default: active profile exists
    profile = MagicMock()
    sketch_mgr.get_active_sketch.return_value = profile
    sketch_mgr.get_active_refaxis.return_value = None
    sketch_mgr.get_accumulated_profiles.return_value = []

    return doc_mgr, sketch_mgr, doc, models, model, profile


@pytest.fixture
def feature_mgr(managers):
    """Create FeatureManager with mocked dependencies."""
    from solidedge_mcp.backends.features import FeatureManager

    doc_mgr, sketch_mgr, *_ = managers
    return FeatureManager(doc_mgr, sketch_mgr)


# ============================================================================
# EXTRUDE
# ============================================================================


class TestCreateExtrude:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, profile = managers
        result = feature_mgr.create_extrude(0.05)
        assert result["status"] == "created"
        assert result["type"] == "extrude"
        assert result["distance"] == 0.05
        models.AddFiniteExtrudedProtrusion.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extrude(0.05)
        assert "error" in result
        assert "No active sketch" in result["error"]


# ============================================================================
# EXTRUDE THROUGH NEXT
# ============================================================================


class TestCreateExtrudeThroughNext:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        protrusions = MagicMock()
        model.ExtrudedProtrusions = protrusions

        result = feature_mgr.create_extrude_through_next()
        assert result["status"] == "created"
        assert result["type"] == "extrude_through_next"
        assert result["direction"] == "Normal"
        protrusions.AddThroughNext.assert_called_once_with(profile, 2)

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        protrusions = MagicMock()
        model.ExtrudedProtrusions = protrusions

        result = feature_mgr.create_extrude_through_next("Reverse")
        assert result["status"] == "created"
        protrusions.AddThroughNext.assert_called_once_with(profile, 1)

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extrude_through_next()
        assert "error" in result

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_extrude_through_next()
        assert "error" in result


# ============================================================================
# EXTRUDE FROM TO
# ============================================================================


class TestCreateExtrudeFromTo:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, profile = managers
        protrusions = MagicMock()
        model.ExtrudedProtrusions = protrusions
        ref_planes = MagicMock()
        ref_planes.Count = 4
        from_plane = MagicMock()
        to_plane = MagicMock()
        ref_planes.Item.side_effect = lambda i: {1: from_plane, 4: to_plane}.get(i)
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_extrude_from_to(1, 4)
        assert result["status"] == "created"
        assert result["type"] == "extrude_from_to"
        assert result["from_plane_index"] == 1
        assert result["to_plane_index"] == 4
        protrusions.AddFromTo.assert_called_once_with(profile, from_plane, to_plane)

    def test_invalid_from_plane(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_extrude_from_to(0, 2)
        assert "error" in result
        assert "from_plane_index" in result["error"]

    def test_invalid_to_plane(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_extrude_from_to(1, 10)
        assert "error" in result
        assert "to_plane_index" in result["error"]

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extrude_from_to(1, 2)
        assert "error" in result

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_extrude_from_to(1, 2)
        assert "error" in result


# ============================================================================
# CREATE EXTRUDE SYMMETRIC
# ============================================================================


class TestCreateExtrudeSymmetric:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, profile = managers
        result = feature_mgr.create_extrude_symmetric(0.05)
        assert result["status"] == "created"
        assert result["type"] == "extrude_symmetric"
        assert result["distance"] == 0.05

    def test_no_active_sketch(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None

        result = feature_mgr.create_extrude_symmetric(0.05)
        assert "error" in result

    def test_negative_distance(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, profile = managers
        result = feature_mgr.create_extrude_symmetric(-0.05)
        # Should still attempt (or error from COM)
        assert isinstance(result, dict)


# ============================================================================
# EXTRUDE THROUGH NEXT V2
# ============================================================================


class TestCreateExtrudeThroughNextV2:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_extrude_through_next_v2("Normal")
        assert result["status"] == "created"
        assert result["type"] == "extrude_through_next_v2"
        model.ExtrudedProtrusions.AddThroughNextMulti.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extrude_through_next_v2()
        assert "error" in result

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_extrude_through_next_v2()
        assert "error" in result


# ============================================================================
# EXTRUDE FROM TO V2
# ============================================================================


class TestCreateExtrudeFromToV2:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        result = feature_mgr.create_extrude_from_to_v2(1, 2)
        assert result["status"] == "created"
        assert result["type"] == "extrude_from_to_v2"
        model.ExtrudedProtrusions.AddFromToMulti.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extrude_from_to_v2(1, 2)
        assert "error" in result

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_extrude_from_to_v2(1, 2)
        assert "error" in result


# ============================================================================
# EXTRUDE BY KEYPOINT
# ============================================================================


class TestCreateExtrudeByKeypoint:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_extrude_by_keypoint("Normal")
        assert result["status"] == "created"
        assert result["type"] == "extrude_by_keypoint"
        model.ExtrudedProtrusions.AddFiniteByKeyPoint.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extrude_by_keypoint()
        assert "error" in result

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_extrude_by_keypoint()
        assert "error" in result


# ============================================================================
# SINGLE-PROFILE VARIANTS
# ============================================================================


class TestCreateExtrudeFromToSingle:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, profile = managers
        ref_planes = MagicMock()
        ref_planes.Count = 10
        ref_planes.Item.return_value = MagicMock()
        doc.RefPlanes = ref_planes
        protrusion = MagicMock()
        protrusion.Name = "Extrude1"
        model.ExtrudedProtrusions.AddFromTo.return_value = protrusion

        result = feature_mgr.create_extrude_from_to_single(4, 5)
        assert result["status"] == "created"
        assert result["type"] == "extrude_from_to_single"
        model.ExtrudedProtrusions.AddFromTo.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extrude_from_to_single(4, 5)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_extrude_from_to_single(4, 5)
        assert "error" in result
        assert "No base feature" in result["error"]


class TestCreateExtrudeThroughNextSingle:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        protrusion = MagicMock()
        protrusion.Name = "Extrude1"
        model.ExtrudedProtrusions.AddThroughNext.return_value = protrusion

        result = feature_mgr.create_extrude_through_next_single()
        assert result["status"] == "created"
        assert result["type"] == "extrude_through_next_single"
        model.ExtrudedProtrusions.AddThroughNext.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extrude_through_next_single()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        protrusion = MagicMock()
        protrusion.Name = "Extrude1"
        model.ExtrudedProtrusions.AddThroughNext.return_value = protrusion

        result = feature_mgr.create_extrude_through_next_single("Reverse")
        assert result["status"] == "created"
        assert result["direction"] == "Reverse"
