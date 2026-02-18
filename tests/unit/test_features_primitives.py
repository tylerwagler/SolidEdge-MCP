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
# BOX CUTOUT
# ============================================================================


class TestBoxCutout:
    def test_success(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        ref_planes = MagicMock()
        doc.RefPlanes = ref_planes
        box_features = MagicMock()
        models.BoxFeatures = box_features

        result = feature_mgr.create_box_cutout_by_two_points(0, 0, 0, 0.05, 0.05, 0.05)
        assert result["status"] == "created"
        assert result["type"] == "box_cutout"
        assert result["corner1"] == [0, 0, 0]
        assert result["corner2"] == [0.05, 0.05, 0.05]
        box_features.AddCutoutByTwoPoints.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_box_cutout_by_two_points(0, 0, 0, 1, 1, 1)
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# BOX CUTOUT BY CENTER
# ============================================================================


class TestBoxCutoutByCenter:
    def test_success(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        ref_planes = MagicMock()
        doc.RefPlanes = ref_planes
        box_features = MagicMock()
        models.BoxFeatures = box_features

        result = feature_mgr.create_box_cutout_by_center(0, 0, 0, 0.1, 0.1, 0.05)
        assert result["status"] == "created"
        assert result["type"] == "box_cutout"
        assert result["method"] == "by_center"
        assert result["center"] == [0, 0, 0]
        box_features.AddCutoutByCenter.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_box_cutout_by_center(0, 0, 0, 1, 1, 1)
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# BOX CUTOUT BY THREE POINTS
# ============================================================================


class TestBoxCutoutByThreePoints:
    def test_success(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        ref_planes = MagicMock()
        doc.RefPlanes = ref_planes
        box_features = MagicMock()
        models.BoxFeatures = box_features

        result = feature_mgr.create_box_cutout_by_three_points(0, 0, 0, 0.1, 0, 0, 0, 0.1, 0)
        assert result["status"] == "created"
        assert result["type"] == "box_cutout"
        assert result["method"] == "by_three_points"
        box_features.AddCutoutByThreePoints.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_box_cutout_by_three_points(0, 0, 0, 1, 0, 0, 0, 1, 0)
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# CYLINDER CUTOUT
# ============================================================================


class TestCylinderCutout:
    def test_success(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        ref_planes = MagicMock()
        doc.RefPlanes = ref_planes
        cyl_features = MagicMock()
        models.CylinderFeatures = cyl_features

        result = feature_mgr.create_cylinder_cutout(0, 0, 0, 0.01, 0.05)
        assert result["status"] == "created"
        assert result["type"] == "cylinder_cutout"
        assert result["radius"] == 0.01
        assert result["height"] == 0.05
        cyl_features.AddCutoutByCenterAndRadius.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_cylinder_cutout(0, 0, 0, 0.01, 0.05)
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# SPHERE CUTOUT
# ============================================================================


class TestSphereCutout:
    def test_success(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        ref_planes = MagicMock()
        doc.RefPlanes = ref_planes
        sph_features = MagicMock()
        models.SphereFeatures = sph_features

        result = feature_mgr.create_sphere_cutout(0, 0, 0, 0.02)
        assert result["status"] == "created"
        assert result["type"] == "sphere_cutout"
        assert result["radius"] == 0.02
        sph_features.AddCutoutByCenterAndRadius.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_sphere_cutout(0, 0, 0, 0.02)
        assert "error" in result
        assert "No base feature" in result["error"]
