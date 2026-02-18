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
# LOFT WITH GUIDES
# ============================================================================


class TestCreateLoftWithGuides:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, models, model, _ = managers
        p1, p2, g1 = MagicMock(), MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2, g1]
        lp = MagicMock()
        model.LoftedProtrusions = lp

        result = feature_mgr.create_loft_with_guides(
            guide_profile_indices=[2], profile_indices=[0, 1]
        )
        assert result["status"] == "created"
        assert result["type"] == "loft_with_guides"
        assert result["num_profiles"] == 2
        assert result["num_guides"] == 1
        lp.AddSimple.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_no_guides(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        p1, p2 = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2]

        result = feature_mgr.create_loft_with_guides(guide_profile_indices=None)
        assert "error" in result
        assert "guide_profile_indices is required" in result["error"]

    def test_empty_guides(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        p1, p2 = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2]

        result = feature_mgr.create_loft_with_guides(guide_profile_indices=[])
        assert "error" in result
        assert "guide_profile_indices is required" in result["error"]

    def test_too_few_cross_sections(self, feature_mgr, managers):
        _, sketch_mgr, doc, models, model, _ = managers
        p1, g1 = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, g1]

        result = feature_mgr.create_loft_with_guides(
            guide_profile_indices=[1], profile_indices=[0]
        )
        assert "error" in result
        assert "at least 2 cross-section profiles" in result["error"]

    def test_auto_select_profiles(self, feature_mgr, managers):
        """When profile_indices is None, non-guide profiles are auto-selected."""
        _, sketch_mgr, doc, models, model, _ = managers
        p1, p2, p3, g1 = MagicMock(), MagicMock(), MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2, p3, g1]
        lp = MagicMock()
        model.LoftedProtrusions = lp

        result = feature_mgr.create_loft_with_guides(
            guide_profile_indices=[3]
        )
        assert result["status"] == "created"
        assert result["num_profiles"] == 3
        assert result["num_guides"] == 1


# ============================================================================
# BOUNDED SURFACE (BlueSurf)
# ============================================================================


class TestCreateBoundedSurface:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, _ = managers
        p1, p2 = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2]
        blue_surfs = MagicMock()
        model.BlueSurfs = blue_surfs

        result = feature_mgr.create_bounded_surface()
        assert result["status"] == "created"
        assert result["type"] == "bounded_surface"
        assert result["num_profiles"] == 2
        assert result["want_end_caps"] is True
        assert result["periodic"] is False
        blue_surfs.Add.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_periodic(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, _ = managers
        p1, p2, p3 = MagicMock(), MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2, p3]
        blue_surfs = MagicMock()
        model.BlueSurfs = blue_surfs

        result = feature_mgr.create_bounded_surface(want_end_caps=False, periodic=True)
        assert result["status"] == "created"
        assert result["want_end_caps"] is False
        assert result["periodic"] is True
        assert result["num_profiles"] == 3

    def test_too_few_profiles(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_accumulated_profiles.return_value = [MagicMock()]

        result = feature_mgr.create_bounded_surface()
        assert "error" in result
        assert "at least 2 profiles" in result["error"]

    def test_no_base_feature(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        sketch_mgr.get_accumulated_profiles.return_value = [MagicMock(), MagicMock()]
        models.Count = 0

        result = feature_mgr.create_bounded_surface()
        assert "error" in result
