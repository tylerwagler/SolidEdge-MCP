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
# REVOLVE (operation routing: Add -> protrusion, Cut -> RevolvedCutouts)
# ============================================================================


class TestCreateRevolve:
    def test_add_uses_finite_revolved_protrusion(self, feature_mgr, managers):
        import math

        from solidedge_mcp.backends.constants import DirectionConstants

        _, sketch_mgr, _, models, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        result = feature_mgr.create_revolve(90)
        assert result["status"] == "created"
        assert result["type"] == "revolve"
        assert result["angle"] == 90
        assert result["operation"] == "Add"
        models.AddFiniteRevolvedProtrusion.assert_called_once_with(
            1, (profile,), refaxis, DirectionConstants.igRight, math.radians(90)
        )
        model.RevolvedCutouts.AddFiniteMulti.assert_not_called()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        sketch_mgr.get_active_refaxis.return_value = MagicMock()
        result = feature_mgr.create_revolve(360)
        assert "error" in result
        assert "No active sketch" in result["error"]
        models.AddFiniteRevolvedProtrusion.assert_not_called()

    def test_no_refaxis(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        result = feature_mgr.create_revolve(360)
        assert "error" in result
        assert "axis of revolution" in result["error"]
        models.AddFiniteRevolvedProtrusion.assert_not_called()

    def test_cut_uses_revolved_cutout_collection(self, feature_mgr, managers):
        import math

        from solidedge_mcp.backends.constants import DirectionConstants

        _, sketch_mgr, _, models, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        result = feature_mgr.create_revolve(180, "Cut")
        assert result["status"] == "created"
        assert result["type"] == "revolve"
        assert result["operation"] == "Cut"
        assert result["angle"] == 180
        assert result["method"] == "RevolvedCutouts.AddFiniteMulti"
        model.RevolvedCutouts.AddFiniteMulti.assert_called_once_with(
            1, (profile,), refaxis, DirectionConstants.igRight, math.radians(180)
        )
        models.AddFiniteRevolvedProtrusion.assert_not_called()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_cut_without_base_feature_is_clear_error(self, feature_mgr, managers):
        _, sketch_mgr, _, models, model, _ = managers
        sketch_mgr.get_active_refaxis.return_value = MagicMock()
        models.Count = 0
        result = feature_mgr.create_revolve(360, "Cut")
        assert "error" in result
        assert "No base feature" in result["error"]
        model.RevolvedCutouts.AddFiniteMulti.assert_not_called()
        models.AddFiniteRevolvedProtrusion.assert_not_called()
        sketch_mgr.clear_accumulated_profiles.assert_not_called()

    def test_cut_without_refaxis_is_clear_error(self, feature_mgr, managers):
        _, _, _, models, model, _ = managers
        result = feature_mgr.create_revolve(360, "Cut")
        assert "error" in result
        assert "axis of revolution" in result["error"]
        model.RevolvedCutouts.AddFiniteMulti.assert_not_called()
        models.AddFiniteRevolvedProtrusion.assert_not_called()

    def test_cut_propagates_com_error(self, feature_mgr, managers):
        _, sketch_mgr, _, models, model, _ = managers
        sketch_mgr.get_active_refaxis.return_value = MagicMock()
        model.RevolvedCutouts.AddFiniteMulti.side_effect = RuntimeError("COM boom")
        result = feature_mgr.create_revolve(360, "Cut")
        assert "error" in result
        assert "COM boom" in result["error"]
        models.AddFiniteRevolvedProtrusion.assert_not_called()

    def test_intersect_is_unsupported(self, feature_mgr, managers):
        _, sketch_mgr, _, models, model, _ = managers
        sketch_mgr.get_active_refaxis.return_value = MagicMock()
        result = feature_mgr.create_revolve(360, "Intersect")
        assert "error" in result
        assert result["unsupported"] is True
        assert "Intersect" in result["error"]
        models.AddFiniteRevolvedProtrusion.assert_not_called()
        model.RevolvedCutouts.AddFiniteMulti.assert_not_called()
        sketch_mgr.clear_accumulated_profiles.assert_not_called()

    def test_unknown_operation_is_error(self, feature_mgr, managers):
        _, sketch_mgr, _, models, model, _ = managers
        sketch_mgr.get_active_refaxis.return_value = MagicMock()
        result = feature_mgr.create_revolve(360, "Subtract")
        assert "error" in result
        assert "Unknown operation" in result["error"]
        models.AddFiniteRevolvedProtrusion.assert_not_called()
        model.RevolvedCutouts.AddFiniteMulti.assert_not_called()


# ============================================================================
# REVOLVE BY KEYPOINT
# ============================================================================


class TestCreateRevolveByKeypoint:
    """AddFiniteByKeyPoint needs a KeyPoint object the server cannot select."""

    def test_reports_unsupported_without_calling_com(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        sketch_mgr.get_active_refaxis.return_value = MagicMock()
        result = feature_mgr.create_revolve_by_keypoint()
        assert result["unsupported"] is True
        assert "KeyPoint" in result["error"]
        model.RevolvedProtrusions.AddFiniteByKeyPoint.assert_not_called()

    def test_no_profile_still_unsupported(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_revolve_by_keypoint()
        assert result["unsupported"] is True
        model.RevolvedProtrusions.AddFiniteByKeyPoint.assert_not_called()

    def test_no_refaxis_still_unsupported(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_revolve_by_keypoint()
        assert result["unsupported"] is True
        model.RevolvedProtrusions.AddFiniteByKeyPoint.assert_not_called()


# ============================================================================
# REVOLVE FULL
# ============================================================================


class TestCreateRevolveFull:
    def test_success(self, feature_mgr, managers):
        import math

        _, sketch_mgr, _, _, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        result = feature_mgr.create_revolve_full(360.0, "None")
        assert result["status"] == "created"
        assert result["type"] == "revolve_full"
        # RevolvedProtrusions.Add(NumberOfProfiles, ProfileArray, RefAxis,
        #   ProfileSide, ExtentType1, ExtentSide1, FiniteAngle1,
        #   KeyPointOrTangentFace1, KeyPointFlags1, ExtentType2, ExtentSide2,
        #   FiniteAngle2, KeyPointOrTangentFace2, KeyPointFlags2)
        model.RevolvedProtrusions.Add.assert_called_once_with(
            1,
            (profile,),
            refaxis,
            2,  # igRight
            13,  # igFinite
            2,  # igRight
            math.radians(360.0),
            None,
            1,  # igTangentNormal
            44,  # igNone
            2,  # igRight
            0.0,
            None,
            1,  # igTangentNormal
        )

    def test_treatment_is_unsupported(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        sketch_mgr.get_active_refaxis.return_value = MagicMock()
        result = feature_mgr.create_revolve_full(360.0, "Draft")
        assert result["unsupported"] is True
        assert "treatment" in result["error"].lower()
        model.RevolvedProtrusions.Add.assert_not_called()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_revolve_full()
        assert "error" in result

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_revolve_full()
        assert "error" in result


# ============================================================================
# REVOLVE BY KEYPOINT SYNC
# ============================================================================


class TestCreateRevolveByKeypointSync:
    """AddFiniteByKeyPointSync needs a KeyPoint object the server cannot select."""

    def test_reports_unsupported_without_calling_com(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        sketch_mgr.get_active_refaxis.return_value = MagicMock()
        result = feature_mgr.create_revolve_by_keypoint_sync()
        assert result["unsupported"] is True
        assert "KeyPoint" in result["error"]
        model.RevolvedProtrusions.AddFiniteByKeyPointSync.assert_not_called()

    def test_no_profile_still_unsupported(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_revolve_by_keypoint_sync()
        assert result["unsupported"] is True
        model.RevolvedProtrusions.AddFiniteByKeyPointSync.assert_not_called()

    def test_no_refaxis_still_unsupported(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_revolve_by_keypoint_sync()
        assert result["unsupported"] is True
        model.RevolvedProtrusions.AddFiniteByKeyPointSync.assert_not_called()


# ============================================================================
# REVOLVE THIN WALL / SYNC (full-parameter Models.* overloads)
# ============================================================================


class TestCreateRevolveThinWall:
    def test_success(self, feature_mgr, managers):
        import math

        _, sketch_mgr, _, models, _, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis

        result = feature_mgr.create_revolve_thin_wall(90.0, 0.002)
        assert result["status"] == "created"
        assert result["type"] == "revolve_thin_wall"
        # AddRevolvedProtrusionWithThinWall: 19 required arguments
        models.AddRevolvedProtrusionWithThinWall.assert_called_once_with(
            1,
            (profile,),
            refaxis,
            2,  # ProfileSide igRight
            13,  # ExtentType1 igFinite
            2,  # ExtentSide1 igRight
            math.radians(90.0),
            None,  # KeyPointOrTangentFace1
            1,  # KeyPointFlags1 igTangentNormal
            44,  # ExtentType2 igNone
            2,  # ExtentSide2 igRight
            0.0,  # FiniteAngle2
            None,  # KeyPointOrTangentFace2
            1,  # KeyPointFlags2
            True,  # ThinWall
            False,  # AddEndCaps
            True,  # RemoveInsideMaterial
            0.002,  # Thickness
            4,  # ThicknessSide igInside
        )

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_revolve_thin_wall(90.0, 0.002)
        assert "error" in result
        models.AddRevolvedProtrusionWithThinWall.assert_not_called()


class TestCreateRevolveSync:
    def test_success(self, feature_mgr, managers):
        import math

        _, sketch_mgr, _, models, _, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis

        result = feature_mgr.create_revolve_sync(180.0)
        assert result["status"] == "created"
        assert result["type"] == "revolve_sync"
        # AddRevolvedProtrusionSync: 14 required arguments
        models.AddRevolvedProtrusionSync.assert_called_once_with(
            1,
            (profile,),
            refaxis,
            2,  # ProfileSide igRight
            13,  # ExtentType1 igFinite
            2,  # ExtentSide1 igRight
            math.radians(180.0),
            None,
            1,  # igTangentNormal
            44,  # ExtentType2 igNone
            2,
            0.0,
            None,
            1,
        )

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_revolve_sync(180.0)
        assert "error" in result
        models.AddRevolvedProtrusionSync.assert_not_called()
