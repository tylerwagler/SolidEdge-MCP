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
# EXTRUDED SURFACE
# ============================================================================


class TestCreateExtrudedSurface:
    def test_success_normal(self):
        from solidedge_mcp.backends.features import FeatureManager

        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)

        doc = MagicMock()
        profile = MagicMock()
        dm.get_active_document.return_value = doc
        sm.get_active_sketch.return_value = profile

        surface = MagicMock()
        doc.Constructions.ExtrudedSurfaces.Add.return_value = surface

        result = fm.create_extruded_surface(0.05, direction="Normal", end_caps=True)
        assert result["status"] == "created"
        assert result["type"] == "extruded_surface"
        assert result["distance"] == 0.05
        assert result["direction"] == "Normal"
        doc.Constructions.ExtrudedSurfaces.Add.assert_called_once()

    def test_success_symmetric(self):
        from solidedge_mcp.backends.features import FeatureManager

        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)

        doc = MagicMock()
        profile = MagicMock()
        dm.get_active_document.return_value = doc
        sm.get_active_sketch.return_value = profile

        result = fm.create_extruded_surface(0.03, direction="Symmetric")
        assert result["status"] == "created"
        assert result["direction"] == "Symmetric"

    def test_no_profile(self):
        from solidedge_mcp.backends.features import FeatureManager

        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)

        dm.get_active_document.return_value = MagicMock()
        sm.get_active_sketch.return_value = None

        result = fm.create_extruded_surface(0.05)
        assert "error" in result


# ============================================================================
# TIER 2: REVOLVED SURFACE
# ============================================================================


class TestRevolvedSurface:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        result = feature_mgr.create_revolved_surface(360)
        assert result["status"] == "created"
        assert result["type"] == "revolved_surface"
        assert result["angle_degrees"] == 360
        model.RevolvedSurfaces.AddFinite.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_revolved_surface()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_revolved_surface()
        assert "error" in result
        assert "axis" in result["error"].lower()

    def test_partial_angle(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        result = feature_mgr.create_revolved_surface(180, want_end_caps=True)
        assert result["status"] == "created"
        assert result["angle_degrees"] == 180
        assert result["want_end_caps"] is True


# ============================================================================
# TIER 2: LOFTED SURFACE
# ============================================================================


class TestLoftedSurface:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        p1, p2 = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2]
        result = feature_mgr.create_lofted_surface()
        assert result["status"] == "created"
        assert result["type"] == "lofted_surface"
        assert result["num_profiles"] == 2
        model.LoftedSurfaces.Add.assert_called_once()

    def test_too_few_profiles(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_accumulated_profiles.return_value = [MagicMock()]
        result = feature_mgr.create_lofted_surface()
        assert "error" in result
        assert "at least 2" in result["error"]

    def test_no_base_feature(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        models.Count = 0
        p1, p2 = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2]
        result = feature_mgr.create_lofted_surface()
        assert "error" in result
        assert "base feature" in result["error"].lower()


# ============================================================================
# TIER 2: SWEPT SURFACE
# ============================================================================


class TestSweptSurface:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        path, cs = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [path, cs]
        result = feature_mgr.create_swept_surface()
        assert result["status"] == "created"
        assert result["type"] == "swept_surface"
        assert result["num_cross_sections"] == 1
        model.SweptSurfaces.Add.assert_called_once()

    def test_too_few_profiles(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_accumulated_profiles.return_value = [MagicMock()]
        result = feature_mgr.create_swept_surface()
        assert "error" in result
        assert "at least 2" in result["error"]

    def test_no_base_feature(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_swept_surface()
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_with_end_caps(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        path, cs = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [path, cs]
        result = feature_mgr.create_swept_surface(want_end_caps=True)
        assert result["status"] == "created"
        assert result["want_end_caps"] is True


# ============================================================================
# BATCH 4: EXTRUDED SURFACE FROM-TO
# ============================================================================


class TestCreateExtrudedSurfaceFromTo:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, _, profile = managers
        ref_planes = MagicMock()
        ref_planes.Count = 4
        doc.RefPlanes = ref_planes
        ref_planes.Item.return_value = MagicMock()
        constructions = MagicMock()
        doc.Constructions = constructions
        extruded_surfaces = MagicMock()
        constructions.ExtrudedSurfaces = extruded_surfaces

        result = feature_mgr.create_extruded_surface_from_to(1, 4)
        assert result["status"] == "created"
        assert result["type"] == "extruded_surface_from_to"
        assert result["from_plane_index"] == 1
        assert result["to_plane_index"] == 4
        extruded_surfaces.AddFromTo.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_no_active_sketch(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None

        result = feature_mgr.create_extruded_surface_from_to(1, 4)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_invalid_from_plane(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_extruded_surface_from_to(10, 1)
        assert "error" in result
        assert "Invalid from_plane_index" in result["error"]


# ============================================================================
# BATCH 4: EXTRUDED SURFACE BY KEYPOINT
# ============================================================================


class TestCreateExtrudedSurfaceByKeypoint:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, _, profile = managers
        constructions = MagicMock()
        doc.Constructions = constructions
        extruded_surfaces = MagicMock()
        constructions.ExtrudedSurfaces = extruded_surfaces

        result = feature_mgr.create_extruded_surface_by_keypoint("End")
        assert result["status"] == "created"
        assert result["type"] == "extruded_surface_by_keypoint"
        assert result["keypoint_type"] == "End"
        extruded_surfaces.AddFiniteByKeyPoint.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_no_active_sketch(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None

        result = feature_mgr.create_extruded_surface_by_keypoint("End")
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_start_keypoint(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        constructions = MagicMock()
        doc.Constructions = constructions
        extruded_surfaces = MagicMock()
        constructions.ExtrudedSurfaces = extruded_surfaces

        result = feature_mgr.create_extruded_surface_by_keypoint("Start")
        assert result["status"] == "created"
        assert result["keypoint_type"] == "Start"


# ============================================================================
# BATCH 4: EXTRUDED SURFACE BY CURVES
# ============================================================================


class TestCreateExtrudedSurfaceByCurves:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, _, profile = managers
        constructions = MagicMock()
        doc.Constructions = constructions
        extruded_surfaces = MagicMock()
        constructions.ExtrudedSurfaces = extruded_surfaces

        result = feature_mgr.create_extruded_surface_by_curves(0.05)
        assert result["status"] == "created"
        assert result["type"] == "extruded_surface_by_curves"
        assert result["distance"] == 0.05
        extruded_surfaces.AddByCurves.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_no_active_sketch(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None

        result = feature_mgr.create_extruded_surface_by_curves(0.05)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_symmetric_direction(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        constructions = MagicMock()
        doc.Constructions = constructions
        extruded_surfaces = MagicMock()
        constructions.ExtrudedSurfaces = extruded_surfaces

        result = feature_mgr.create_extruded_surface_by_curves(0.05, "Symmetric")
        assert result["status"] == "created"
        assert result["direction"] == "Symmetric"


# ============================================================================
# BATCH 4: REVOLVED SURFACE SYNC
# ============================================================================


class TestCreateRevolvedSurfaceSync:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, models, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        rev_surfaces = MagicMock()
        model.RevolvedSurfaces = rev_surfaces

        result = feature_mgr.create_revolved_surface_sync(360.0)
        assert result["status"] == "created"
        assert result["type"] == "revolved_surface_sync"
        assert result["angle_degrees"] == 360.0
        rev_surfaces.AddFiniteSync.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_no_active_sketch(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None

        result = feature_mgr.create_revolved_surface_sync(360.0)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_axis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None

        result = feature_mgr.create_revolved_surface_sync(360.0)
        assert "error" in result
        assert "axis" in result["error"].lower()


# ============================================================================
# BATCH 4: REVOLVED SURFACE BY KEYPOINT
# ============================================================================


class TestCreateRevolvedSurfaceByKeypoint:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, models, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        rev_surfaces = MagicMock()
        model.RevolvedSurfaces = rev_surfaces

        result = feature_mgr.create_revolved_surface_by_keypoint("End")
        assert result["status"] == "created"
        assert result["type"] == "revolved_surface_by_keypoint"
        assert result["keypoint_type"] == "End"
        rev_surfaces.AddFiniteByKeyPoint.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_no_active_sketch(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None

        result = feature_mgr.create_revolved_surface_by_keypoint("End")
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_axis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None

        result = feature_mgr.create_revolved_surface_by_keypoint("End")
        assert "error" in result
        assert "axis" in result["error"].lower()


# ============================================================================
# BATCH 4: LOFTED SURFACE V2
# ============================================================================


class TestCreateLoftedSurfaceV2:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, models, model, _ = managers
        p1, p2 = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2]
        loft_surfaces = MagicMock()
        model.LoftedSurfaces = loft_surfaces

        result = feature_mgr.create_lofted_surface_v2()
        assert result["status"] == "created"
        assert result["type"] == "lofted_surface_v2"
        assert result["num_profiles"] == 2
        loft_surfaces.Add2.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_too_few_profiles(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_accumulated_profiles.return_value = [MagicMock()]

        result = feature_mgr.create_lofted_surface_v2()
        assert "error" in result
        assert "at least 2 profiles" in result["error"]

    def test_no_base_feature(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        sketch_mgr.get_accumulated_profiles.return_value = [MagicMock(), MagicMock()]
        models.Count = 0

        result = feature_mgr.create_lofted_surface_v2()
        assert "error" in result
        assert "base feature" in result["error"]


# ============================================================================
# BATCH 4: SWEPT SURFACE EX
# ============================================================================


class TestCreateSweptSurfaceEx:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, models, model, _ = managers
        path, cs = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [path, cs]
        swept_surfaces = MagicMock()
        model.SweptSurfaces = swept_surfaces

        result = feature_mgr.create_swept_surface_ex()
        assert result["status"] == "created"
        assert result["type"] == "swept_surface_ex"
        assert result["num_cross_sections"] == 1
        swept_surfaces.AddEx.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_too_few_profiles(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_accumulated_profiles.return_value = [MagicMock()]

        result = feature_mgr.create_swept_surface_ex()
        assert "error" in result
        assert "at least 2 profiles" in result["error"]

    def test_no_base_feature(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        models.Count = 0

        result = feature_mgr.create_swept_surface_ex()
        assert "error" in result
        assert "base feature" in result["error"]


# ============================================================================
# BATCH 4: EXTRUDED SURFACE FULL
# ============================================================================


class TestCreateExtrudedSurfaceFull:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, _, profile = managers
        constructions = MagicMock()
        doc.Constructions = constructions
        extruded_surfaces = MagicMock()
        constructions.ExtrudedSurfaces = extruded_surfaces

        result = feature_mgr.create_extruded_surface_full(0.05)
        assert result["status"] == "created"
        assert result["type"] == "extruded_surface_full"
        assert result["distance"] == 0.05
        assert result["treatment_type"] == "None"
        extruded_surfaces.Add.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_no_active_sketch(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None

        result = feature_mgr.create_extruded_surface_full(0.05)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_with_draft_treatment(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        constructions = MagicMock()
        doc.Constructions = constructions
        extruded_surfaces = MagicMock()
        constructions.ExtrudedSurfaces = extruded_surfaces

        result = feature_mgr.create_extruded_surface_full(
            0.05, treatment_type="Draft", draft_angle=5.0
        )
        assert result["status"] == "created"
        assert result["treatment_type"] == "Draft"
        assert result["draft_angle"] == 5.0


# ============================================================================
# REVOLVED SURFACE FULL
# ============================================================================


class TestCreateRevolvedSurfaceFull:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        surface = MagicMock()
        surface.Name = "RevSurfFull1"
        model.RevolvedSurfaces.Add.return_value = surface

        result = feature_mgr.create_revolved_surface_full(180.0)
        assert result["status"] == "created"
        assert result["type"] == "revolved_surface_full"
        model.RevolvedSurfaces.Add.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_revolved_surface_full()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_revolved_surface_full()
        assert "error" in result
        assert "axis" in result["error"].lower()


class TestCreateRevolvedSurfaceFullSync:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        surface = MagicMock()
        surface.Name = "RevSurfFullSync1"
        model.RevolvedSurfaces.AddSync.return_value = surface

        result = feature_mgr.create_revolved_surface_full_sync(180.0)
        assert result["status"] == "created"
        assert result["type"] == "revolved_surface_full_sync"
        model.RevolvedSurfaces.AddSync.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_revolved_surface_full_sync()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_revolved_surface_full_sync()
        assert "error" in result
        assert "axis" in result["error"].lower()
