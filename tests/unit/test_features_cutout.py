"""
Unit tests for FeatureManager backend methods.

Uses unittest.mock to simulate COM objects so tests run without Solid Edge.
"""

from unittest.mock import MagicMock

import pytest

from solidedge_mcp.backends.constants import DocumentTypeConstants

IG_PART_DOCUMENT = DocumentTypeConstants.igPartDocument
IG_SHEET_METAL_DOCUMENT = DocumentTypeConstants.igSheetMetalDocument


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
# EXTRUDED CUTOUT
# ============================================================================


class TestExtrudedCutout:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, profile = managers
        result = feature_mgr.create_extruded_cutout(0.01)
        assert result["status"] == "created"
        assert result["type"] == "extruded_cutout"
        model.ExtrudedCutouts.AddFiniteMulti.assert_called_once_with(
            1,
            (profile,),
            2,
            0.01,  # igRight=2
        )
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extruded_cutout(0.01)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_extruded_cutout(0.01)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        result = feature_mgr.create_extruded_cutout(0.01, direction="Reverse")
        assert result["direction"] == "Reverse"
        model.ExtrudedCutouts.AddFiniteMulti.assert_called_once_with(
            1,
            (profile,),
            1,
            0.01,  # igLeft=1
        )


# ============================================================================
# EXTRUDED CUTOUT THROUGH ALL
# ============================================================================


class TestExtrudedCutoutThroughAll:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, profile = managers
        result = feature_mgr.create_extruded_cutout_through_all()
        assert result["status"] == "created"
        assert result["type"] == "extruded_cutout_through_all"
        model.ExtrudedCutouts.AddThroughAllMulti.assert_called_once_with(
            1,
            (profile,),
            2,  # igRight=2
        )
        sketch_mgr.clear_accumulated_profiles.assert_called_once()


# ============================================================================
# REVOLVED CUTOUT
# ============================================================================


class TestRevolvedCutout:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis

        result = feature_mgr.create_revolved_cutout(360)
        assert result["status"] == "created"
        assert result["type"] == "revolved_cutout"
        model.RevolvedCutouts.AddFiniteMulti.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_no_axis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_revolved_cutout(360)
        assert "error" in result
        assert "axis" in result["error"].lower()


# ============================================================================
# NORMAL CUTOUT
# ============================================================================


class TestNormalCutout:
    """A finite normal cutout is a sheet metal feature.

    Verified on Solid Edge 2026: it cuts on a sheet metal document and
    removes nothing from an ordinary part, while the through_all and
    through_next variants work on both.
    """

    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, profile = managers
        doc.Type = IG_SHEET_METAL_DOCUMENT
        result = feature_mgr.create_normal_cutout(0.005)
        assert result["status"] == "created"
        assert result["type"] == "normal_cutout"
        assert result["distance"] == 0.005
        # Method (igSMFaceCutout=205) is required; without it Solid Edge
        # raises 0x8002000F Parameter not optional.
        model.NormalCutouts.AddFiniteMulti.assert_called_once_with(
            1,
            (profile,),
            2,  # igRight
            0.005,
            205,  # igSMFaceCutout
        )
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_normal_cutout(0.005)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_normal_cutout(0.005)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_refuses_an_ordinary_part(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        doc.Type = IG_PART_DOCUMENT

        result = feature_mgr.create_normal_cutout(0.005)

        assert "error" in result
        assert "sheet metal feature" in result["error"]
        assert result["unsupported"] is True
        model.NormalCutouts.AddFiniteMulti.assert_not_called()

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, doc, _, model, profile = managers
        doc.Type = IG_SHEET_METAL_DOCUMENT
        result = feature_mgr.create_normal_cutout(0.01, direction="Reverse")
        assert result["direction"] == "Reverse"
        model.NormalCutouts.AddFiniteMulti.assert_called_once_with(
            1,
            (profile,),
            1,  # igLeft
            0.01,
            205,  # igSMFaceCutout
        )


# ============================================================================
# LOFTED CUTOUT
# ============================================================================


class TestLoftedCutout:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        p1, p2 = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2]

        result = feature_mgr.create_lofted_cutout()
        assert result["status"] == "created"
        assert result["type"] == "lofted_cutout"
        assert result["num_profiles"] == 2
        model.LoftedCutouts.AddSimple.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_insufficient_profiles(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_accumulated_profiles.return_value = [MagicMock()]
        result = feature_mgr.create_lofted_cutout()
        assert "error" in result
        assert "at least 2 profiles" in result["error"]

    def test_no_profiles(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_accumulated_profiles.return_value = []
        result = feature_mgr.create_lofted_cutout()
        assert "error" in result
        assert "at least 2 profiles" in result["error"]

    def test_no_base_feature(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        models.Count = 0
        sketch_mgr.get_accumulated_profiles.return_value = [MagicMock(), MagicMock()]
        result = feature_mgr.create_lofted_cutout()
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_with_profile_indices(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        p1, p2, p3 = MagicMock(), MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2, p3]

        result = feature_mgr.create_lofted_cutout(profile_indices=[0, 2])
        assert result["status"] == "created"
        assert result["num_profiles"] == 2


# ============================================================================
# SWEPT CUTOUT
# ============================================================================


class TestSweptCutout:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        profile1 = MagicMock()
        profile2 = MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [profile1, profile2]

        result = feature_mgr.create_swept_cutout()
        assert result["status"] == "created"
        assert result["type"] == "swept_cutout"
        assert result["num_cross_sections"] == 1
        model.SweptCutouts.Add.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_swept_cutout()
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_insufficient_profiles(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_accumulated_profiles.return_value = [MagicMock()]
        result = feature_mgr.create_swept_cutout()
        assert "error" in result
        assert "at least 2 profiles" in result["error"]

    def test_custom_path_index(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        p0, p1, p2 = MagicMock(), MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p0, p1, p2]

        result = feature_mgr.create_swept_cutout(path_profile_index=1)
        assert result["status"] == "created"
        assert result["num_cross_sections"] == 2


# ============================================================================
# HELIX CUTOUT
# ============================================================================


class TestHelixCutout:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis

        result = feature_mgr.create_helix_cutout(pitch=0.005, height=0.03, revolutions=6)
        assert result["status"] == "created"
        assert result["type"] == "helix_cutout"
        assert result["pitch"] == 0.005
        assert result["height"] == 0.03
        assert result["revolutions"] == 6
        model.HelixCutouts.AddFinite.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_helix_cutout(pitch=0.005, height=0.03)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_helix_cutout(pitch=0.005, height=0.03)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_helix_cutout(pitch=0.005, height=0.03)
        assert "error" in result
        assert "No axis" in result["error"]

    def test_auto_revolutions(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        sketch_mgr.get_active_refaxis.return_value = MagicMock()

        result = feature_mgr.create_helix_cutout(pitch=0.01, height=0.05)
        assert result["status"] == "created"
        assert result["revolutions"] == 5.0

    def test_left_hand(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        sketch_mgr.get_active_refaxis.return_value = MagicMock()

        result = feature_mgr.create_helix_cutout(pitch=0.005, height=0.03, direction="Left")
        assert result["status"] == "created"
        assert result["direction"] == "Left"


# ============================================================================
# EXTRUDED CUTOUT THROUGH NEXT
# ============================================================================


class TestExtrudedCutoutThroughNext:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        result = feature_mgr.create_extruded_cutout_through_next("Normal")
        assert result["status"] == "created"
        assert result["type"] == "extruded_cutout_through_next"
        model.ExtrudedCutouts.AddThroughNextMulti.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extruded_cutout_through_next()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_extruded_cutout_through_next()
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_extruded_cutout_through_next("Reverse")
        assert result["status"] == "created"
        assert result["direction"] == "Reverse"


# ============================================================================
# NORMAL CUTOUT THROUGH ALL
# ============================================================================


class TestNormalCutoutThroughAll:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_normal_cutout_through_all("Normal")
        assert result["status"] == "created"
        assert result["type"] == "normal_cutout_through_all"
        model.NormalCutouts.AddThroughAllMulti.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_normal_cutout_through_all()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_normal_cutout_through_all()
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# EXTRUDED CUTOUT FROM TO
# ============================================================================


class TestCreateExtrudedCutoutFromTo:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, profile = managers
        cutouts = MagicMock()
        model.ExtrudedCutouts = cutouts
        ref_planes = MagicMock()
        ref_planes.Count = 4
        from_plane = MagicMock()
        to_plane = MagicMock()
        ref_planes.Item.side_effect = lambda i: {2: from_plane, 3: to_plane}.get(i)
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_extruded_cutout_from_to(2, 3)
        assert result["status"] == "created"
        assert result["type"] == "extruded_cutout_from_to"
        cutouts.AddFromToMulti.assert_called_once_with(1, (profile,), from_plane, to_plane)

    def test_invalid_from_plane(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_extruded_cutout_from_to(0, 2)
        assert "error" in result

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extruded_cutout_from_to(1, 2)
        assert "error" in result

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_extruded_cutout_from_to(1, 2)
        assert "error" in result


# ============================================================================
# EXTRUDED CUTOUT FROM TO V2
# ============================================================================


class TestCreateExtrudedCutoutFromToV2:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        result = feature_mgr.create_extruded_cutout_from_to_v2(1, 2)
        assert result["status"] == "created"
        assert result["type"] == "extruded_cutout_from_to_v2"
        model.ExtrudedCutouts.AddFromToMulti.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extruded_cutout_from_to_v2(1, 2)
        assert "error" in result

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_extruded_cutout_from_to_v2(1, 2)
        assert "error" in result


# ============================================================================
# EXTRUDED CUTOUT BY KEYPOINT
# ============================================================================


class TestCreateExtrudedCutoutByKeypoint:
    """AddFiniteByKeyPointMulti needs a KeyPoint object the server cannot select."""

    def test_reports_unsupported_without_calling_com(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_extruded_cutout_by_keypoint("Normal")
        assert result["unsupported"] is True
        assert "KeyPoint" in result["error"]
        assert result["direction"] == "Normal"
        model.ExtrudedCutouts.AddFiniteByKeyPointMulti.assert_not_called()

    def test_no_profile_still_unsupported(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extruded_cutout_by_keypoint()
        assert result["unsupported"] is True
        model.ExtrudedCutouts.AddFiniteByKeyPointMulti.assert_not_called()

    def test_no_model_still_unsupported(self, feature_mgr, managers):
        _, _, _, models, model, _ = managers
        models.Count = 0
        result = feature_mgr.create_extruded_cutout_by_keypoint()
        assert result["unsupported"] is True
        model.ExtrudedCutouts.AddFiniteByKeyPointMulti.assert_not_called()


# ============================================================================
# REVOLVED CUTOUT SYNC
# ============================================================================


class TestCreateRevolvedCutoutSync:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        result = feature_mgr.create_revolved_cutout_sync(360.0)
        assert result["status"] == "created"
        assert result["type"] == "revolved_cutout_sync"
        model.RevolvedCutouts.AddFiniteMultiSync.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_revolved_cutout_sync()
        assert "error" in result

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_revolved_cutout_sync()
        assert "error" in result


# ============================================================================
# REVOLVED CUTOUT BY KEYPOINT
# ============================================================================


class TestCreateRevolvedCutoutByKeypoint:
    """AddFiniteByKeyPointMulti needs a KeyPoint object the server cannot select."""

    def test_reports_unsupported_without_calling_com(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        sketch_mgr.get_active_refaxis.return_value = MagicMock()
        result = feature_mgr.create_revolved_cutout_by_keypoint()
        assert result["unsupported"] is True
        assert "KeyPoint" in result["error"]
        model.RevolvedCutouts.AddFiniteByKeyPointMulti.assert_not_called()

    def test_no_profile_still_unsupported(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_revolved_cutout_by_keypoint()
        assert result["unsupported"] is True
        model.RevolvedCutouts.AddFiniteByKeyPointMulti.assert_not_called()

    def test_no_refaxis_still_unsupported(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_revolved_cutout_by_keypoint()
        assert result["unsupported"] is True
        model.RevolvedCutouts.AddFiniteByKeyPointMulti.assert_not_called()


# ============================================================================
# NORMAL CUTOUT FROM TO
# ============================================================================


class TestCreateNormalCutoutFromTo:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        result = feature_mgr.create_normal_cutout_from_to(1, 2)
        assert result["status"] == "created"
        assert result["type"] == "normal_cutout_from_to"
        model.NormalCutouts.AddFromToMulti.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_normal_cutout_from_to(1, 2)
        assert "error" in result

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_normal_cutout_from_to(1, 2)
        assert "error" in result


# ============================================================================
# NORMAL CUTOUT THROUGH NEXT
# ============================================================================


class TestCreateNormalCutoutThroughNext:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_normal_cutout_through_next("Normal")
        assert result["status"] == "created"
        assert result["type"] == "normal_cutout_through_next"
        model.NormalCutouts.AddThroughNextMulti.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_normal_cutout_through_next()
        assert "error" in result

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_normal_cutout_through_next()
        assert "error" in result


# ============================================================================
# NORMAL CUTOUT BY KEYPOINT
# ============================================================================


class TestCreateNormalCutoutByKeypoint:
    def test_reports_unsupported(self, feature_mgr, managers):
        """AddFiniteByKeyPointMulti needs a KeyPoint object we cannot select."""
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_normal_cutout_by_keypoint("Normal")
        assert result["unsupported"] is True
        assert "KeyPoint" in result["error"]
        model.NormalCutouts.AddFiniteByKeyPointMulti.assert_not_called()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_normal_cutout_by_keypoint()
        assert "error" in result

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_normal_cutout_by_keypoint()
        assert "error" in result


# ============================================================================
# LOFTED CUTOUT FULL
# ============================================================================


class TestCreateLoftedCutoutFull:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        p1, p2 = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2]
        result = feature_mgr.create_lofted_cutout_full()
        assert result["status"] == "created"
        assert result["type"] == "lofted_cutout_full"
        assert result["num_profiles"] == 2
        # LoftedCutouts.Add(NumSections, CrossSections, CrossSectionTypes,
        #   Origins, SegmentMaps, MaterialSide, StartExtentType,
        #   StartExtentDistance, StartSurfaceOrRefPlane, EndExtentType,
        #   EndExtentDistance, EndSurfaceOrRefPlane, StartTangentType,
        #   StartTangentMagnitude, EndTangentType, EndTangentMagnitude)
        assert model.LoftedCutouts.Add.call_count == 1
        args = model.LoftedCutouts.Add.call_args[0]
        assert len(args) == 16
        assert args[0] == 2
        assert list(args[1]) == [p1, p2]
        assert list(args[2]) == [48, 48]  # igProfileBasedCrossSection
        assert len(args[3]) == 2  # per-profile origins
        assert args[5:] == (2, 44, 0.0, None, 44, 0.0, None, 44, 0.0, 44, 0.0)

    def test_too_few_profiles(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_accumulated_profiles.return_value = [MagicMock()]
        result = feature_mgr.create_lofted_cutout_full()
        assert "error" in result
        assert "at least 2 profiles" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_lofted_cutout_full()
        assert "error" in result


# ============================================================================
# SWEPT CUTOUT MULTI BODY
# ============================================================================


class TestCreateSweptCutoutMultiBody:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        path, cs = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [path, cs]
        result = feature_mgr.create_swept_cutout_multi_body()
        assert result["status"] == "created"
        assert result["type"] == "swept_cutout_multi_body"
        # SweptCutouts.AddMultiBody(NumCurves, TraceCurves, TraceCurveTypes,
        #   NumSections, CrossSections, CrossSectionTypes, Origins, SegmentMaps,
        #   MaterialSide, StartExtentType, StartExtentDistance,
        #   StartSurfaceOrRefPlane, EndExtentType, EndExtentDistance,
        #   EndSurfaceOrRefPlane, NumberOfBodies, BodyArray)
        assert model.SweptCutouts.AddMultiBody.call_count == 1
        args = model.SweptCutouts.AddMultiBody.call_args[0]
        assert len(args) == 17
        assert args[0] == 1
        assert list(args[1]) == [path]
        assert list(args[2]) == [48]
        assert args[3] == 1
        assert list(args[4]) == [cs]
        assert list(args[5]) == [48]
        assert args[8:16] == (2, 44, 0.0, None, 44, 0.0, None, 1)
        assert list(args[16]) == [model.Body]

    def test_too_few_profiles(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_accumulated_profiles.return_value = [MagicMock()]
        result = feature_mgr.create_swept_cutout_multi_body()
        assert "error" in result
        assert "at least 2 profiles" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_swept_cutout_multi_body()
        assert "error" in result


# ============================================================================
# HELIX FROM TO
# ============================================================================


class TestCreateHelixFromTo:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        ref_planes = MagicMock()
        ref_planes.Count = 3
        from_plane, to_plane = MagicMock(), MagicMock()
        ref_planes.Item.side_effect = lambda i: {1: from_plane, 2: to_plane}[i]
        doc.RefPlanes = ref_planes
        result = feature_mgr.create_helix_from_to(1, 2, 0.005)
        assert result["status"] == "created"
        assert result["type"] == "helix_from_to"
        # HelixProtrusions.AddFromTo(HelixAxis, AxisStart, NumCrossSections,
        #   CrossSectionArray, ProfileSide, Height, Pitch, NumberOfTurns,
        #   HelixDir, FromPlane, ToPlane)
        assert model.HelixProtrusions.AddFromTo.call_count == 1
        args = model.HelixProtrusions.AddFromTo.call_args[0]
        assert len(args) == 11
        assert args[0] is refaxis
        assert args[1] == 29  # igStart
        assert args[2] == 1
        # Verified on SE 2026: the helix APIs take a plain list here; a
        # VARIANT(VT_ARRAY | VT_DISPATCH, ...) wrapper is rejected.
        assert args[3] == [profile]
        assert args[4:] == (2, 0.0, 0.005, 0.0, 2, from_plane, to_plane)

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_helix_from_to(1, 2, 0.005)
        assert "error" in result

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_helix_from_to(1, 2, 0.005)
        assert "error" in result


# ============================================================================
# HELIX FROM TO THIN WALL
# ============================================================================


class TestCreateHelixFromToThinWall:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        ref_planes = MagicMock()
        ref_planes.Count = 3
        from_plane, to_plane = MagicMock(), MagicMock()
        ref_planes.Item.side_effect = lambda i: {1: from_plane, 2: to_plane}[i]
        doc.RefPlanes = ref_planes
        result = feature_mgr.create_helix_from_to_thin_wall(1, 2, 0.005, 0.001)
        assert result["status"] == "created"
        assert result["type"] == "helix_from_to_thin_wall"
        # AddFromToWithThinWall(HelixAxis, AxisStart, NumCrossSections,
        #   CrossSectionArray, ProfileSide, Height, Pitch, NumberOfTurns,
        #   HelixDir, FromPlane, ToPlane, ThinWall, AddEndCaps,
        #   RemoveInsideMaterial, Thickness, ThicknessSide)
        assert model.HelixProtrusions.AddFromToWithThinWall.call_count == 1
        args = model.HelixProtrusions.AddFromToWithThinWall.call_args[0]
        assert len(args) == 16
        assert args[0] is refaxis
        assert args[1] == 29  # igStart
        assert args[2] == 1
        # Verified on SE 2026: the helix APIs take a plain list here; a
        # VARIANT(VT_ARRAY | VT_DISPATCH, ...) wrapper is rejected.
        assert args[3] == [profile]
        assert args[4:] == (
            2,
            0.0,
            0.005,
            0.0,
            2,
            from_plane,
            to_plane,
            True,
            False,
            True,
            0.001,
            4,  # igInside
        )

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_helix_from_to_thin_wall(1, 2, 0.005, 0.001)
        assert "error" in result

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_helix_from_to_thin_wall(1, 2, 0.005, 0.001)
        assert "error" in result


# ============================================================================
# HELIX CUTOUT SYNC
# ============================================================================


class TestCreateHelixCutoutSync:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        result = feature_mgr.create_helix_cutout_sync(0.005, 0.05)
        assert result["status"] == "created"
        assert result["type"] == "helix_cutout_sync"
        model.HelixCutouts.AddFiniteSync.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_helix_cutout_sync(0.005, 0.05)
        assert "error" in result

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_helix_cutout_sync(0.005, 0.05)
        assert "error" in result


# ============================================================================
# HELIX CUTOUT FROM TO
# ============================================================================


class TestCreateHelixCutoutFromTo:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        ref_planes = MagicMock()
        ref_planes.Count = 3
        from_plane, to_plane = MagicMock(), MagicMock()
        ref_planes.Item.side_effect = lambda i: {1: from_plane, 2: to_plane}[i]
        doc.RefPlanes = ref_planes
        result = feature_mgr.create_helix_cutout_from_to(1, 2, 0.005)
        assert result["status"] == "created"
        assert result["type"] == "helix_cutout_from_to"
        # HelixCutouts.AddFromTo(HelixAxis, AxisStart, NumCrossSections,
        #   CrossSectionArray, ProfileSide, Height, Pitch, NumberOfTurns,
        #   HelixDir, FromPlane, ToPlane)
        assert model.HelixCutouts.AddFromTo.call_count == 1
        args = model.HelixCutouts.AddFromTo.call_args[0]
        assert len(args) == 11
        assert args[0] is refaxis
        assert args[1] == 29  # igStart
        assert args[2] == 1
        assert list(args[3]) == [profile]
        assert args[4:] == (2, 0.0, 0.005, 0.0, 2, from_plane, to_plane)

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_helix_cutout_from_to(1, 2, 0.005)
        assert "error" in result

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_helix_cutout_from_to(1, 2, 0.005)
        assert "error" in result


# ============================================================================
# EXTRUDED CUTOUT THROUGH NEXT SINGLE
# ============================================================================


class TestCreateExtrudedCutoutThroughNextSingle:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        cutout = MagicMock()
        cutout.Name = "Cutout1"
        model.ExtrudedCutouts.AddThroughNext.return_value = cutout

        result = feature_mgr.create_extruded_cutout_through_next_single()
        assert result["status"] == "created"
        assert result["type"] == "extruded_cutout_through_next_single"
        model.ExtrudedCutouts.AddThroughNext.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extruded_cutout_through_next_single()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_extruded_cutout_through_next_single()
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# EXTRUDED CUTOUT MULTI BODY
# ============================================================================


class TestCreateExtrudedCutoutMultiBody:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        cutout = MagicMock()
        cutout.Name = "CutoutMB1"
        model.ExtrudedCutouts.AddFiniteMultiBody.return_value = cutout

        result = feature_mgr.create_extruded_cutout_multi_body(0.05)
        assert result["status"] == "created"
        assert result["type"] == "extruded_cutout_multi_body"
        assert result["distance"] == 0.05
        model.ExtrudedCutouts.AddFiniteMultiBody.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extruded_cutout_multi_body(0.05)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_extruded_cutout_multi_body(0.05)
        assert "error" in result
        assert "No base feature" in result["error"]


class TestCreateExtrudedCutoutFromToMultiBody:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, profile = managers
        ref_planes = MagicMock()
        ref_planes.Count = 10
        ref_planes.Item.return_value = MagicMock()
        doc.RefPlanes = ref_planes
        cutout = MagicMock()
        cutout.Name = "CutoutFTMB1"
        model.ExtrudedCutouts.AddFromToMultiBody.return_value = cutout

        result = feature_mgr.create_extruded_cutout_from_to_multi_body(4, 5)
        assert result["status"] == "created"
        assert result["type"] == "extruded_cutout_from_to_multi_body"
        model.ExtrudedCutouts.AddFromToMultiBody.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extruded_cutout_from_to_multi_body(4, 5)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_extruded_cutout_from_to_multi_body(4, 5)
        assert "error" in result
        assert "No base feature" in result["error"]


class TestCreateExtrudedCutoutThroughAllMultiBody:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        cutout = MagicMock()
        cutout.Name = "CutoutTAMB1"
        model.ExtrudedCutouts.AddThroughAllMultiBody.return_value = cutout

        result = feature_mgr.create_extruded_cutout_through_all_multi_body()
        assert result["status"] == "created"
        assert result["type"] == "extruded_cutout_through_all_multi_body"
        model.ExtrudedCutouts.AddThroughAllMultiBody.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_extruded_cutout_through_all_multi_body()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        cutout = MagicMock()
        cutout.Name = "CutoutTAMB1"
        model.ExtrudedCutouts.AddThroughAllMultiBody.return_value = cutout

        result = feature_mgr.create_extruded_cutout_through_all_multi_body("Reverse")
        assert result["status"] == "created"
        assert result["direction"] == "Reverse"


class TestCreateRevolvedCutoutMultiBody:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        cutout = MagicMock()
        cutout.Name = "RevCutMB1"
        model.RevolvedCutouts.AddFiniteMultiBody.return_value = cutout

        result = feature_mgr.create_revolved_cutout_multi_body(180.0)
        assert result["status"] == "created"
        assert result["type"] == "revolved_cutout_multi_body"
        model.RevolvedCutouts.AddFiniteMultiBody.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_revolved_cutout_multi_body()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_revolved_cutout_multi_body()
        assert "error" in result
        assert "axis" in result["error"].lower()


class TestCreateRevolvedCutoutFull:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        cutout = MagicMock()
        cutout.Name = "RevCutFull1"
        model.RevolvedCutouts.Add.return_value = cutout

        result = feature_mgr.create_revolved_cutout_full(180.0)
        assert result["status"] == "created"
        assert result["type"] == "revolved_cutout_full"
        model.RevolvedCutouts.Add.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_revolved_cutout_full()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_revolved_cutout_full()
        assert "error" in result
        assert "axis" in result["error"].lower()


class TestCreateRevolvedCutoutFullSync:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        cutout = MagicMock()
        cutout.Name = "RevCutFullSync1"
        model.RevolvedCutouts.AddSync.return_value = cutout

        result = feature_mgr.create_revolved_cutout_full_sync(180.0)
        assert result["status"] == "created"
        assert result["type"] == "revolved_cutout_full_sync"
        model.RevolvedCutouts.AddSync.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_revolved_cutout_full_sync()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_revolved_cutout_full_sync()
        assert "error" in result
        assert "axis" in result["error"].lower()


class TestCreateHelixFromToSync:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        ref_planes = MagicMock()
        ref_planes.Count = 10
        ref_planes.Item.return_value = MagicMock()
        doc.RefPlanes = ref_planes
        helix = MagicMock()
        helix.Name = "HelixSync1"
        model.HelixProtrusions.AddFromToSync.return_value = helix

        result = feature_mgr.create_helix_from_to_sync(4, 5, 0.01)
        assert result["status"] == "created"
        assert result["type"] == "helix_from_to_sync"
        assert result["pitch"] == 0.01
        model.HelixProtrusions.AddFromToSync.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_helix_from_to_sync(4, 5, 0.01)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_helix_from_to_sync(4, 5, 0.01)
        assert "error" in result
        assert "axis" in result["error"].lower()


class TestCreateHelixFromToSyncThinWall:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        ref_planes = MagicMock()
        ref_planes.Count = 10
        ref_planes.Item.return_value = MagicMock()
        doc.RefPlanes = ref_planes
        helix = MagicMock()
        helix.Name = "HelixSyncTW1"
        model.HelixProtrusions.AddFromToSyncWithThinWall.return_value = helix

        result = feature_mgr.create_helix_from_to_sync_thin_wall(4, 5, 0.01, 0.001)
        assert result["status"] == "created"
        assert result["type"] == "helix_from_to_sync_thin_wall"
        assert result["wall_thickness"] == 0.001
        model.HelixProtrusions.AddFromToSyncWithThinWall.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_helix_from_to_sync_thin_wall(4, 5, 0.01, 0.001)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_helix_from_to_sync_thin_wall(4, 5, 0.01, 0.001)
        assert "error" in result
        assert "axis" in result["error"].lower()


class TestCreateHelixCutoutFromToSync:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis
        ref_planes = MagicMock()
        ref_planes.Count = 10
        ref_planes.Item.return_value = MagicMock()
        doc.RefPlanes = ref_planes
        helix = MagicMock()
        helix.Name = "HelixCutSync1"
        model.HelixCutouts.AddFromToSync.return_value = helix

        result = feature_mgr.create_helix_cutout_from_to_sync(4, 5, 0.01)
        assert result["status"] == "created"
        assert result["type"] == "helix_cutout_from_to_sync"
        model.HelixCutouts.AddFromToSync.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_helix_cutout_from_to_sync(4, 5, 0.01)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_helix_cutout_from_to_sync(4, 5, 0.01)
        assert "error" in result
        assert "axis" in result["error"].lower()
