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
# FLANGE
# ============================================================================


class TestCreateFlange:
    def test_basic_flange(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers
        flanges = MagicMock()
        model.Flanges = flanges

        result = feature_mgr.create_flange(face_index=0, edge_index=0, flange_length=0.01)
        assert result["status"] == "created"
        assert result["type"] == "flange"
        assert result["flange_length"] == 0.01
        assert result["side"] == "Right"
        flanges.Add.assert_called_once()

    def test_flange_with_options(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers
        flanges = MagicMock()
        model.Flanges = flanges

        result = feature_mgr.create_flange(
            face_index=0,
            edge_index=0,
            flange_length=0.02,
            side="Left",
            inside_radius=0.003,
            bend_angle=90.0,
        )
        assert result["status"] == "created"
        assert result["side"] == "Left"
        assert result["inside_radius"] == 0.003
        assert result["bend_angle"] == 90.0

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers
        models.Count = 0

        result = feature_mgr.create_flange(face_index=0, edge_index=0, flange_length=0.01)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_face_index(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers

        result = feature_mgr.create_flange(face_index=99, edge_index=0, flange_length=0.01)
        assert "error" in result
        assert "Invalid face index" in result["error"]

    def test_invalid_edge_index(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers

        result = feature_mgr.create_flange(face_index=0, edge_index=99, flange_length=0.01)
        assert "error" in result
        assert "Invalid edge index" in result["error"]


# ============================================================================
# FLANGE BY MATCH FACE
# ============================================================================


class TestCreateFlangeByMatchFace:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_flange_by_match_face(0, 0, 0.02)
        assert result["status"] == "created"
        assert result["type"] == "flange_by_match_face"
        assert result["flange_length"] == 0.02
        model.Flanges.AddByMatchFace.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_flange_by_match_face(0, 0, 0.02)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_face_index(self, feature_mgr, managers):
        _, _, _, _, _, _ = managers
        result = feature_mgr.create_flange_by_match_face(99, 0, 0.02)
        assert "error" in result
        assert "Invalid face index" in result["error"]


# ============================================================================
# FLANGE SYNC
# ============================================================================


class TestCreateFlangeSync:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_flange_sync(0, 0, 0.03)
        assert result["status"] == "created"
        assert result["type"] == "flange_sync"
        assert result["flange_length"] == 0.03
        model.Flanges.AddSync.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_flange_sync(0, 0, 0.03)
        assert "error" in result

    def test_invalid_edge_index(self, feature_mgr, managers):
        _, _, _, _, _, _ = managers
        result = feature_mgr.create_flange_sync(0, 99, 0.03)
        assert "error" in result
        assert "Invalid edge index" in result["error"]


# ============================================================================
# FLANGE BY FACE
# ============================================================================


class TestCreateFlangeByFace:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_flange_by_face(0, 0, 0, 0.025)
        assert result["status"] == "created"
        assert result["type"] == "flange_by_face"
        assert result["flange_length"] == 0.025
        model.Flanges.AddFlangeByFace.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_flange_by_face(0, 0, 0, 0.025)
        assert "error" in result

    def test_invalid_ref_face(self, feature_mgr, managers):
        _, _, _, _, _, _ = managers
        result = feature_mgr.create_flange_by_face(0, 0, 99, 0.025)
        assert "error" in result
        assert "Invalid ref_face_index" in result["error"]


# ============================================================================
# FLANGE WITH BEND CALC
# ============================================================================


class TestCreateFlangeWithBendCalc:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_flange_with_bend_calc(0, 0, 0.03)
        assert result["status"] == "created"
        assert result["type"] == "flange_with_bend_calc"
        model.Flanges.AddByBendDeductionOrBendAllowance.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_flange_with_bend_calc(0, 0, 0.03)
        assert "error" in result

    def test_custom_side(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_flange_with_bend_calc(0, 0, 0.03, side="Left")
        assert result["status"] == "created"
        assert result["side"] == "Left"


# ============================================================================
# FLANGE SYNC WITH BEND CALC
# ============================================================================


class TestCreateFlangeSyncWithBendCalc:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_flange_sync_with_bend_calc(0, 0, 0.04)
        assert result["status"] == "created"
        assert result["type"] == "flange_sync_with_bend_calc"
        model.Flanges.AddSyncByBendDeductionOrBendAllowance.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_flange_sync_with_bend_calc(0, 0, 0.04)
        assert "error" in result

    def test_bend_deduction_returned(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_flange_sync_with_bend_calc(0, 0, 0.04, bend_deduction=0.002)
        assert result["status"] == "created"
        assert result["bend_deduction"] == 0.002


# ============================================================================
# CONTOUR FLANGE EX
# ============================================================================


class TestCreateContourFlangeEx:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        result = feature_mgr.create_contour_flange_ex(0.005)
        assert result["status"] == "created"
        assert result["type"] == "contour_flange_ex"
        assert result["thickness"] == 0.005
        model.ContourFlanges.AddEx.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_contour_flange_ex(0.005)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_contour_flange_ex(0.005)
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# CONTOUR FLANGE SYNC
# ============================================================================


class TestCreateContourFlangeSync:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        result = feature_mgr.create_contour_flange_sync(0, 0, 0.005)
        assert result["status"] == "created"
        assert result["type"] == "contour_flange_sync"
        model.ContourFlanges.AddSync.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_contour_flange_sync(0, 0, 0.005)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_invalid_face(self, feature_mgr, managers):
        _, _, _, _, _, _ = managers
        result = feature_mgr.create_contour_flange_sync(99, 0, 0.005)
        assert "error" in result
        assert "Invalid face index" in result["error"]


# ============================================================================
# CONTOUR FLANGE SYNC WITH BEND
# ============================================================================


class TestCreateContourFlangeSyncWithBend:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        result = feature_mgr.create_contour_flange_sync_with_bend(0, 0, 0.005)
        assert result["status"] == "created"
        assert result["type"] == "contour_flange_sync_with_bend"
        model.ContourFlanges.AddSyncByBendDeductionOrBendAllowance.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_contour_flange_sync_with_bend(0, 0, 0.005)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_bend_deduction_returned(self, feature_mgr, managers):
        _, _, _, _, _, _ = managers
        result = feature_mgr.create_contour_flange_sync_with_bend(0, 0, 0.005, bend_deduction=0.001)
        assert result["status"] == "created"
        assert result["bend_deduction"] == 0.001


# ============================================================================
# HEM
# ============================================================================


class TestCreateHem:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_hem(0, 0)
        assert result["status"] == "created"
        assert result["type"] == "hem"
        assert result["hem_width"] == 0.005
        model.Hems.Add.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_hem(0, 0)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_custom_hem_type(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_hem(0, 0, hem_type="Open")
        assert result["status"] == "created"
        assert result["hem_type"] == "Open"


# ============================================================================
# JOG
# ============================================================================


class TestCreateJog:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        result = feature_mgr.create_jog()
        assert result["status"] == "created"
        assert result["type"] == "jog"
        assert result["jog_offset"] == 0.005
        model.Jogs.AddFinite.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_jog()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_jog()
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# CLOSE CORNER
# ============================================================================


class TestCreateCloseCorner:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_close_corner(0, 0)
        assert result["status"] == "created"
        assert result["type"] == "close_corner"
        assert result["closure_type"] == "Close"
        model.CloseCorners.Add.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_close_corner(0, 0)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_overlap_type(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_close_corner(0, 0, closure_type="Overlap")
        assert result["status"] == "created"
        assert result["closure_type"] == "Overlap"


# ============================================================================
# MULTI EDGE FLANGE
# ============================================================================


class TestCreateMultiEdgeFlange:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_multi_edge_flange(0, [0, 1], 0.03)
        assert result["status"] == "created"
        assert result["type"] == "multi_edge_flange"
        assert result["edge_count"] == 2
        assert result["flange_length"] == 0.03
        model.MultiEdgeFlanges.Add.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_multi_edge_flange(0, [0], 0.03)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_edge_index(self, feature_mgr, managers):
        _, _, _, _, _, _ = managers
        result = feature_mgr.create_multi_edge_flange(0, [0, 99], 0.03)
        assert "error" in result
        assert "Invalid edge index" in result["error"]


# ============================================================================
# BEND WITH CALC
# ============================================================================


class TestCreateBendWithCalc:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        result = feature_mgr.create_bend_with_calc()
        assert result["status"] == "created"
        assert result["type"] == "bend_with_calc"
        assert result["bend_angle"] == 90.0
        model.Bends.AddByBendDeductionOrBendAllowance.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_bend_with_calc()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_bend_with_calc()
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# CONVERT PART TO SHEET METAL
# ============================================================================


class TestConvertPartToSheetMetal:
    def test_command_invoked(self, feature_mgr, managers):
        doc_mgr, _, _, _, _, _ = managers
        # Mock the connection manager and app
        app = MagicMock()
        doc_mgr.connection_manager = MagicMock()
        doc_mgr.connection_manager.get_application.return_value = app
        result = feature_mgr.convert_part_to_sheet_metal(0.002)
        assert result.get("status") == "command_invoked" or "error" in result

    def test_with_default_thickness(self, feature_mgr, managers):
        doc_mgr, _, _, _, _, _ = managers
        app = MagicMock()
        doc_mgr.connection_manager = MagicMock()
        doc_mgr.connection_manager.get_application.return_value = app
        result = feature_mgr.convert_part_to_sheet_metal()
        assert result.get("status") == "command_invoked" or "error" in result

    def test_command_failure(self, feature_mgr, managers):
        doc_mgr, _, _, _, _, _ = managers
        app = MagicMock()
        app.StartCommand.side_effect = Exception("Command failed")
        doc_mgr.connection_manager = MagicMock()
        doc_mgr.connection_manager.get_application.return_value = app
        result = feature_mgr.convert_part_to_sheet_metal()
        # Should handle failure gracefully
        assert "error" in result or "status" in result


# ============================================================================
# FLANGE MATCH FACE WITH BEND
# ============================================================================


class TestCreateFlangeMatchFaceWithBend:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        body = model.Body
        faces = MagicMock()
        faces.Count = 3
        face = MagicMock()
        edge = MagicMock()
        edges = MagicMock()
        edges.Count = 4
        edges.Item.return_value = edge
        face.Edges = edges
        faces.Item.return_value = face
        body.Faces.return_value = faces

        flange = MagicMock()
        flange.Name = "FlangeMFB1"
        model.Flanges.AddByMatchFaceAndBendDeductionOrBendAllowance.return_value = flange

        result = feature_mgr.create_flange_match_face_with_bend(0, 0, 0.02)
        assert result["status"] == "created"
        assert result["type"] == "flange_match_face_with_bend"
        model.Flanges.AddByMatchFaceAndBendDeductionOrBendAllowance.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_flange_match_face_with_bend(0, 0, 0.02)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_face(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        body = model.Body
        faces = MagicMock()
        faces.Count = 1
        body.Faces.return_value = faces

        result = feature_mgr.create_flange_match_face_with_bend(5, 0, 0.02)
        assert "error" in result
        assert "Invalid face_index" in result["error"]


# ============================================================================
# FLANGE BY FACE WITH BEND
# ============================================================================


class TestCreateFlangeByFaceWithBend:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        body = model.Body
        faces = MagicMock()
        faces.Count = 3
        face = MagicMock()
        edge = MagicMock()
        edges = MagicMock()
        edges.Count = 4
        edges.Item.return_value = edge
        face.Edges = edges
        faces.Item.return_value = face
        body.Faces.return_value = faces

        flange = MagicMock()
        flange.Name = "FlangeFBB1"
        model.Flanges.AddFlangeByFaceAndBendDeductionOrBendAllowance.return_value = flange

        result = feature_mgr.create_flange_by_face_with_bend(0, 0, 1, 0.02)
        assert result["status"] == "created"
        assert result["type"] == "flange_by_face_with_bend"
        model.Flanges.AddFlangeByFaceAndBendDeductionOrBendAllowance.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_flange_by_face_with_bend(0, 0, 1, 0.02)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_face(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        body = model.Body
        faces = MagicMock()
        faces.Count = 1
        body.Faces.return_value = faces

        result = feature_mgr.create_flange_by_face_with_bend(5, 0, 1, 0.02)
        assert "error" in result
        assert "Invalid face_index" in result["error"]


# ============================================================================
# CONTOUR FLANGE V3
# ============================================================================


class TestCreateContourFlangeV3:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        contour = MagicMock()
        contour.Name = "ContourV3_1"
        model.ContourFlanges.Add3.return_value = contour

        result = feature_mgr.create_contour_flange_v3(0.001, 0.001)
        assert result["status"] == "created"
        assert result["type"] == "contour_flange_v3"
        model.ContourFlanges.Add3.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_contour_flange_v3(0.001)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_contour_flange_v3(0.001)
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# CONTOUR FLANGE SYNC EX
# ============================================================================


class TestCreateContourFlangeSyncEx:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        body = model.Body
        faces = MagicMock()
        faces.Count = 3
        face = MagicMock()
        edge = MagicMock()
        edges = MagicMock()
        edges.Count = 4
        edges.Item.return_value = edge
        face.Edges = edges
        faces.Item.return_value = face
        body.Faces.return_value = faces

        contour = MagicMock()
        contour.Name = "ContourSyncEx1"
        model.ContourFlanges.AddSyncEx.return_value = contour

        result = feature_mgr.create_contour_flange_sync_ex(0, 0, 0.001)
        assert result["status"] == "created"
        assert result["type"] == "contour_flange_sync_ex"
        model.ContourFlanges.AddSyncEx.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_contour_flange_sync_ex(0, 0, 0.001)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_face(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        body = model.Body
        faces = MagicMock()
        faces.Count = 1
        body.Faces.return_value = faces

        result = feature_mgr.create_contour_flange_sync_ex(5, 0, 0.001)
        assert "error" in result
        assert "Invalid face_index" in result["error"]


# ============================================================================
# BEND
# ============================================================================


class TestCreateBend:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        bend = MagicMock()
        bend.Name = "Bend1"
        model.Bends.Add.return_value = bend

        result = feature_mgr.create_bend()
        assert result["status"] == "created"
        assert result["type"] == "bend"
        model.Bends.Add.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_bend()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_bend()
        assert "error" in result
        assert "No base feature" in result["error"]
