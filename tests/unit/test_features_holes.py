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
# HOLES
# ============================================================================


class TestCreateHole:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        # Mock ProfileSets for hole creation
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        hole_profile = MagicMock()
        ps.Profiles.Add.return_value = hole_profile
        plane = MagicMock()
        doc.RefPlanes.Item.return_value = plane

        result = feature_mgr.create_hole(0.01, 0.01, 0.005, 0.01)
        assert result["status"] == "created"
        assert result["type"] == "hole"
        assert result["diameter"] == 0.005
        model.ExtrudedCutouts.AddFiniteMulti.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_hole(0.01, 0.01, 0.005, 0.01)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_hole_position(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        hole_profile = MagicMock()
        ps.Profiles.Add.return_value = hole_profile
        plane = MagicMock()
        doc.RefPlanes.Item.return_value = plane

        result = feature_mgr.create_hole(0.05, 0.03, 0.01, 0.02)
        assert result["position"] == [0.05, 0.03]
        assert result["depth"] == 0.02

    def test_hole_reverse_direction(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        hole_profile = MagicMock()
        ps.Profiles.Add.return_value = hole_profile
        plane = MagicMock()
        doc.RefPlanes.Item.return_value = plane

        result = feature_mgr.create_hole(0, 0, 0.005, 0.01, direction="Reverse")
        assert result["status"] == "created"
        # Verify igLeft (1) was passed for Reverse direction
        call_args = model.ExtrudedCutouts.AddFiniteMulti.call_args
        assert call_args[0][2] == 1  # igLeft


# ============================================================================
# HOLE THROUGH ALL
# ============================================================================


class TestHoleThroughAll:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        profile_sets = MagicMock()
        ps = MagicMock()
        profile = MagicMock()
        ps.Profiles.Add.return_value = profile
        profile_sets.Add.return_value = ps
        doc.ProfileSets = profile_sets
        ref_planes = MagicMock()
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_hole_through_all(0.01, 0.02, 0.005)
        assert result["status"] == "created"
        assert result["type"] == "hole_through_all"
        assert result["position"] == [0.01, 0.02]
        assert result["diameter"] == 0.005
        model.ExtrudedCutouts.AddThroughAllMulti.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_hole_through_all(0, 0, 0.005)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        profile_sets = MagicMock()
        ps = MagicMock()
        profile = MagicMock()
        ps.Profiles.Add.return_value = profile
        profile_sets.Add.return_value = ps
        doc.ProfileSets = profile_sets
        ref_planes = MagicMock()
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_hole_through_all(0, 0, 0.01, direction="Reverse")
        assert result["status"] == "created"
        assert result["direction"] == "Reverse"


# ============================================================================
# DELETE HOLE
# ============================================================================


class TestDeleteHole:
    def test_success_all(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_delete_hole(0.01, "All")
        assert result["status"] == "created"
        assert result["type"] == "delete_hole"
        assert result["max_diameter"] == 0.01
        assert result["hole_type"] == "All"
        model.DeleteHoles.Add.assert_called_once_with(0, 0.01)

    def test_success_round_only(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_delete_hole(0.005, "Round")
        assert result["status"] == "created"
        model.DeleteHoles.Add.assert_called_once_with(1, 0.005)

    def test_success_nonround(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_delete_hole(0.02, "NonRound")
        assert result["status"] == "created"
        model.DeleteHoles.Add.assert_called_once_with(2, 0.02)

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_delete_hole()
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# DELETE HOLE BY FACE
# ============================================================================


class TestDeleteHoleByFace:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.delete_hole_by_face(0)
        assert result["status"] == "created"
        assert result["type"] == "delete_hole_by_face"
        assert result["face_index"] == 0
        model.DeleteHoles.AddByFace.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.delete_hole_by_face(0)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_face_index(self, feature_mgr, managers):
        result = feature_mgr.delete_hole_by_face(99)
        assert "error" in result
        assert "Invalid face index" in result["error"]


# ============================================================================
# DELETE BLEND
# ============================================================================


class TestDeleteBlend:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_delete_blend(0)
        assert result["status"] == "created"
        assert result["type"] == "delete_blend"
        assert result["face_index"] == 0
        model.DeleteBlends.Add.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_delete_blend(0)
        assert "error" in result

    def test_invalid_face_index(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        body = model.Body
        faces = body.Faces.return_value
        faces.Count = 3
        result = feature_mgr.create_delete_blend(5)
        assert "error" in result
        assert "Invalid face index" in result["error"]


# ============================================================================
# HOLE FROM TO
# ============================================================================


class TestCreateHoleFromTo:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        profile_mock = MagicMock()
        ps.Profiles.Add.return_value = profile_mock
        result = feature_mgr.create_hole_from_to(0.0, 0.0, 0.01, 1, 2)
        assert result["status"] == "created"
        assert result["type"] == "hole_from_to"
        model.ExtrudedCutouts.AddFromToMulti.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_hole_from_to(0.0, 0.0, 0.01, 1, 2)
        assert "error" in result

    def test_invalid_plane(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        result = feature_mgr.create_hole_from_to(0.0, 0.0, 0.01, 0, 2)
        assert "error" in result
        assert "Invalid from_plane_index" in result["error"]


# ============================================================================
# HOLE THROUGH NEXT
# ============================================================================


class TestCreateHoleThroughNext:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        profile_mock = MagicMock()
        ps.Profiles.Add.return_value = profile_mock
        result = feature_mgr.create_hole_through_next(0.0, 0.0, 0.01)
        assert result["status"] == "created"
        assert result["type"] == "hole_through_next"
        model.ExtrudedCutouts.AddThroughNextMulti.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_hole_through_next(0.0, 0.0, 0.01)
        assert "error" in result

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        ps.Profiles.Add.return_value = MagicMock()
        result = feature_mgr.create_hole_through_next(0.0, 0.0, 0.01, "Reverse")
        assert result["status"] == "created"
        assert result["direction"] == "Reverse"


# ============================================================================
# HOLE SYNC
# ============================================================================


class TestCreateHoleSync:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        ps.Profiles.Add.return_value = MagicMock()
        result = feature_mgr.create_hole_sync(0.0, 0.0, 0.01, 0.05)
        assert result["status"] == "created"
        assert result["type"] == "hole_sync"
        model.Holes.AddSync.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_hole_sync(0.0, 0.0, 0.01, 0.05)
        assert "error" in result

    def test_diameter_passed(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        ps.Profiles.Add.return_value = MagicMock()
        result = feature_mgr.create_hole_sync(0.01, 0.02, 0.02, 0.03)
        assert result["status"] == "created"
        assert result["diameter"] == 0.02


# ============================================================================
# HOLE FINITE EX
# ============================================================================


class TestCreateHoleFiniteEx:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        ps.Profiles.Add.return_value = MagicMock()
        result = feature_mgr.create_hole_finite_ex(0.0, 0.0, 0.01, 0.05)
        assert result["status"] == "created"
        assert result["type"] == "hole_finite_ex"
        model.Holes.AddFiniteEx.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_hole_finite_ex(0.0, 0.0, 0.01, 0.05)
        assert "error" in result

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        ps.Profiles.Add.return_value = MagicMock()
        result = feature_mgr.create_hole_finite_ex(0.0, 0.0, 0.01, 0.05, "Reverse")
        assert result["status"] == "created"
        assert result["direction"] == "Reverse"


# ============================================================================
# HOLE FROM TO EX
# ============================================================================


class TestCreateHoleFromToEx:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        ps.Profiles.Add.return_value = MagicMock()
        result = feature_mgr.create_hole_from_to_ex(0.0, 0.0, 0.01, 1, 2)
        assert result["status"] == "created"
        assert result["type"] == "hole_from_to_ex"
        model.Holes.AddFromToEx.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_hole_from_to_ex(0.0, 0.0, 0.01, 1, 2)
        assert "error" in result

    def test_invalid_plane(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        result = feature_mgr.create_hole_from_to_ex(0.0, 0.0, 0.01, 0, 2)
        assert "error" in result
        assert "Invalid from_plane_index" in result["error"]


# ============================================================================
# HOLE THROUGH NEXT EX
# ============================================================================


class TestCreateHoleThroughNextEx:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        ps.Profiles.Add.return_value = MagicMock()
        result = feature_mgr.create_hole_through_next_ex(0.0, 0.0, 0.01)
        assert result["status"] == "created"
        assert result["type"] == "hole_through_next_ex"
        model.Holes.AddThroughNextEx.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_hole_through_next_ex(0.0, 0.0, 0.01)
        assert "error" in result

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        ps.Profiles.Add.return_value = MagicMock()
        result = feature_mgr.create_hole_through_next_ex(0.0, 0.0, 0.01, "Reverse")
        assert result["status"] == "created"
        assert result["direction"] == "Reverse"


# ============================================================================
# HOLE THROUGH ALL EX
# ============================================================================


class TestCreateHoleThroughAllEx:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        ps.Profiles.Add.return_value = MagicMock()
        result = feature_mgr.create_hole_through_all_ex(0.0, 0.0, 0.01)
        assert result["status"] == "created"
        assert result["type"] == "hole_through_all_ex"
        model.Holes.AddThroughAllEx.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_hole_through_all_ex(0.0, 0.0, 0.01)
        assert "error" in result

    def test_plane_index_passed(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        ps.Profiles.Add.return_value = MagicMock()
        result = feature_mgr.create_hole_through_all_ex(0.0, 0.0, 0.01, 2)
        assert result["status"] == "created"
        assert result["plane_index"] == 2


# ============================================================================
# HOLE SYNC EX
# ============================================================================


class TestCreateHoleSyncEx:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        ps.Profiles.Add.return_value = MagicMock()
        result = feature_mgr.create_hole_sync_ex(0.0, 0.0, 0.01, 0.05)
        assert result["status"] == "created"
        assert result["type"] == "hole_sync_ex"
        model.Holes.AddSyncEx.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_hole_sync_ex(0.0, 0.0, 0.01, 0.05)
        assert "error" in result

    def test_depth_returned(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        ps.Profiles.Add.return_value = MagicMock()
        result = feature_mgr.create_hole_sync_ex(0.0, 0.0, 0.01, 0.03)
        assert result["status"] == "created"
        assert result["depth"] == 0.03


# ============================================================================
# HOLE MULTI BODY
# ============================================================================


class TestCreateHoleMultiBody:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        ps.Profiles.Add.return_value = MagicMock()
        result = feature_mgr.create_hole_multi_body(0.0, 0.0, 0.01, 0.05)
        assert result["status"] == "created"
        assert result["type"] == "hole_multi_body"
        model.Holes.AddMultiBody.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_hole_multi_body(0.0, 0.0, 0.01, 0.05)
        assert "error" in result

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        ps.Profiles.Add.return_value = MagicMock()
        result = feature_mgr.create_hole_multi_body(0.0, 0.0, 0.01, 0.05, 1, "Reverse")
        assert result["status"] == "created"
        assert result["direction"] == "Reverse"


# ============================================================================
# HOLE SYNC MULTI BODY
# ============================================================================


class TestCreateHoleSyncMultiBody:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        ps.Profiles.Add.return_value = MagicMock()
        result = feature_mgr.create_hole_sync_multi_body(0.0, 0.0, 0.01, 0.05)
        assert result["status"] == "created"
        assert result["type"] == "hole_sync_multi_body"
        model.Holes.AddSyncMultiBody.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_hole_sync_multi_body(0.0, 0.0, 0.01, 0.05)
        assert "error" in result

    def test_depth_returned(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ps = MagicMock()
        doc.ProfileSets.Add.return_value = ps
        ps.Profiles.Add.return_value = MagicMock()
        result = feature_mgr.create_hole_sync_multi_body(0.0, 0.0, 0.01, 0.04)
        assert result["status"] == "created"
        assert result["depth"] == 0.04
