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
# ROUNDS (FILLETS)
# ============================================================================

class TestCreateRound:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_round(0.002)
        assert result["status"] == "created"
        assert result["type"] == "round"
        assert result["radius"] == 0.002
        assert result["edge_count"] == 2
        model.Rounds.Add.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_round(0.002)
        assert "error" in result
        assert "No features" in result["error"]

    def test_no_faces(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        faces = MagicMock()
        faces.Count = 0
        model.Body.Faces.return_value = faces
        result = feature_mgr.create_round(0.002)
        assert "error" in result
        assert "No faces" in result["error"]


# ============================================================================
# CHAMFERS
# ============================================================================

class TestCreateChamfer:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_chamfer(0.001)
        assert result["status"] == "created"
        assert result["type"] == "chamfer"
        assert result["distance"] == 0.001
        assert result["edge_count"] == 2
        model.Chamfers.AddEqualSetback.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_chamfer(0.001)
        assert "error" in result
        assert "No features" in result["error"]


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
# EXTRUDED CUTOUT
# ============================================================================

class TestExtrudedCutout:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, profile = managers
        result = feature_mgr.create_extruded_cutout(0.01)
        assert result["status"] == "created"
        assert result["type"] == "extruded_cutout"
        model.ExtrudedCutouts.AddFiniteMulti.assert_called_once_with(
            1, (profile,), 2, 0.01  # igRight=2
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
            1, (profile,), 1, 0.01  # igLeft=1
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
            1, (profile,), 2  # igRight=2
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
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, profile = managers
        result = feature_mgr.create_normal_cutout(0.005)
        assert result["status"] == "created"
        assert result["type"] == "normal_cutout"
        assert result["distance"] == 0.005
        model.NormalCutouts.AddFiniteMulti.assert_called_once_with(
            1, (profile,), 2, 0.005  # igRight=2
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

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        result = feature_mgr.create_normal_cutout(0.01, direction="Reverse")
        assert result["direction"] == "Reverse"
        model.NormalCutouts.AddFiniteMulti.assert_called_once_with(
            1, (profile,), 1, 0.01  # igLeft=1
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
# ROUND ON FACE (SELECTIVE)
# ============================================================================

class TestCreateRoundOnFace:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_round_on_face(0.002, 0)
        assert result["status"] == "created"
        assert result["type"] == "round"
        assert result["radius"] == 0.002
        assert result["face_index"] == 0
        assert result["edge_count"] == 2
        model.Rounds.Add.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_round_on_face(0.002, 0)
        assert "error" in result

    def test_invalid_face_index(self, feature_mgr, managers):
        result = feature_mgr.create_round_on_face(0.002, 99)
        assert "error" in result
        assert "Invalid face index" in result["error"]

    def test_face_no_edges(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        face = MagicMock()
        no_edges = MagicMock()
        no_edges.Count = 0
        face.Edges = no_edges
        faces = MagicMock()
        faces.Count = 1
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces
        result = feature_mgr.create_round_on_face(0.002, 0)
        assert "error" in result
        assert "no edges" in result["error"]


# ============================================================================
# CHAMFER ON FACE (SELECTIVE)
# ============================================================================

class TestCreateChamferOnFace:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_chamfer_on_face(0.001, 0)
        assert result["status"] == "created"
        assert result["type"] == "chamfer"
        assert result["distance"] == 0.001
        assert result["face_index"] == 0
        assert result["edge_count"] == 2
        model.Chamfers.AddEqualSetback.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_chamfer_on_face(0.001, 0)
        assert "error" in result

    def test_invalid_face_index(self, feature_mgr, managers):
        result = feature_mgr.create_chamfer_on_face(0.001, 99)
        assert "error" in result
        assert "Invalid face index" in result["error"]


# ============================================================================
# DELETE FACES
# ============================================================================

class TestDeleteFaces:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.delete_faces([0])
        assert result["status"] == "created"
        assert result["type"] == "delete_faces"
        assert result["face_count"] == 1
        model.DeleteFaces.Add.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.delete_faces([0])
        assert "error" in result

    def test_invalid_face_index(self, feature_mgr, managers):
        result = feature_mgr.delete_faces([99])
        assert "error" in result
        assert "Invalid face index" in result["error"]

    def test_multiple_faces(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        # Set up 3 faces
        faces = MagicMock()
        faces.Count = 3
        face1, face2, face3 = MagicMock(), MagicMock(), MagicMock()
        faces.Item.side_effect = lambda i: {1: face1, 2: face2, 3: face3}[i]
        model.Body.Faces.return_value = faces
        result = feature_mgr.delete_faces([0, 2])
        assert result["status"] == "created"
        assert result["face_count"] == 2


# ============================================================================
# GET FEATURE INFO
# ============================================================================

class TestGetFeatureInfo:
    def test_success(self, feature_mgr, managers):
        _, _, _, models, model, _ = managers
        model.Name = "ExtrudedProtrusion_1"
        model.Type = 3
        model.Visible = True
        model.Suppressed = False
        result = feature_mgr.get_feature_info(0)
        assert result["index"] == 0
        assert result["name"] == "ExtrudedProtrusion_1"

    def test_invalid_index(self, feature_mgr, managers):
        result = feature_mgr.get_feature_info(99)
        assert "error" in result

    def test_negative_index(self, feature_mgr, managers):
        result = feature_mgr.get_feature_info(-1)
        assert "error" in result


# ============================================================================
# DIMPLE
# ============================================================================

class TestCreateDimple:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        dimples = MagicMock()
        model.Dimples = dimples
        result = feature_mgr.create_dimple(0.01)
        assert result["status"] == "created"
        assert result["type"] == "dimple"
        assert result["depth"] == 0.01
        dimples.Add.assert_called_once_with(profile, 0.01, 1, 2)

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        dimples = MagicMock()
        model.Dimples = dimples
        result = feature_mgr.create_dimple(0.005, "Reverse")
        assert result["status"] == "created"
        dimples.Add.assert_called_once_with(profile, 0.005, 2, 1)

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_dimple(0.01)
        assert "error" in result

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_dimple(0.01)
        assert "error" in result


# ============================================================================
# ETCH
# ============================================================================

class TestCreateEtch:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        etches = MagicMock()
        model.Etches = etches
        result = feature_mgr.create_etch()
        assert result["status"] == "created"
        assert result["type"] == "etch"
        etches.Add.assert_called_once_with(profile)

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_etch()
        assert "error" in result

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_etch()
        assert "error" in result


# ============================================================================
# RIB
# ============================================================================

class TestCreateRib:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        ribs = MagicMock()
        model.Ribs = ribs
        result = feature_mgr.create_rib(0.005)
        assert result["status"] == "created"
        assert result["type"] == "rib"
        ribs.Add.assert_called_once_with(profile, 1, 0, 2, 0.005)

    def test_symmetric(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        ribs = MagicMock()
        model.Ribs = ribs
        result = feature_mgr.create_rib(0.01, "Symmetric")
        assert result["status"] == "created"
        assert result["direction"] == "Symmetric"

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_rib(0.005)
        assert "error" in result


# ============================================================================
# LIP
# ============================================================================

class TestCreateLip:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        lips = MagicMock()
        model.Lips = lips
        result = feature_mgr.create_lip(0.003)
        assert result["status"] == "created"
        assert result["type"] == "lip"
        lips.Add.assert_called_once_with(profile, 2, 0.003)

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_lip(0.003)
        assert "error" in result


# ============================================================================
# DRAWN CUTOUT
# ============================================================================

class TestCreateDrawnCutout:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        drawn_cutouts = MagicMock()
        model.DrawnCutouts = drawn_cutouts
        result = feature_mgr.create_drawn_cutout(0.002)
        assert result["status"] == "created"
        assert result["type"] == "drawn_cutout"
        drawn_cutouts.Add.assert_called_once_with(profile, 2, 0.002)

    def test_reverse(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        drawn_cutouts = MagicMock()
        model.DrawnCutouts = drawn_cutouts
        result = feature_mgr.create_drawn_cutout(0.002, "Reverse")
        assert result["status"] == "created"
        drawn_cutouts.Add.assert_called_once_with(profile, 1, 0.002)

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_drawn_cutout(0.002)
        assert "error" in result


# ============================================================================
# BEAD
# ============================================================================

class TestCreateBead:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        beads = MagicMock()
        model.Beads = beads
        result = feature_mgr.create_bead(0.003)
        assert result["status"] == "created"
        assert result["type"] == "bead"
        beads.Add.assert_called_once_with(profile, 2, 0.003)

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_bead(0.003)
        assert "error" in result


# ============================================================================
# LOUVER
# ============================================================================

class TestCreateLouver:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        louvers = MagicMock()
        model.Louvers = louvers
        result = feature_mgr.create_louver(0.005)
        assert result["status"] == "created"
        assert result["type"] == "louver"
        louvers.Add.assert_called_once_with(profile, 2, 0.005)

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_louver(0.005)
        assert "error" in result


# ============================================================================
# GUSSET
# ============================================================================

class TestCreateGusset:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        gussets = MagicMock()
        model.Gussets = gussets
        result = feature_mgr.create_gusset(0.002)
        assert result["status"] == "created"
        assert result["type"] == "gusset"
        gussets.Add.assert_called_once_with(profile, 2, 0.002)

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_gusset(0.002)
        assert "error" in result


# ============================================================================
# THREAD
# ============================================================================

class TestCreateThread:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        face = MagicMock()
        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        threads = MagicMock()
        model.Threads = threads

        result = feature_mgr.create_thread(2, 0.001)
        assert result["status"] == "created"
        assert result["type"] == "thread"
        assert result["face_index"] == 2
        threads.Add.assert_called_once_with(face, 0.001)

    def test_invalid_face(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        faces = MagicMock()
        faces.Count = 3
        model.Body.Faces.return_value = faces

        result = feature_mgr.create_thread(5)
        assert "error" in result

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = feature_mgr.create_thread(0)
        assert "error" in result


# ============================================================================
# SLOT
# ============================================================================

class TestCreateSlot:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        slots = MagicMock()
        model.Slots = slots
        result = feature_mgr.create_slot(0.01)
        assert result["status"] == "created"
        assert result["type"] == "slot"
        slots.Add.assert_called_once_with(profile, 2, 0.01)

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_slot(0.01)
        assert "error" in result

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        slots = MagicMock()
        model.Slots = slots
        result = feature_mgr.create_slot(0.01, "Reverse")
        assert result["status"] == "created"
        slots.Add.assert_called_once_with(profile, 1, 0.01)


# ============================================================================
# SPLIT
# ============================================================================

class TestCreateSplit:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        splits = MagicMock()
        model.Splits = splits
        result = feature_mgr.create_split()
        assert result["status"] == "created"
        assert result["type"] == "split"
        splits.Add.assert_called_once_with(profile, 2)

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_split()
        assert "error" in result


# ============================================================================
# SKETCH FILLET
# ============================================================================

class TestSketchFillet:
    def test_success(self):
        from solidedge_mcp.backends.sketching import SketchManager
        dm = MagicMock()
        sm = SketchManager(dm)

        profile = MagicMock()
        line1 = MagicMock()
        line2 = MagicMock()
        lines = MagicMock()
        lines.Count = 2
        lines.Item.side_effect = lambda i: [None, line1, line2][i]
        profile.Lines2d = lines
        sm.active_profile = profile

        result = sm.sketch_fillet(0.005)
        assert result["status"] == "created"
        assert result["type"] == "sketch_fillet"
        profile.Arcs2d.AddByFillet.assert_called_once_with(line1, line2, 0.005)

    def test_no_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager
        dm = MagicMock()
        sm = SketchManager(dm)

        result = sm.sketch_fillet(0.005)
        assert "error" in result


# ============================================================================
# SKETCH CHAMFER
# ============================================================================

class TestSketchChamfer:
    def test_success(self):
        from solidedge_mcp.backends.sketching import SketchManager
        dm = MagicMock()
        sm = SketchManager(dm)

        profile = MagicMock()
        line1 = MagicMock()
        line2 = MagicMock()
        lines = MagicMock()
        lines.Count = 2
        lines.Item.side_effect = lambda i: [None, line1, line2][i]
        profile.Lines2d = lines
        sm.active_profile = profile

        result = sm.sketch_chamfer(0.003)
        assert result["status"] == "created"
        assert result["type"] == "sketch_chamfer"
        profile.Lines2d.AddByChamfer.assert_called_once_with(line1, line2, 0.003, 0.003)

    def test_no_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager
        dm = MagicMock()
        sm = SketchManager(dm)

        result = sm.sketch_chamfer(0.003)
        assert "error" in result


# ============================================================================
# SKETCH MIRROR
# ============================================================================

class TestSketchMirror:
    def test_mirror_x(self):
        from solidedge_mcp.backends.sketching import SketchManager
        dm = MagicMock()
        sm = SketchManager(dm)

        profile = MagicMock()
        line = MagicMock()
        line.StartPoint.X = 0.01
        line.StartPoint.Y = 0.02
        line.EndPoint.X = 0.03
        line.EndPoint.Y = 0.04
        lines = MagicMock()
        lines.Count = 1
        lines.Item.return_value = line
        profile.Lines2d = lines

        circles = MagicMock()
        circles.Count = 0
        profile.Circles2d = circles
        sm.active_profile = profile

        result = sm.sketch_mirror("X")
        assert result["status"] == "created"
        assert result["mirrored_elements"] == 1
        profile.Lines2d.AddBy2Points.assert_called_once_with(0.01, -0.02, 0.03, -0.04)

    def test_no_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager
        dm = MagicMock()
        sm = SketchManager(dm)

        result = sm.sketch_mirror()
        assert "error" in result


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
# UNEQUAL CHAMFER
# ============================================================================

class TestChamferUnequal:
    def _make_fm_with_body(self):
        from solidedge_mcp.backends.features import FeatureManager
        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)

        doc = MagicMock()
        dm.get_active_document.return_value = doc
        model = MagicMock()
        doc.Models.Count = 1
        doc.Models.Item.return_value = model
        body = MagicMock()
        model.Body = body

        edge1, edge2 = MagicMock(), MagicMock()
        face = MagicMock()
        face_edges = MagicMock()
        face_edges.Count = 2
        face_edges.Item.side_effect = lambda i: [None, edge1, edge2][i]
        face.Edges = face_edges

        faces = MagicMock()
        faces.Count = 3
        faces.Item.return_value = face
        body.Faces.return_value = faces

        return fm, model, face

    def test_success(self):
        fm, model, face = self._make_fm_with_body()
        result = fm.create_chamfer_unequal(0.009, 0.001, face_index=0)
        assert result["status"] == "created"
        assert result["type"] == "chamfer_unequal"
        assert result["distance1"] == 0.009
        assert result["distance2"] == 0.001
        model.Chamfers.AddUnequalSetback.assert_called_once()

    def test_no_features(self):
        from solidedge_mcp.backends.features import FeatureManager
        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)
        doc = MagicMock()
        dm.get_active_document.return_value = doc
        doc.Models.Count = 0
        result = fm.create_chamfer_unequal(0.009, 0.001)
        assert "error" in result


# ============================================================================
# ANGLE CHAMFER
# ============================================================================

class TestChamferAngle:
    def _make_fm_with_body(self):
        from solidedge_mcp.backends.features import FeatureManager
        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)

        doc = MagicMock()
        dm.get_active_document.return_value = doc
        model = MagicMock()
        doc.Models.Count = 1
        doc.Models.Item.return_value = model
        body = MagicMock()
        model.Body = body

        edge1 = MagicMock()
        face = MagicMock()
        face_edges = MagicMock()
        face_edges.Count = 1
        face_edges.Item.return_value = edge1
        face.Edges = face_edges

        faces = MagicMock()
        faces.Count = 3
        faces.Item.return_value = face
        body.Faces.return_value = faces

        return fm, model, face

    def test_success(self):
        fm, model, face = self._make_fm_with_body()
        result = fm.create_chamfer_angle(0.003, 45.0, face_index=0)
        assert result["status"] == "created"
        assert result["type"] == "chamfer_angle"
        assert result["angle_degrees"] == 45.0
        model.Chamfers.AddSetbackAngle.assert_called_once()

    def test_no_features(self):
        from solidedge_mcp.backends.features import FeatureManager
        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)
        doc = MagicMock()
        dm.get_active_document.return_value = doc
        doc.Models.Count = 0
        result = fm.create_chamfer_angle(0.003, 45.0)
        assert "error" in result


# ============================================================================
# FACE ROTATE BY EDGE
# ============================================================================

class TestFaceRotateByEdge:
    def _make_fm_with_body(self):
        from solidedge_mcp.backends.features import FeatureManager
        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)

        doc = MagicMock()
        dm.get_active_document.return_value = doc
        model = MagicMock()
        doc.Models.Count = 1
        doc.Models.Item.return_value = model
        body = MagicMock()
        model.Body = body

        edge = MagicMock()
        face = MagicMock()
        face_edges = MagicMock()
        face_edges.Count = 4
        face_edges.Item.return_value = edge
        face.Edges = face_edges

        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        body.Faces.return_value = faces

        return fm, model

    def test_success(self):
        fm, model = self._make_fm_with_body()
        result = fm.create_face_rotate_by_edge(1, 0, 5.0)
        assert result["status"] == "created"
        assert result["type"] == "face_rotate"
        assert result["method"] == "by_edge"
        model.FaceRotates.Add.assert_called_once()

    def test_invalid_face(self):
        fm, model = self._make_fm_with_body()
        result = fm.create_face_rotate_by_edge(10, 0, 5.0)
        assert "error" in result

    def test_no_features(self):
        from solidedge_mcp.backends.features import FeatureManager
        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)
        doc = MagicMock()
        dm.get_active_document.return_value = doc
        doc.Models.Count = 0
        result = fm.create_face_rotate_by_edge(0, 0, 5.0)
        assert "error" in result


# ============================================================================
# FACE ROTATE BY POINTS
# ============================================================================

class TestFaceRotateByPoints:
    def test_success(self):
        from solidedge_mcp.backends.features import FeatureManager
        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)

        doc = MagicMock()
        dm.get_active_document.return_value = doc
        model = MagicMock()
        doc.Models.Count = 1
        doc.Models.Item.return_value = model
        body = MagicMock()
        model.Body = body

        face = MagicMock()
        v1, v2 = MagicMock(), MagicMock()
        vertices = MagicMock()
        vertices.Count = 4
        vertices.Item.side_effect = lambda i: [None, v1, v2, MagicMock(), MagicMock()][i]
        face.Vertices = vertices

        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        body.Faces.return_value = faces

        result = fm.create_face_rotate_by_points(0, 0, 1, 5.0)
        assert result["status"] == "created"
        assert result["method"] == "by_points"
        model.FaceRotates.Add.assert_called_once()

    def test_no_features(self):
        from solidedge_mcp.backends.features import FeatureManager
        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)
        doc = MagicMock()
        dm.get_active_document.return_value = doc
        doc.Models.Count = 0
        result = fm.create_face_rotate_by_points(0, 0, 1, 5.0)
        assert "error" in result


# ============================================================================
# DRAFT ANGLE
# ============================================================================

class TestDraftAngle:
    def test_success(self):
        from solidedge_mcp.backends.features import FeatureManager
        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)

        doc = MagicMock()
        dm.get_active_document.return_value = doc
        model = MagicMock()
        doc.Models.Count = 1
        doc.Models.Item.return_value = model
        body = MagicMock()
        model.Body = body

        face = MagicMock()
        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        body.Faces.return_value = faces

        ref_plane = MagicMock()
        doc.RefPlanes.Item.return_value = ref_plane

        result = fm.create_draft_angle(0, 3.0, plane_index=1)
        assert result["status"] == "created"
        assert result["type"] == "draft_angle"
        assert result["angle_degrees"] == 3.0
        model.Drafts.Add.assert_called_once()

    def test_no_features(self):
        from solidedge_mcp.backends.features import FeatureManager
        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)
        doc = MagicMock()
        dm.get_active_document.return_value = doc
        doc.Models.Count = 0
        result = fm.create_draft_angle(0, 3.0)
        assert "error" in result


# ============================================================================
# REF PLANE NORMAL TO CURVE
# ============================================================================

class TestRefPlaneNormalToCurve:
    def test_success(self):
        from solidedge_mcp.backends.features import FeatureManager
        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)

        doc = MagicMock()
        dm.get_active_document.return_value = doc
        profile = MagicMock()
        sm.get_active_sketch.return_value = profile

        pivot_plane = MagicMock()
        doc.RefPlanes.Item.return_value = pivot_plane
        doc.RefPlanes.Count = 4

        result = fm.create_ref_plane_normal_to_curve("End", 2)
        assert result["status"] == "created"
        assert result["type"] == "ref_plane_normal_to_curve"
        assert result["new_plane_index"] == 4
        doc.RefPlanes.AddNormalToCurve.assert_called_once()

    def test_no_profile(self):
        from solidedge_mcp.backends.features import FeatureManager
        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)

        dm.get_active_document.return_value = MagicMock()
        sm.get_active_sketch.return_value = None

        result = fm.create_ref_plane_normal_to_curve()
        assert "error" in result


# ============================================================================
# DO IDLE
# ============================================================================

class TestDoIdle:
    def test_success(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection
        cm = SolidEdgeConnection()
        cm._is_connected = True
        cm.application = MagicMock()

        result = cm.do_idle()
        assert result["status"] == "success"
        cm.application.DoIdle.assert_called_once()

    def test_not_connected(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection
        cm = SolidEdgeConnection()
        cm._is_connected = False
        cm.application = None

        result = cm.do_idle()
        assert "error" in result


# ============================================================================
# TIER 1: SWEPT CUTOUT
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
# TIER 1: HELIX CUTOUT
# ============================================================================

class TestHelixCutout:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis

        result = feature_mgr.create_helix_cutout(
            pitch=0.005, height=0.03, revolutions=6
        )
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

        result = feature_mgr.create_helix_cutout(
            pitch=0.005, height=0.03, direction="Left"
        )
        assert result["status"] == "created"
        assert result["direction"] == "Left"


# ============================================================================
# TIER 1: VARIABLE ROUND
# ============================================================================

class TestVariableRound:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_variable_round([0.001, 0.002])
        assert result["status"] == "created"
        assert result["type"] == "variable_round"
        assert result["edge_count"] == 2
        model.Rounds.AddVariable.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_variable_round([0.001])
        assert "error" in result
        assert "No features" in result["error"]

    def test_radii_extends_to_edge_count(self, feature_mgr, managers):
        """If fewer radii than edges, last radius should be repeated."""
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_variable_round([0.003])
        assert result["status"] == "created"
        # 2 edges in fixture, 1 radius provided -> extended to [0.003, 0.003]
        assert len(result["radii"]) == 2
        assert result["radii"] == [0.003, 0.003]

    def test_on_specific_face(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_variable_round([0.001, 0.002], face_index=0)
        assert result["status"] == "created"
        assert result["edge_count"] == 2

    def test_invalid_face_index(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_variable_round([0.001], face_index=5)
        assert "error" in result
        assert "Invalid face index" in result["error"]


# ============================================================================
# TIER 1: BLEND
# ============================================================================

class TestBlend:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_blend(0.003)
        assert result["status"] == "created"
        assert result["type"] == "blend"
        assert result["radius"] == 0.003
        assert result["edge_count"] == 2
        model.Blends.Add.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_blend(0.003)
        assert "error" in result
        assert "No features" in result["error"]

    def test_on_specific_face(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_blend(0.002, face_index=0)
        assert result["status"] == "created"
        assert result["edge_count"] == 2

    def test_invalid_face_index(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_blend(0.002, face_index=99)
        assert "error" in result
        assert "Invalid face index" in result["error"]


# ============================================================================
# TIER 1: REFERENCE PLANE BY ANGLE
# ============================================================================

class TestRefPlaneByAngle:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 4  # 3 default + 1 new
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_by_angle(1, 45.0)
        assert result["status"] == "created"
        assert result["type"] == "reference_plane"
        assert result["method"] == "angular_by_angle"
        assert result["angle_degrees"] == 45.0
        ref_planes.AddAngularByAngle.assert_called_once()

    def test_invalid_plane_index(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_by_angle(5, 30.0)
        assert "error" in result
        assert "Invalid plane index" in result["error"]

    def test_reverse_side(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 4
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_by_angle(2, 60.0, normal_side="Reverse")
        assert result["status"] == "created"
        assert result["normal_side"] == "Reverse"


# ============================================================================
# TIER 1: REFERENCE PLANE BY 3 POINTS
# ============================================================================

class TestRefPlaneBy3Points:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 4
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_by_3_points(
            0, 0, 0,
            0.1, 0, 0,
            0, 0.1, 0
        )
        assert result["status"] == "created"
        assert result["type"] == "reference_plane"
        assert result["method"] == "by_3_points"
        assert result["point1"] == [0, 0, 0]
        assert result["point2"] == [0.1, 0, 0]
        assert result["point3"] == [0, 0.1, 0]
        ref_planes.AddBy3Points.assert_called_once_with(
            0, 0, 0, 0.1, 0, 0, 0, 0.1, 0
        )


# ============================================================================
# TIER 1: REFERENCE PLANE MID-PLANE
# ============================================================================

class TestRefPlaneMidPlane:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 4
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_midplane(1, 3)
        assert result["status"] == "created"
        assert result["type"] == "reference_plane"
        assert result["method"] == "mid_plane"
        assert result["plane1_index"] == 1
        assert result["plane2_index"] == 3
        ref_planes.AddMidPlane.assert_called_once()

    def test_invalid_plane1(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_midplane(5, 1)
        assert "error" in result
        assert "Invalid plane1" in result["error"]

    def test_invalid_plane2(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_midplane(1, 5)
        assert "error" in result
        assert "Invalid plane2" in result["error"]


# ============================================================================
# TIER 1: HOLE THROUGH ALL
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
# TIER 1: BOX CUTOUT
# ============================================================================

class TestBoxCutout:
    def test_success(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        ref_planes = MagicMock()
        doc.RefPlanes = ref_planes
        box_features = MagicMock()
        models.BoxFeatures = box_features

        result = feature_mgr.create_box_cutout_by_two_points(
            0, 0, 0, 0.05, 0.05, 0.05
        )
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
# TIER 1: CYLINDER CUTOUT
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
# TIER 1: SPHERE CUTOUT
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


# ============================================================================
# TIER 2: EXTRUDED CUTOUT THROUGH NEXT
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
# TIER 2: NORMAL CUTOUT THROUGH ALL
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
# TIER 2: DELETE HOLE
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
# TIER 2: DELETE BLEND
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
# TIER 3: CONVERT FEATURE TYPE
# ============================================================================

class TestConvertFeatureType:
    def test_convert_to_cutout(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        feat = MagicMock()
        feat.Name = "Protrusion_1"
        cutout = MagicMock()
        cutout.Name = "Cutout_1"
        feat.ConvertToCutout.return_value = cutout

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = feature_mgr.convert_feature_type("Protrusion_1", "cutout")
        assert result["status"] == "converted"
        assert result["target_type"] == "cutout"
        assert result["new_name"] == "Cutout_1"
        feat.ConvertToCutout.assert_called_once()

    def test_convert_to_protrusion(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        feat = MagicMock()
        feat.Name = "Cutout_1"
        protrusion = MagicMock()
        protrusion.Name = "Protrusion_1"
        feat.ConvertToProtrusion.return_value = protrusion

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = feature_mgr.convert_feature_type("Cutout_1", "protrusion")
        assert result["status"] == "converted"
        assert result["target_type"] == "protrusion"
        assert result["new_name"] == "Protrusion_1"

    def test_feature_not_found(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        feat = MagicMock()
        feat.Name = "Other_1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = feature_mgr.convert_feature_type("NonExistent", "cutout")
        assert "error" in result
        assert "not found" in result["error"]

    def test_invalid_target_type(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        feat = MagicMock()
        feat.Name = "Protrusion_1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = feature_mgr.convert_feature_type("Protrusion_1", "invalid")
        assert "error" in result
        assert "Invalid target_type" in result["error"]


# ============================================================================
# EMBOSS
# ============================================================================

class TestCreateEmboss:
    def test_basic_emboss(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers
        emboss_features = MagicMock()
        model.EmbossFeatures = emboss_features

        result = feature_mgr.create_emboss([0])
        assert result["status"] == "created"
        assert result["type"] == "emboss"
        assert result["face_count"] == 1
        emboss_features.Add.assert_called_once()

    def test_multiple_faces(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers
        body = model.Body
        face2 = MagicMock()
        faces = body.Faces.return_value
        faces.Count = 3
        faces.Item.side_effect = lambda i: MagicMock()

        emboss_features = MagicMock()
        model.EmbossFeatures = emboss_features

        result = feature_mgr.create_emboss([0, 1, 2], clearance=0.002, thickness=0.001, thicken=True)
        assert result["status"] == "created"
        assert result["face_count"] == 3
        assert result["clearance"] == 0.002
        assert result["thickness"] == 0.001
        assert result["thicken"] is True

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers
        models.Count = 0

        result = feature_mgr.create_emboss([0])
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_face_index(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers

        result = feature_mgr.create_emboss([99])
        assert "error" in result
        assert "Invalid face index" in result["error"]

    def test_empty_face_indices(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers

        result = feature_mgr.create_emboss([])
        assert "error" in result
        assert "at least one" in result["error"]


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
            face_index=0, edge_index=0, flange_length=0.02,
            side="Left", inside_radius=0.003, bend_angle=90.0
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
