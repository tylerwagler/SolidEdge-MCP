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
# DELETE FACES NO HEAL
# ============================================================================


class TestDeleteFacesNoHeal:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.delete_faces_no_heal([0])
        assert result["status"] == "created"
        assert result["type"] == "delete_faces_no_heal"
        assert result["face_count"] == 1
        model.DeleteFaces.AddNoHeal.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.delete_faces_no_heal([0])
        assert "error" in result

    def test_invalid_face_index(self, feature_mgr, managers):
        result = feature_mgr.delete_faces_no_heal([99])
        assert "error" in result
        assert "Invalid face index" in result["error"]


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
        _, _, doc, _, model, _ = managers
        cyl_face = MagicMock()
        end_face = MagicMock()
        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = cyl_face
        model.Body.Faces.return_value = faces

        # Mock geometry for auto-diameter detection
        cyl_face.Geometry.Radius = 0.005

        # Mock _find_cylinder_end_face to return end_face
        cyl_edges = MagicMock()
        cyl_edges.Count = 1
        cyl_edge = MagicMock()
        cyl_edges.Item.return_value = cyl_edge
        cyl_face.Edges = cyl_edges
        # The end_face shares an edge with cyl_face
        end_edges = MagicMock()
        end_edges.Count = 1
        end_edges.Item.return_value = cyl_edge  # same edge object
        end_face.Edges = end_edges
        # Make faces.Item return end_face for second call (fi=1 candidate)
        # but cyl_face for the thread target (face_index+1=3)
        def faces_item_side_effect(idx):
            if idx == 3:  # face_index=2, 1-based=3
                return cyl_face
            return end_face
        faces.Item.side_effect = faces_item_side_effect

        # Mock HoleDataCollection
        hole_data = MagicMock()
        doc.HoleDataCollection.Add.return_value = hole_data

        threads = MagicMock()
        model.Threads = threads

        result = feature_mgr.create_thread(2)
        assert result["status"] == "created"
        assert result["type"] == "cosmetic_thread"
        assert result["face_index"] == 2
        threads.Add.assert_called_once()

    def test_invalid_face(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        faces = MagicMock()
        faces.Count = 3
        model.Body.Faces.return_value = faces

        result = feature_mgr.create_thread(5)
        assert "error" in result
        assert "Invalid face index" in result["error"]

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
# CONVERT FEATURE TYPE
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
        MagicMock()
        faces = body.Faces.return_value
        faces.Count = 3
        faces.Item.side_effect = lambda i: MagicMock()

        emboss_features = MagicMock()
        model.EmbossFeatures = emboss_features

        result = feature_mgr.create_emboss(
            [0, 1, 2],
            clearance=0.002,
            thickness=0.001,
            thicken=True,
        )
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
# BATCH 9: DIMPLE EX
# ============================================================================


class TestCreateDimpleEx:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        result = feature_mgr.create_dimple_ex(0.003)
        assert result["status"] == "created"
        assert result["type"] == "dimple_ex"
        assert result["depth"] == 0.003
        model.Dimples.AddEx.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_dimple_ex(0.003)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_dimple_ex(0.003)
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# BATCH 9: THREAD EX
# ============================================================================


class TestCreateThreadEx:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        cyl_face = MagicMock()
        end_face = MagicMock()
        faces = MagicMock()
        faces.Count = 2
        # face_index=0 → 1-based=1 → cyl_face; fi=2 → end_face (candidate)
        def faces_item_side_effect(idx):
            if idx == 1:
                return cyl_face
            return end_face
        faces.Item.side_effect = faces_item_side_effect
        model.Body.Faces.return_value = faces

        cyl_face.Geometry.Radius = 0.005
        cyl_edges = MagicMock()
        cyl_edges.Count = 1
        cyl_edge = MagicMock()
        cyl_edges.Item.return_value = cyl_edge
        cyl_face.Edges = cyl_edges
        end_edges = MagicMock()
        end_edges.Count = 1
        end_edges.Item.return_value = cyl_edge
        end_face.Edges = end_edges

        hole_data = MagicMock()
        doc.HoleDataCollection.Add.return_value = hole_data

        threads = MagicMock()
        model.Threads = threads

        result = feature_mgr.create_thread_ex(0, 0.01, 0.001)
        assert result["status"] == "created"
        assert result["type"] == "physical_thread"
        assert result["face_index"] == 0
        threads.AddEx.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_thread_ex(0, 0.01)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_face(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        faces = MagicMock()
        faces.Count = 1
        model.Body.Faces.return_value = faces
        result = feature_mgr.create_thread_ex(5, 0.01)
        assert "error" in result
        assert "Invalid face index" in result["error"]


# ============================================================================
# BATCH 9: SLOT EX
# ============================================================================


class TestCreateSlotEx:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        result = feature_mgr.create_slot_ex(0.005, 0.01)
        assert result["status"] == "created"
        assert result["type"] == "slot_ex"
        assert result["width"] == 0.005
        assert result["depth"] == 0.01
        model.Slots.AddEx.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_slot_ex(0.005, 0.01)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_slot_ex(0.005, 0.01)
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# BATCH 9: SLOT SYNC
# ============================================================================


class TestCreateSlotSync:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        result = feature_mgr.create_slot_sync(0.005, 0.01)
        assert result["status"] == "created"
        assert result["type"] == "slot_sync"
        assert result["width"] == 0.005
        assert result["depth"] == 0.01
        model.Slots.AddSync.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_slot_sync(0.005, 0.01)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_slot_sync(0.005, 0.01)
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# BATCH 9: DRAWN CUTOUT EX
# ============================================================================


class TestCreateDrawnCutoutEx:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        result = feature_mgr.create_drawn_cutout_ex(0.005)
        assert result["status"] == "created"
        assert result["type"] == "drawn_cutout_ex"
        assert result["depth"] == 0.005
        model.DrawnCutouts.AddEx.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_drawn_cutout_ex(0.005)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_drawn_cutout_ex(0.005)
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# BATCH 9: LOUVER SYNC
# ============================================================================


class TestCreateLouverSync:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        result = feature_mgr.create_louver_sync(0.003)
        assert result["status"] == "created"
        assert result["type"] == "louver_sync"
        assert result["depth"] == 0.003
        model.Louvers.AddSync.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_louver_sync(0.003)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_louver_sync(0.003)
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# BATCH 9: THICKEN SYNC
# ============================================================================


class TestCreateThickenSync:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_thicken_sync(0.002)
        assert result["status"] == "created"
        assert result["type"] == "thicken_sync"
        assert result["thickness"] == 0.002
        model.Thickens.AddSync.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_thicken_sync(0.002)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_direction_reverse(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_thicken_sync(0.003, direction="Reverse")
        assert result["status"] == "created"
        assert result["direction"] == "Reverse"
        model.Thickens.AddSync.assert_called_once()


# ============================================================================
# BATCH 9: MIRROR SYNC EX
# ============================================================================


def _setup_feature_lookup(doc, feature_name="Extrude1"):
    """Helper to set up DesignEdgebarFeatures mock for feature lookup."""
    feat = MagicMock()
    feat.Name = feature_name
    features = MagicMock()
    features.Count = 1
    features.Item.return_value = feat
    doc.DesignEdgebarFeatures = features

    ref_planes = MagicMock()
    ref_planes.Count = 3
    ref_planes.Item.return_value = MagicMock()
    doc.RefPlanes = ref_planes

    return feat


class TestCreateMirrorSyncEx:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Extrude1")
        mc = MagicMock()
        mirror = MagicMock()
        mirror.Name = "Mirror1"
        mc.AddSyncEx.return_value = mirror
        model.MirrorCopies = mc

        result = feature_mgr.create_mirror_sync_ex("Extrude1", 1)
        assert result["status"] == "created"
        assert result["type"] == "mirror_sync_ex"
        assert result["feature"] == "Extrude1"
        mc.AddSyncEx.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_mirror_sync_ex("Extrude1", 1)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_feature_not_found(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        feat = MagicMock()
        feat.Name = "OtherFeature"
        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = feature_mgr.create_mirror_sync_ex("NonExistent", 1)
        assert "error" in result
        assert "not found" in result["error"]


# ============================================================================
# BATCH 9: PATTERN RECTANGULAR EX
# ============================================================================


class TestCreatePatternRectangularEx:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")
        pattern = MagicMock()
        pattern.Name = "Pattern1"
        model.Patterns.AddByRectangularEx.return_value = pattern

        result = feature_mgr.create_pattern_rectangular_ex("Hole1", 3, 2, 0.01, 0.02)
        assert result["status"] == "created"
        assert result["type"] == "pattern_rectangular_ex"
        assert result["x_count"] == 3
        assert result["y_count"] == 2
        model.Patterns.AddByRectangularEx.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_pattern_rectangular_ex("Hole1", 3, 2, 0.01, 0.02)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_feature_not_found(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        feat = MagicMock()
        feat.Name = "OtherFeature"
        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = feature_mgr.create_pattern_rectangular_ex("Missing", 3, 2, 0.01, 0.02)
        assert "error" in result
        assert "not found" in result["error"]


# ============================================================================
# BATCH 9: PATTERN CIRCULAR EX
# ============================================================================


class TestCreatePatternCircularEx:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")
        pattern = MagicMock()
        pattern.Name = "CircPattern1"
        model.Patterns.AddByCircularEx.return_value = pattern

        result = feature_mgr.create_pattern_circular_ex("Hole1", 6, 360.0, 0)
        assert result["status"] == "created"
        assert result["type"] == "pattern_circular_ex"
        assert result["count"] == 6
        assert result["angle"] == 360.0
        model.Patterns.AddByCircularEx.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_pattern_circular_ex("Hole1", 6, 360.0, 0)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_face(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")
        faces = MagicMock()
        faces.Count = 1
        model.Body.Faces.return_value = faces

        result = feature_mgr.create_pattern_circular_ex("Hole1", 6, 360.0, 5)
        assert "error" in result
        assert "Invalid axis_face_index" in result["error"]


# ============================================================================
# BATCH 9: PATTERN DUPLICATE
# ============================================================================


class TestCreatePatternDuplicate:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Extrude1")
        pattern = MagicMock()
        pattern.Name = "Dup1"
        model.Patterns.AddDuplicate.return_value = pattern

        result = feature_mgr.create_pattern_duplicate("Extrude1")
        assert result["status"] == "created"
        assert result["type"] == "pattern_duplicate"
        assert result["feature"] == "Extrude1"
        model.Patterns.AddDuplicate.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_pattern_duplicate("Extrude1")
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_feature_not_found(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        feat = MagicMock()
        feat.Name = "OtherFeature"
        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = feature_mgr.create_pattern_duplicate("Missing")
        assert "error" in result
        assert "not found" in result["error"]


# ============================================================================
# BATCH 9: PATTERN BY FILL
# ============================================================================


class TestCreatePatternByFill:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")
        pattern = MagicMock()
        pattern.Name = "FillPattern1"
        model.Patterns.AddByFill.return_value = pattern

        result = feature_mgr.create_pattern_by_fill("Hole1", 0, 0.01, 0.01)
        assert result["status"] == "created"
        assert result["type"] == "pattern_by_fill"
        assert result["fill_region_face_index"] == 0
        model.Patterns.AddByFill.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_pattern_by_fill("Hole1", 0, 0.01, 0.01)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_face(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")
        faces = MagicMock()
        faces.Count = 1
        model.Body.Faces.return_value = faces

        result = feature_mgr.create_pattern_by_fill("Hole1", 5, 0.01, 0.01)
        assert "error" in result
        assert "Invalid fill_region_face_index" in result["error"]


# ============================================================================
# BATCH 9: PATTERN BY TABLE
# ============================================================================


class TestCreatePatternByTable:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")
        pattern = MagicMock()
        pattern.Name = "TablePattern1"
        model.Patterns.AddPatternByTable.return_value = pattern

        result = feature_mgr.create_pattern_by_table("Hole1", [0.01, 0.02], [0.01, 0.02])
        assert result["status"] == "created"
        assert result["type"] == "pattern_by_table"
        assert result["point_count"] == 2
        model.Patterns.AddPatternByTable.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_pattern_by_table("Hole1", [0.01], [0.01])
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_mismatched_offsets(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        _setup_feature_lookup(doc, "Hole1")
        result = feature_mgr.create_pattern_by_table("Hole1", [0.01, 0.02], [0.01])
        assert "error" in result
        assert "same length" in result["error"]


# ============================================================================
# PATTERN BY TABLE SYNC
# ============================================================================


class TestCreatePatternByTableSync:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")
        pattern = MagicMock()
        pattern.Name = "TablePatternSync1"
        model.Patterns.AddPatternByTableSync.return_value = pattern

        result = feature_mgr.create_pattern_by_table_sync("Hole1", [0.01, 0.02], [0.01, 0.02])
        assert result["status"] == "created"
        assert result["type"] == "pattern_by_table_sync"
        assert result["point_count"] == 2
        model.Patterns.AddPatternByTableSync.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_pattern_by_table_sync("Hole1", [0.01], [0.01])
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_mismatched_offsets(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        _setup_feature_lookup(doc, "Hole1")
        result = feature_mgr.create_pattern_by_table_sync("Hole1", [0.01, 0.02], [0.01])
        assert "error" in result
        assert "same length" in result["error"]


# ============================================================================
# PATTERN BY FILL EX
# ============================================================================


class TestCreatePatternByFillEx:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")
        body = model.Body
        faces = MagicMock()
        faces.Count = 3
        faces.Item.return_value = MagicMock()
        body.Faces.return_value = faces

        pattern = MagicMock()
        pattern.Name = "FillPatternEx1"
        model.Patterns.AddByFillEx.return_value = pattern

        result = feature_mgr.create_pattern_by_fill_ex("Hole1", 0, 0.01, 0.01)
        assert result["status"] == "created"
        assert result["type"] == "pattern_by_fill_ex"
        model.Patterns.AddByFillEx.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_pattern_by_fill_ex("Hole1", 0, 0.01, 0.01)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_face(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")
        body = model.Body
        faces = MagicMock()
        faces.Count = 1
        body.Faces.return_value = faces

        result = feature_mgr.create_pattern_by_fill_ex("Hole1", 5, 0.01, 0.01)
        assert "error" in result
        assert "Invalid fill_region_face_index" in result["error"]


# ============================================================================
# PATTERN BY CURVE EX
# ============================================================================


class TestCreatePatternByCurveEx:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")
        body = model.Body
        edges = MagicMock()
        edges.Count = 3
        edges.Item.return_value = MagicMock()
        body.Edges.return_value = edges

        pattern = MagicMock()
        pattern.Name = "CurvePatternEx1"
        model.Patterns.AddByCurveEx.return_value = pattern

        result = feature_mgr.create_pattern_by_curve_ex("Hole1", 0, 5, 0.01)
        assert result["status"] == "created"
        assert result["type"] == "pattern_by_curve_ex"
        assert result["count"] == 5
        model.Patterns.AddByCurveEx.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_pattern_by_curve_ex("Hole1", 0, 5, 0.01)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_edge(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")
        body = model.Body
        edges = MagicMock()
        edges.Count = 1
        body.Edges.return_value = edges

        result = feature_mgr.create_pattern_by_curve_ex("Hole1", 5, 5, 0.01)
        assert "error" in result
        assert "Invalid curve_edge_index" in result["error"]


# ============================================================================
# SAVE AS MIRROR PART
# ============================================================================


class TestSaveAsMirrorPart:
    def test_success(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        plane = MagicMock()
        ref_planes = MagicMock()
        ref_planes.Count = 3
        ref_planes.Item.return_value = plane
        doc.RefPlanes = ref_planes

        result = feature_mgr.save_as_mirror_part("C:/temp/mirror.par", 3, True)
        assert result["status"] == "saved"
        assert result["type"] == "mirror_part"
        assert result["path"] == "C:/temp/mirror.par"
        assert result["linked"] is True
        models.SaveAsMirrorPart.assert_called_once_with("C:/temp/mirror.par", plane, True)

    def test_no_model(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        models.Count = 0

        result = feature_mgr.save_as_mirror_part("C:/temp/mirror.par")
        assert "error" in result
        assert "No model" in result["error"]

    def test_invalid_plane_index(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = feature_mgr.save_as_mirror_part("C:/temp/mirror.par", 5)
        assert "error" in result
        assert "Invalid mirror_plane_index" in result["error"]

    def test_unlinked(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        plane = MagicMock()
        ref_planes = MagicMock()
        ref_planes.Count = 3
        ref_planes.Item.return_value = plane
        doc.RefPlanes = ref_planes

        result = feature_mgr.save_as_mirror_part("C:/temp/mirror.par", 2, False)
        assert result["status"] == "saved"
        assert result["linked"] is False
        models.SaveAsMirrorPart.assert_called_once_with("C:/temp/mirror.par", plane, False)


# ============================================================================
# USER DEFINED PATTERN
# ============================================================================


class TestCreateUserDefinedPattern:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, _ = managers
        seed = MagicMock()
        seed.Name = "Extrude 1"
        features = MagicMock()
        features.Count = 1
        features.Item.return_value = seed
        doc.DesignEdgebarFeatures = features

        p1, p2 = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2]

        udp = MagicMock()
        model.UserDefinedPatterns = udp

        result = feature_mgr.create_user_defined_pattern("Extrude 1")
        assert result["status"] == "created"
        assert result["type"] == "user_defined_pattern"
        assert result["seed_feature"] == "Extrude 1"
        assert result["num_occurrences"] == 2
        udp.AddByProfiles.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0

        result = feature_mgr.create_user_defined_pattern("Extrude 1")
        assert "error" in result
        assert "No model" in result["error"]

    def test_feature_not_found(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, _ = managers
        feat = MagicMock()
        feat.Name = "Other Feature"
        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = feature_mgr.create_user_defined_pattern("NonExistent")
        assert "error" in result
        assert "not found" in result["error"]

    def test_no_profiles(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, _ = managers
        seed = MagicMock()
        seed.Name = "Extrude 1"
        features = MagicMock()
        features.Count = 1
        features.Item.return_value = seed
        doc.DesignEdgebarFeatures = features

        sketch_mgr.get_accumulated_profiles.return_value = []

        result = feature_mgr.create_user_defined_pattern("Extrude 1")
        assert "error" in result
        assert "at least 1 profile" in result["error"]


# ============================================================================
# SLOT MULTI BODY
# ============================================================================


class TestCreateSlotMultiBody:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        model.Slots = MagicMock()

        result = feature_mgr.create_slot_multi_body(0.005, 0.01)
        assert result["status"] == "created"
        assert result["type"] == "slot_multi_body"
        assert result["width"] == 0.005
        assert result["depth"] == 0.01
        model.Slots.AddMultiBody.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_slot_multi_body(0.005, 0.01)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_slot_multi_body(0.005, 0.01)
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# SLOT SYNC MULTI BODY
# ============================================================================


class TestCreateSlotSyncMultiBody:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        model.Slots = MagicMock()

        result = feature_mgr.create_slot_sync_multi_body(0.005, 0.01)
        assert result["status"] == "created"
        assert result["type"] == "slot_sync_multi_body"
        assert result["width"] == 0.005
        assert result["depth"] == 0.01
        model.Slots.AddSyncMultiBody.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_slot_sync_multi_body(0.005, 0.01)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        model.Slots = MagicMock()

        result = feature_mgr.create_slot_sync_multi_body(0.005, 0.01, "Reverse")
        assert result["status"] == "created"
        assert result["direction"] == "Reverse"
