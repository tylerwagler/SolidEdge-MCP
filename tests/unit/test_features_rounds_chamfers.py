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
# VARIABLE ROUND
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
# BLEND
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
# CREATE CHAMFER UNEQUAL ON FACE
# ============================================================================


class TestCreateChamferUnequalOnFace:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_chamfer_unequal_on_face(0.002, 0.004, 0)
        assert result["status"] == "created"
        assert result["type"] == "chamfer_unequal_on_face"
        assert result["distance1"] == 0.002
        assert result["distance2"] == 0.004
        model.Chamfers.AddUnequalSetback.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_chamfer_unequal_on_face(0.002, 0.004, 0)
        assert "error" in result
        assert "No features" in result["error"]

    def test_invalid_face_index(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        body = model.Body
        faces = body.Faces.return_value
        faces.Count = 1

        result = feature_mgr.create_chamfer_unequal_on_face(0.002, 0.004, 5)
        assert "error" in result
        assert "face index" in result["error"]


# ============================================================================
# ROUND BLEND
# ============================================================================


class TestCreateRoundBlend:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_round_blend(0, 0, 0.002)
        assert result["status"] == "created"
        assert result["type"] == "round_blend"
        model.Rounds.AddBlend.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_round_blend(0, 0, 0.002)
        assert "error" in result

    def test_invalid_face_index(self, feature_mgr, managers):
        _, _, _, _, _, _ = managers
        result = feature_mgr.create_round_blend(99, 0, 0.002)
        assert "error" in result
        assert "Invalid face_index1" in result["error"]


# ============================================================================
# ROUND SURFACE BLEND
# ============================================================================


class TestCreateRoundSurfaceBlend:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_round_surface_blend(0, 0, 0.002)
        assert result["status"] == "created"
        assert result["type"] == "round_surface_blend"
        model.Rounds.AddSurfaceBlend.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_round_surface_blend(0, 0, 0.002)
        assert "error" in result

    def test_invalid_face_index(self, feature_mgr, managers):
        _, _, _, _, _, _ = managers
        result = feature_mgr.create_round_surface_blend(0, 99, 0.002)
        assert "error" in result
        assert "Invalid face_index2" in result["error"]


# ============================================================================
# BLEND VARIABLE
# ============================================================================


class TestCreateBlendVariable:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_blend_variable(0.001, 0.003)
        assert result["status"] == "created"
        assert result["type"] == "blend_variable"
        assert result["radius1"] == 0.001
        assert result["radius2"] == 0.003
        model.Blends.AddVariable.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_blend_variable(0.001, 0.003)
        assert "error" in result

    def test_no_faces(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        faces = MagicMock()
        faces.Count = 0
        model.Body.Faces.return_value = faces
        result = feature_mgr.create_blend_variable(0.001, 0.003)
        assert "error" in result
        assert "No faces" in result["error"]


# ============================================================================
# BLEND SURFACE
# ============================================================================


class TestCreateBlendSurface:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_blend_surface(0, 0)
        assert result["status"] == "created"
        assert result["type"] == "blend_surface"
        model.Blends.AddSurfaceBlend.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_blend_surface(0, 0)
        assert "error" in result

    def test_invalid_face_index(self, feature_mgr, managers):
        _, _, _, _, _, _ = managers
        result = feature_mgr.create_blend_surface(0, 99)
        assert "error" in result
        assert "Invalid face_index2" in result["error"]
