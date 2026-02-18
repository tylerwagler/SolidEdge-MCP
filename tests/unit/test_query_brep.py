"""
Unit tests for QueryManager backend methods (_brep.py mixin).

Uses unittest.mock to simulate COM objects.
"""

from unittest.mock import MagicMock

import pytest


@pytest.fixture
def doc_mgr():
    """Create mock doc_manager."""
    dm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc
    return dm, doc


@pytest.fixture
def query_mgr(doc_mgr):
    """Create QueryManager with mocked dependencies."""
    from solidedge_mcp.backends.query import QueryManager

    dm, doc = doc_mgr
    return QueryManager(dm), doc


def _setup_face_edge(doc):
    """Helper to set up mock doc with model > body > face > edge."""
    model = MagicMock()
    models = MagicMock()
    models.Count = 1
    models.Item.return_value = model
    doc.Models = models

    edge = MagicMock()
    edges = MagicMock()
    edges.Count = 3
    edges.Item.return_value = edge

    face = MagicMock()
    face.Edges = edges

    faces = MagicMock()
    faces.Count = 6
    faces.Item.return_value = face
    model.Body.Faces.return_value = faces

    return model, face, edge


def _setup_body(doc):
    """Helper to set up mock doc with model > body."""
    model = MagicMock()
    models = MagicMock()
    models.Count = 1
    models.Item.return_value = model
    doc.Models = models
    return model, model.Body


# ============================================================================
# TOPOLOGY QUERIES
# ============================================================================


class TestGetBodyFaces:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        # Set up model
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        # Set up faces
        face1 = MagicMock()
        face1.Type = 1
        face1.Area = 0.01
        face1.Edges.Count = 4

        face2 = MagicMock()
        face2.Type = 2
        face2.Area = 0.005
        face2.Edges.Count = 3

        faces = MagicMock()
        faces.Count = 2
        faces.Item.side_effect = lambda i: [None, face1, face2][i]
        model.Body.Faces.return_value = faces

        result = qm.get_body_faces()
        assert result["count"] == 2
        assert result["faces"][0]["area"] == 0.01
        assert result["faces"][0]["edge_count"] == 4


class TestGetBodyEdges:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        face = MagicMock()
        face.Edges.Count = 4

        faces = MagicMock()
        faces.Count = 1
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.get_body_edges()
        assert result["total_face_count"] == 1
        assert result["total_edge_references"] == 4


class TestGetFaceInfo:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        face = MagicMock()
        face.Type = 1
        face.Area = 0.01
        face.Edges.Count = 4
        face.Vertices.Count = 4

        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.get_face_info(0)
        assert result["index"] == 0
        assert result["area"] == 0.01

    def test_invalid_index(self, query_mgr):
        qm, doc = query_mgr

        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        faces = MagicMock()
        faces.Count = 6
        model.Body.Faces.return_value = faces

        result = qm.get_face_info(10)
        assert "error" in result


# ============================================================================
# GET EDGE INFO
# ============================================================================


class TestGetEdgeInfo:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        edge = MagicMock()
        edge.Length = 0.1
        edge.Type = 1

        edges = MagicMock()
        edges.Count = 3
        edges.Item.return_value = edge

        face = MagicMock()
        face.Edges = edges

        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.get_edge_info(0, 0)
        assert result["face_index"] == 0
        assert result["edge_index"] == 0
        assert result["length"] == 0.1

    def test_invalid_face(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        faces = MagicMock()
        faces.Count = 6
        model.Body.Faces.return_value = faces

        result = qm.get_edge_info(10, 0)
        assert "error" in result

    def test_invalid_edge(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        face = MagicMock()
        face.Edges.Count = 2

        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.get_edge_info(0, 5)
        assert "error" in result


# ============================================================================
# GET EDGE COUNT
# ============================================================================


class TestGetEdgeCount:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        face1 = MagicMock()
        face1.Edges.Count = 4

        face2 = MagicMock()
        face2.Edges.Count = 4

        faces = MagicMock()
        faces.Count = 2
        faces.Item.side_effect = lambda i: [None, face1, face2][i]
        model.Body.Faces.return_value = faces

        result = qm.get_edge_count()
        assert result["total_edge_references"] == 8
        assert result["face_count"] == 2

    def test_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_edge_count()
        assert "error" in result


# ============================================================================
# GET FACE COUNT
# ============================================================================


class TestGetFaceCount:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        faces = MagicMock()
        faces.Count = 6
        model.Body.Faces.return_value = faces

        result = qm.get_face_count()
        assert result["face_count"] == 6

    def test_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_face_count()
        assert "error" in result


# ============================================================================
# GET BODY VERTICES
# ============================================================================


class TestGetBodyVertices:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        _model, body = _setup_body(doc)

        v1 = MagicMock()
        v1.GetPointData.return_value = ((0.0, 0.0, 0.0),)

        v2 = MagicMock()
        v2.GetPointData.return_value = ((0.1, 0.0, 0.0),)

        vertices = MagicMock()
        vertices.Count = 2
        vertices.Item.side_effect = lambda i: [None, v1, v2][i]
        body.Vertices = vertices

        result = qm.get_body_vertices()
        assert result["vertex_count"] == 2
        assert result["vertices"][0]["point"] == [0.0, 0.0, 0.0]
        assert result["vertices"][1]["point"] == [0.1, 0.0, 0.0]

    def test_empty(self, query_mgr):
        qm, doc = query_mgr
        _model, body = _setup_body(doc)

        vertices = MagicMock()
        vertices.Count = 0
        body.Vertices = vertices

        result = qm.get_body_vertices()
        assert result["vertex_count"] == 0
        assert result["vertices"] == []

    def test_error_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_body_vertices()
        assert "error" in result


# ============================================================================
# GET VERTEX COUNT
# ============================================================================


class TestGetVertexCount:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        body = MagicMock()

        face1 = MagicMock()
        verts1 = MagicMock()
        verts1.Count = 4
        face1.Vertices = verts1

        face2 = MagicMock()
        verts2 = MagicMock()
        verts2.Count = 3
        face2.Vertices = verts2

        faces = MagicMock()
        faces.Count = 2
        faces.Item.side_effect = lambda i: {1: face1, 2: face2}[i]
        body.Faces.return_value = faces

        model.Body = body
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        result = qm.get_vertex_count()
        assert result["total_vertex_references"] == 7
        assert result["face_count"] == 2

    def test_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_vertex_count()
        assert "error" in result


# ============================================================================
# B-REP TOPOLOGY: GET FACE NORMAL
# ============================================================================


class TestGetFaceNormal:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        face = MagicMock()
        face.GetNormal.return_value = (None, (0.0, 0.0, 1.0))
        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.get_face_normal(0, 0.5, 0.5)
        assert "normal" in result
        assert result["normal"] == [0.0, 0.0, 1.0]

    def test_error_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_face_normal(0)
        assert "error" in result

    def test_invalid_face_index(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        faces = MagicMock()
        faces.Count = 2
        model.Body.Faces.return_value = faces

        result = qm.get_face_normal(10)
        assert "error" in result


# ============================================================================
# B-REP TOPOLOGY: GET FACE GEOMETRY
# ============================================================================


class TestGetFaceGeometry:
    def test_plane(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        geom = MagicMock()
        geom.GetPlaneData.return_value = ((0.0, 0.0, 0.0), (0.0, 0.0, 1.0))

        face = MagicMock()
        face.Geometry = geom
        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.get_face_geometry(0)
        assert result["geometry_type"] == "Plane"
        assert result["root_point"] == [0.0, 0.0, 0.0]
        assert result["normal"] == [0.0, 0.0, 1.0]

    def test_cylinder(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        geom = MagicMock()
        geom.GetPlaneData.side_effect = Exception("Not a plane")
        geom.GetCylinderData.return_value = ((0.0, 0.0, 0.0), (0.0, 0.0, 1.0), 0.05)

        face = MagicMock()
        face.Geometry = geom
        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.get_face_geometry(0)
        assert result["geometry_type"] == "Cylinder"
        assert result["radius"] == 0.05

    def test_error_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_face_geometry(0)
        assert "error" in result


# ============================================================================
# B-REP TOPOLOGY: GET FACE LOOPS
# ============================================================================


class TestGetFaceLoops:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        loop1 = MagicMock()
        loop1.IsOuterLoop = True
        loop1.Edges.Count = 4

        loop2 = MagicMock()
        loop2.IsOuterLoop = False
        loop2.Edges.Count = 1

        loops = MagicMock()
        loops.Count = 2
        loops.Item.side_effect = lambda i: [None, loop1, loop2][i]

        face = MagicMock()
        face.Loops = loops
        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.get_face_loops(0)
        assert result["loop_count"] == 2
        assert result["loops"][0]["is_outer"] is True
        assert result["loops"][0]["edge_count"] == 4
        assert result["loops"][1]["is_outer"] is False

    def test_error_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_face_loops(0)
        assert "error" in result

    def test_invalid_face_index(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        faces = MagicMock()
        faces.Count = 2
        model.Body.Faces.return_value = faces

        result = qm.get_face_loops(10)
        assert "error" in result


# ============================================================================
# B-REP TOPOLOGY: GET FACE CURVATURE
# ============================================================================


class TestGetFaceCurvature:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        face = MagicMock()
        face.GetCurvatures.return_value = ((1.0, 0.0, 0.0), (20.0,), (0.0,))

        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.get_face_curvature(0, 0.5, 0.5)
        assert result["max_curvature"] == 20.0
        assert result["min_curvature"] == 0.0
        assert result["max_tangent"] == [1.0, 0.0, 0.0]

    def test_error_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_face_curvature(0)
        assert "error" in result

    def test_invalid_face_index(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        faces = MagicMock()
        faces.Count = 2
        model.Body.Faces.return_value = faces

        result = qm.get_face_curvature(10)
        assert "error" in result


# ============================================================================
# B-REP TOPOLOGY: GET EDGE ENDPOINTS
# ============================================================================


class TestGetEdgeEndpoints:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        _model, _face, edge = _setup_face_edge(doc)
        edge.GetEndPoints.return_value = ((0.0, 0.0, 0.0), (0.1, 0.0, 0.0))

        result = qm.get_edge_endpoints(0, 0)
        assert "start" in result
        assert "end" in result
        assert result["start"] == [0.0, 0.0, 0.0]
        assert result["end"] == [0.1, 0.0, 0.0]

    def test_error_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_edge_endpoints(0, 0)
        assert "error" in result

    def test_invalid_face_index(self, query_mgr):
        qm, doc = query_mgr
        _model, _face, _edge = _setup_face_edge(doc)

        result = qm.get_edge_endpoints(99, 0)
        assert "error" in result


# ============================================================================
# B-REP TOPOLOGY: GET EDGE LENGTH
# ============================================================================


class TestGetEdgeLength:
    def test_success_via_property(self, query_mgr):
        qm, doc = query_mgr
        _model, _face, edge = _setup_face_edge(doc)
        edge.Length = 0.1

        result = qm.get_edge_length(0, 0)
        assert result["length"] == 0.1
        assert result["length_mm"] == 100.0

    def test_error_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_edge_length(0, 0)
        assert "error" in result

    def test_invalid_edge_index(self, query_mgr):
        qm, doc = query_mgr
        _model, _face, _edge = _setup_face_edge(doc)

        result = qm.get_edge_length(0, 99)
        assert "error" in result


# ============================================================================
# B-REP TOPOLOGY: GET EDGE TANGENT
# ============================================================================


class TestGetEdgeTangent:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        _model, _face, edge = _setup_face_edge(doc)
        edge.GetTangent.return_value = (None, (1.0, 0.0, 0.0))

        result = qm.get_edge_tangent(0, 0, 0.5)
        assert "tangent" in result
        assert result["tangent"] == [1.0, 0.0, 0.0]
        assert result["param"] == 0.5

    def test_error_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_edge_tangent(0, 0)
        assert "error" in result

    def test_invalid_face_index(self, query_mgr):
        qm, doc = query_mgr
        _model, _face, _edge = _setup_face_edge(doc)

        result = qm.get_edge_tangent(99, 0)
        assert "error" in result


# ============================================================================
# B-REP TOPOLOGY: GET EDGE GEOMETRY
# ============================================================================


class TestGetEdgeGeometry:
    def test_circle(self, query_mgr):
        qm, doc = query_mgr
        _model, _face, edge = _setup_face_edge(doc)
        geom = MagicMock()
        geom.GetCircleData.return_value = ((0.0, 0.0, 0.0), (0.0, 0.0, 1.0), 0.05)
        edge.Geometry = geom

        result = qm.get_edge_geometry(0, 0)
        assert result["geometry_type"] == "Circle"
        assert result["radius"] == 0.05

    def test_line_fallback(self, query_mgr):
        qm, doc = query_mgr
        _model, _face, edge = _setup_face_edge(doc)
        geom = MagicMock()
        geom.GetCircleData.side_effect = Exception("Not a circle")
        geom.GetEllipseData.side_effect = Exception("Not an ellipse")
        geom.GetBSplineInfo.side_effect = Exception("Not a BSpline")
        geom.Type = 0
        edge.Geometry = geom

        result = qm.get_edge_geometry(0, 0)
        assert result["geometry_type"] == "Line"

    def test_error_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_edge_geometry(0, 0)
        assert "error" in result


# ============================================================================
# B-REP TOPOLOGY: GET EDGE CURVATURE
# ============================================================================


class TestGetEdgeCurvature:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        _model, _face, edge = _setup_face_edge(doc)
        edge.GetCurvature.return_value = ((0.0, 1.0, 0.0), (20.0,))

        result = qm.get_edge_curvature(0, 0, 0.5)
        assert result["curvature"] == 20.0
        assert result["direction"] == [0.0, 1.0, 0.0]
        assert result["param"] == 0.5

    def test_error_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_edge_curvature(0, 0)
        assert "error" in result

    def test_invalid_edge_index(self, query_mgr):
        qm, doc = query_mgr
        _model, _face, _edge = _setup_face_edge(doc)

        result = qm.get_edge_curvature(0, 99)
        assert "error" in result


# ============================================================================
# B-REP TOPOLOGY: GET VERTEX POINT
# ============================================================================


class TestGetVertexPoint:
    def test_start_vertex(self, query_mgr):
        qm, doc = query_mgr
        _model, _face, edge = _setup_face_edge(doc)

        vertex = MagicMock()
        vertex.GetPointData.return_value = ((0.1, 0.2, 0.3),)
        vertex.ID = 42
        edge.StartVertex = vertex

        result = qm.get_vertex_point(0, 0, "start")
        assert result["point"] == [0.1, 0.2, 0.3]
        assert result["vertex_id"] == 42
        assert result["which"] == "start"

    def test_end_vertex(self, query_mgr):
        qm, doc = query_mgr
        _model, _face, edge = _setup_face_edge(doc)

        vertex = MagicMock()
        vertex.GetPointData.return_value = ((0.5, 0.6, 0.7),)
        vertex.ID = 99
        edge.EndVertex = vertex

        result = qm.get_vertex_point(0, 0, "end")
        assert result["point"] == [0.5, 0.6, 0.7]
        assert result["which"] == "end"

    def test_invalid_which(self, query_mgr):
        qm, doc = query_mgr
        _model, _face, _edge = _setup_face_edge(doc)

        result = qm.get_vertex_point(0, 0, "middle")
        assert "error" in result
        assert "middle" in result["error"]


# ============================================================================
# B-REP TOPOLOGY: GET SHELL INFO
# ============================================================================


class TestGetShellInfo:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        _model, body = _setup_body(doc)

        shell = MagicMock()
        shell.IsClosed = True
        shell.Volume = 0.001
        shell.IsVoid = False
        shell.Faces.Count = 6
        shell.Edges.Count = 12

        shells = MagicMock()
        shells.Count = 1
        shells.Item.return_value = shell
        body.Shells = shells

        result = qm.get_shell_info(0)
        assert result["is_closed"] is True
        assert result["volume"] == 0.001
        assert result["is_void"] is False
        assert result["face_count"] == 6
        assert result["edge_count"] == 12

    def test_invalid_index(self, query_mgr):
        qm, doc = query_mgr
        _model, body = _setup_body(doc)

        shells = MagicMock()
        shells.Count = 1
        body.Shells = shells

        result = qm.get_shell_info(5)
        assert "error" in result

    def test_error_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_shell_info(0)
        assert "error" in result


# ============================================================================
# B-REP TOPOLOGY: GET BODY SHELLS
# ============================================================================


class TestGetBodyShells:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        _model, body = _setup_body(doc)

        shell1 = MagicMock()
        shell1.IsClosed = True
        shell1.Volume = 0.001

        shell2 = MagicMock()
        shell2.IsClosed = False
        shell2.Volume = 0.0

        shells = MagicMock()
        shells.Count = 2
        shells.Item.side_effect = lambda i: [None, shell1, shell2][i]
        body.Shells = shells

        result = qm.get_body_shells()
        assert result["shell_count"] == 2
        assert result["shells"][0]["is_closed"] is True
        assert result["shells"][0]["volume"] == 0.001
        assert result["shells"][1]["is_closed"] is False

    def test_empty(self, query_mgr):
        qm, doc = query_mgr
        _model, body = _setup_body(doc)

        shells = MagicMock()
        shells.Count = 0
        body.Shells = shells

        result = qm.get_body_shells()
        assert result["shell_count"] == 0
        assert result["shells"] == []

    def test_error_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_body_shells()
        assert "error" in result


# ============================================================================
# B-REP TOPOLOGY: GET BSPLINE CURVE INFO
# ============================================================================


class TestGetBSplineCurveInfo:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        _model, _face, edge = _setup_face_edge(doc)

        geom = MagicMock()
        geom.GetBSplineInfo.return_value = (4, 10, 14, True, False, False, True)
        edge.Geometry = geom

        result = qm.get_bspline_curve_info(0, 0)
        assert result["order"] == 4
        assert result["num_poles"] == 10
        assert result["num_knots"] == 14
        assert result["rational"] is True
        assert result["closed"] is False
        assert result["planar"] is True

    def test_not_bspline(self, query_mgr):
        qm, doc = query_mgr
        _model, _face, edge = _setup_face_edge(doc)

        geom = MagicMock()
        geom.GetBSplineInfo.side_effect = Exception("Not a BSpline curve")
        edge.Geometry = geom

        result = qm.get_bspline_curve_info(0, 0)
        assert "error" in result

    def test_invalid_edge_index(self, query_mgr):
        qm, doc = query_mgr
        _model, _face, _edge = _setup_face_edge(doc)

        result = qm.get_bspline_curve_info(0, 99)
        assert "error" in result


# ============================================================================
# B-REP TOPOLOGY: GET BSPLINE SURFACE INFO
# ============================================================================


class TestGetBSplineSurfaceInfo:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        geom = MagicMock()
        geom.GetBSplineInfo.return_value = (4, 4, 8, 8, 12, 12, True, False)

        face = MagicMock()
        face.Geometry = geom
        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.get_bspline_surface_info(0)
        assert result["order"] == [4, 4]
        assert result["num_poles"] == [8, 8]
        assert result["num_knots"] == [12, 12]
        assert result["rational"] is True

    def test_not_bspline(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        geom = MagicMock()
        geom.GetBSplineInfo.side_effect = Exception("Not a BSpline surface")

        face = MagicMock()
        face.Geometry = geom
        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.get_bspline_surface_info(0)
        assert "error" in result

    def test_invalid_face_index(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        faces = MagicMock()
        faces.Count = 2
        model.Body.Faces.return_value = faces

        result = qm.get_bspline_surface_info(10)
        assert "error" in result


# ============================================================================
# B-REP TOPOLOGY: GET FACES BY RAY
# ============================================================================


class TestGetFacesByRay:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        _model, body = _setup_body(doc)

        hit_face = MagicMock()
        hit_face.Area = 0.01
        hit_face.ID = 5

        hit_faces = MagicMock()
        hit_faces.Count = 1
        hit_faces.Item.return_value = hit_face
        body.FacesByRay.return_value = hit_faces

        result = qm.get_faces_by_ray(0, 0, 0, 0, 0, 1)
        assert result["face_count"] == 1
        assert result["faces"][0]["area"] == 0.01
        assert result["faces"][0]["id"] == 5

    def test_no_hits(self, query_mgr):
        qm, doc = query_mgr
        _model, body = _setup_body(doc)

        hit_faces = MagicMock()
        hit_faces.Count = 0
        body.FacesByRay.return_value = hit_faces

        result = qm.get_faces_by_ray(0, 0, 0, 1, 0, 0)
        assert result["face_count"] == 0

    def test_error_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_faces_by_ray(0, 0, 0, 0, 0, 1)
        assert "error" in result


# ============================================================================
# B-REP TOPOLOGY: IS POINT INSIDE BODY
# ============================================================================


class TestIsPointInsideBody:
    def test_inside(self, query_mgr):
        qm, doc = query_mgr
        _model, body = _setup_body(doc)

        shell = MagicMock()
        shell.IsPointInside.return_value = True

        shells = MagicMock()
        shells.Count = 1
        shells.Item.return_value = shell
        body.Shells = shells

        result = qm.is_point_inside_body(0.05, 0.05, 0.025)
        assert result["is_inside"] is True
        assert result["point"] == [0.05, 0.05, 0.025]

    def test_outside(self, query_mgr):
        qm, doc = query_mgr
        _model, body = _setup_body(doc)

        shell = MagicMock()
        shell.IsPointInside.return_value = False

        shells = MagicMock()
        shells.Count = 1
        shells.Item.return_value = shell
        body.Shells = shells

        result = qm.is_point_inside_body(10.0, 10.0, 10.0)
        assert result["is_inside"] is False

    def test_no_shells(self, query_mgr):
        qm, doc = query_mgr
        _model, body = _setup_body(doc)

        shells = MagicMock()
        shells.Count = 0
        body.Shells = shells

        result = qm.is_point_inside_body(0, 0, 0)
        assert "error" in result


# ============================================================================
# B-REP TOPOLOGY: GET BODY EXTREME POINT
# ============================================================================


class TestGetBodyExtremePoint:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        _model, body = _setup_body(doc)
        body.GetExtremePoint.return_value = (0.1, 0.2, 0.3)

        result = qm.get_body_extreme_point(1.0, 0.0, 0.0)
        assert result["extreme_point"] == [0.1, 0.2, 0.3]
        assert result["direction"] == [1.0, 0.0, 0.0]

    def test_error_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_body_extreme_point(1.0, 0.0, 0.0)
        assert "error" in result

    def test_negative_direction(self, query_mgr):
        qm, doc = query_mgr
        _model, body = _setup_body(doc)
        body.GetExtremePoint.return_value = (-0.05, 0.0, 0.0)

        result = qm.get_body_extreme_point(-1.0, 0.0, 0.0)
        assert result["extreme_point"] == [-0.05, 0.0, 0.0]
        assert result["direction"] == [-1.0, 0.0, 0.0]


# ============================================================================
# SET FACE COLOR
# ============================================================================


class TestSetFaceColor:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        face = MagicMock()
        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.set_face_color(0, 255, 0, 0)
        assert result["status"] == "updated"
        assert result["color"] == [255, 0, 0]

    def test_invalid_face(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        faces = MagicMock()
        faces.Count = 2
        model.Body.Faces.return_value = faces

        result = qm.set_face_color(5, 0, 0, 255)
        assert "error" in result
