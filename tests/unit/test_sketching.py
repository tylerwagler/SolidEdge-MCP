"""
Unit tests for SketchManager backend methods.

Uses unittest.mock to simulate COM objects so tests run without Solid Edge.
"""

from unittest.mock import MagicMock

import pytest


@pytest.fixture
def sketch_mgr():
    """Create SketchManager with mocked dependencies."""
    from solidedge_mcp.backends.sketching import SketchManager

    dm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc
    sm = SketchManager(dm)
    return sm, doc


# ============================================================================
# PROJECT REF PLANE
# ============================================================================


class TestProjectRefPlane:
    def test_success(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        ref_plane = MagicMock()
        ref_planes = MagicMock()
        ref_planes.Count = 3
        ref_planes.Item.return_value = ref_plane
        doc.RefPlanes = ref_planes

        projected = MagicMock()
        profile.ProjectRefPlane.return_value = projected

        result = sm.project_ref_plane(1)
        assert result["status"] == "projected"
        assert result["plane_index"] == 1
        profile.ProjectRefPlane.assert_called_once_with(ref_plane)

    def test_no_active_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        sm.active_profile = None

        result = sm.project_ref_plane(1)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_invalid_plane_index(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = sm.project_ref_plane(10)
        assert "error" in result
        assert "plane_index" in result["error"]

    def test_plane_index_zero(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = sm.project_ref_plane(0)
        assert "error" in result


# ============================================================================
# OFFSET SKETCH 2D
# ============================================================================


class TestOffsetSketch2d:
    def test_success(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        result = sm.offset_sketch_2d(1.0, 0.0, 0.005)
        assert result["status"] == "offset"
        assert result["offset_side_x"] == 1.0
        assert result["offset_side_y"] == 0.0
        assert result["offset_distance"] == 0.005
        profile.Offset2d.assert_called_once_with(1.0, 0.0, 0.005)

    def test_no_active_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        sm.active_profile = None

        result = sm.offset_sketch_2d(1.0, 0.0, 0.005)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_negative_values(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        result = sm.offset_sketch_2d(-1.0, -1.0, 0.01)
        assert result["status"] == "offset"
        profile.Offset2d.assert_called_once_with(-1.0, -1.0, 0.01)


# ============================================================================
# SKETCH ROTATE
# ============================================================================


class TestSketchRotate:
    def test_success(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        lines = MagicMock()
        lines.Count = 0
        profile.Lines2d = lines

        circles = MagicMock()
        circles.Count = 0
        profile.Circles2d = circles

        result = sm.sketch_rotate(0.0, 0.0, 90.0)
        assert result["status"] == "rotated"
        assert result["angle_degrees"] == 90.0

    def test_no_active_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        sm.active_profile = None

        result = sm.sketch_rotate(0.0, 0.0, 45.0)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_with_lines(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        line = MagicMock()
        line.StartPoint.X = 0.1
        line.StartPoint.Y = 0.0
        line.EndPoint.X = 0.1
        line.EndPoint.Y = 0.1
        lines = MagicMock()
        lines.Count = 1
        lines.Item.return_value = line
        profile.Lines2d = lines

        circles = MagicMock()
        circles.Count = 0
        profile.Circles2d = circles

        result = sm.sketch_rotate(0.0, 0.0, 90.0)
        assert result["status"] == "rotated"
        assert result["elements_rotated"] == 1


# ============================================================================
# SKETCH SCALE
# ============================================================================


class TestSketchScale:
    def test_success(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        lines = MagicMock()
        lines.Count = 0
        profile.Lines2d = lines

        circles = MagicMock()
        circles.Count = 0
        profile.Circles2d = circles

        result = sm.sketch_scale(0.0, 0.0, 2.0)
        assert result["status"] == "scaled"
        assert result["scale_factor"] == 2.0

    def test_no_active_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        sm.active_profile = None

        result = sm.sketch_scale(0.0, 0.0, 2.0)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_with_circles(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        lines = MagicMock()
        lines.Count = 0
        profile.Lines2d = lines

        circle = MagicMock()
        circle.CenterPoint.X = 0.05
        circle.CenterPoint.Y = 0.0
        circle.Radius = 0.01
        circles = MagicMock()
        circles.Count = 1
        circles.Item.return_value = circle
        profile.Circles2d = circles

        result = sm.sketch_scale(0.0, 0.0, 2.0)
        assert result["status"] == "scaled"
        assert result["elements_scaled"] == 1


# ============================================================================
# GET SKETCH MATRIX
# ============================================================================


class TestGetSketchMatrix:
    def test_success(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        profile.GetMatrix.return_value = tuple(float(i) for i in range(16))
        result = sm.get_sketch_matrix()
        assert result["status"] == "ok"
        assert len(result["matrix"]) == 16

    def test_no_active_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        sm.active_profile = None

        result = sm.get_sketch_matrix()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_com_error(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile
        profile.GetMatrix.side_effect = Exception("COM error")

        result = sm.get_sketch_matrix()
        assert "error" in result


# ============================================================================
# CLEAN SKETCH GEOMETRY
# ============================================================================


class TestCleanSketchGeometry:
    def test_success(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        result = sm.clean_sketch_geometry()
        assert result["status"] == "cleaned"
        profile.CleanGeometry2d.assert_called_once()

    def test_no_active_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        sm.active_profile = None

        result = sm.clean_sketch_geometry()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_custom_params(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        result = sm.clean_sketch_geometry(
            clean_points=False,
            clean_splines=False,
            clean_identical=True,
            clean_small=True,
            small_tolerance=0.001,
        )
        assert result["status"] == "cleaned"
        profile.CleanGeometry2d.assert_called_once_with(0, False, False, True, True, None, 0.001)


# ============================================================================
# PROJECT SILHOUETTE EDGES
# ============================================================================


class TestProjectSilhouetteEdges:
    def test_success(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        result = sm.project_silhouette_edges()
        assert result["status"] == "projected"
        assert result["type"] == "silhouette_edges"
        profile.ProjectSilhouetteEdges.assert_called_once()

    def test_no_active_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        sm.active_profile = None

        result = sm.project_silhouette_edges()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_com_error(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile
        profile.ProjectSilhouetteEdges.side_effect = Exception("COM error")

        result = sm.project_silhouette_edges()
        assert "error" in result
        assert "traceback" in result


# ============================================================================
# INCLUDE REGION FACES
# ============================================================================


class TestIncludeRegionFaces:
    def test_success(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        body = MagicMock()
        model = MagicMock()
        model.Body = body
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        face1 = MagicMock()
        face2 = MagicMock()
        faces = MagicMock()
        faces.Count = 3
        faces.Item.side_effect = lambda i: {1: face1, 2: face2, 3: MagicMock()}[i]
        body.Faces.return_value = faces

        result = sm.include_region_faces([0, 1])
        assert result["status"] == "included"
        assert result["face_count"] == 2
        profile.IncludeRegionFaces.assert_called_once()

    def test_no_active_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        sm.active_profile = None

        result = sm.include_region_faces([0])
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_empty_face_indices(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        result = sm.include_region_faces([])
        assert "error" in result
        assert "No face indices" in result["error"]


# ============================================================================
# CHAIN LOCATE
# ============================================================================


class TestChainLocate:
    def test_success(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile
        profile.ChainLocate.return_value = MagicMock()

        result = sm.chain_locate(0.05, 0.05)
        assert result["status"] == "located"
        assert result["type"] == "chain"
        assert result["x"] == 0.05
        assert result["y"] == 0.05
        profile.ChainLocate.assert_called_once_with(0.05, 0.05, 0.001)

    def test_no_active_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        sm.active_profile = None

        result = sm.chain_locate(0.0, 0.0)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_custom_tolerance(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile
        profile.ChainLocate.return_value = MagicMock()

        result = sm.chain_locate(0.1, 0.2, tolerance=0.005)
        assert result["status"] == "located"
        assert result["tolerance"] == 0.005
        profile.ChainLocate.assert_called_once_with(0.1, 0.2, 0.005)


# ============================================================================
# CONVERT TO CURVE
# ============================================================================


class TestConvertToCurve:
    def test_success(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile
        profile.ConvertToCurve.return_value = MagicMock()

        result = sm.convert_to_curve()
        assert result["status"] == "converted"
        assert result["type"] == "curve"
        profile.ConvertToCurve.assert_called_once()

    def test_no_active_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        sm.active_profile = None

        result = sm.convert_to_curve()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_com_error(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile
        profile.ConvertToCurve.side_effect = Exception("COM error")

        result = sm.convert_to_curve()
        assert "error" in result
        assert "traceback" in result


# ============================================================================
# SKETCH PASTE
# ============================================================================


class TestSketchPaste:
    def test_success(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        result = sm.sketch_paste()
        assert result["status"] == "pasted"
        assert result["type"] == "sketch_paste"
        profile.Paste.assert_called_once()

    def test_no_active_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        sm.active_profile = None

        result = sm.sketch_paste()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_com_error(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile
        profile.Paste.side_effect = Exception("Clipboard empty")

        result = sm.sketch_paste()
        assert "error" in result
        assert "traceback" in result


# ============================================================================
# GET ORDERED GEOMETRY
# ============================================================================


class TestGetOrderedGeometry:
    def test_success(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        elem1 = MagicMock()
        elem1.StartPoint.X = 0.0
        elem1.StartPoint.Y = 0.0
        elem1.EndPoint.X = 0.1
        elem1.EndPoint.Y = 0.0
        elem1.Length = 0.1
        type(elem1).__name__ = "Line2d"

        elem2 = MagicMock()
        elem2.CenterPoint.X = 0.05
        elem2.CenterPoint.Y = 0.05
        elem2.Radius = 0.02
        type(elem2).__name__ = "Circle2d"

        profile.OrderedGeometry.return_value = (2, [elem1, elem2])

        result = sm.get_ordered_geometry()
        assert result["status"] == "ok"
        assert result["num_elements"] == 2
        assert len(result["elements"]) == 2
        assert result["elements"][0]["start_x"] == 0.0
        assert result["elements"][0]["end_x"] == 0.1
        assert result["elements"][1]["center_x"] == 0.05
        assert result["elements"][1]["radius"] == 0.02

    def test_no_active_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        sm.active_profile = None

        result = sm.get_ordered_geometry()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_empty_geometry(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile
        profile.OrderedGeometry.return_value = (0, [])

        result = sm.get_ordered_geometry()
        assert result["status"] == "ok"
        assert result["num_elements"] == 0
        assert result["elements"] == []
