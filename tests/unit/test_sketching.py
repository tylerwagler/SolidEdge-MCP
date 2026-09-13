"""
Unit tests for SketchManager backend methods.

Uses unittest.mock to simulate COM objects so tests run without Solid Edge.
"""

from unittest.mock import MagicMock

import pytest
import pythoncom
from win32com.client import VARIANT

# --- COM argument assertions -------------------------------------------------
# win32com's VARIANT has no __eq__, so assert_called_once_with cannot compare
# [in,out] SAFEARRAY buffers directly. Normalise to (varianttype, value).

VT_R8_ARRAY = pythoncom.VT_ARRAY | pythoncom.VT_R8
VT_DISPATCH_ARRAY = pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH


def _shape(arg):
    if isinstance(arg, VARIANT):
        return (arg.varianttype, list(arg.value))
    return arg


def call_shape(mock_method):
    """(positional_args, keyword_args) of the single call, VARIANTs normalised."""
    assert mock_method.call_count == 1, mock_method.call_args_list
    args, kwargs = mock_method.call_args
    return tuple(_shape(a) for a in args), {k: _shape(v) for k, v in kwargs.items()}


def _collection(items):
    coll = MagicMock()
    coll.Count = len(items)
    coll.Item.side_effect = lambda i: items[i - 1]
    return coll


def _profile_with_geometry(lines=(), arcs=(), circles=(), splines=()):
    """A Profile whose 2D geometry collections hold exactly these elements."""
    profile = MagicMock()
    profile.Lines2d = _collection(list(lines))
    profile.Arcs2d = _collection(list(arcs))
    profile.Circles2d = _collection(list(circles))
    profile.Ellipses2d = _collection([])
    profile.EllipticalArcs2d = _collection([])
    profile.BSplineCurves2d = _collection(list(splines))
    profile.Conics2d = _collection([])
    return profile


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
# CLOSE SKETCH  (validation-flag + honest validation-code reporting)
# ============================================================================


class TestCloseSketch:
    def _profile(self, sm, end_ret=0):
        profile = MagicMock()
        profile.End.return_value = end_ret
        sm.active_profile = profile
        return profile

    def test_no_active_sketch(self, sketch_mgr):
        sm, _ = sketch_mgr
        sm.active_profile = None
        assert "error" in sm.close_sketch()

    def test_closed_true_uses_igProfileClosed(self, sketch_mgr):
        sm, _ = sketch_mgr
        profile = self._profile(sm, end_ret=0)
        result = sm.close_sketch(closed=True)
        profile.End.assert_called_once_with(1)  # igProfileClosed
        assert result["status"] == "closed"
        assert result["validation_code"] == 0
        assert "note" not in result  # clean close -> no hint

    def test_closed_false_uses_igProfileDefault(self, sketch_mgr):
        sm, _ = sketch_mgr
        profile = self._profile(sm, end_ret=0)
        sm.close_sketch(closed=False)
        profile.End.assert_called_once_with(0)  # igProfileDefault

    def test_refaxis_forces_revolve_flag(self, sketch_mgr):
        sm, _ = sketch_mgr
        profile = self._profile(sm, end_ret=0)
        sm.active_refaxis = MagicMock()
        sm.close_sketch(closed=True)
        profile.End.assert_called_once_with(17)  # igProfileForRevolve

    def test_nonzero_code_is_informational_not_failure(self, sketch_mgr):
        # End() returns -103 for a polyline whose endpoints igProfileClosed
        # auto-connects -- it still extrudes, so this must NOT be flagged as a
        # failure. We surface the code as a hint only.
        sm, _ = sketch_mgr
        self._profile(sm, end_ret=-103)
        result = sm.close_sketch(closed=True)
        assert result["status"] == "closed"
        assert result["validation_code"] == -103
        assert "note" in result
        assert "valid" not in result

    def test_end_exception_surfaced(self, sketch_mgr):
        sm, _ = sketch_mgr
        profile = MagicMock()
        profile.End.side_effect = Exception("COM boom")
        sm.active_profile = profile
        result = sm.close_sketch()
        assert "error" in result


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

        # Profile.GetMatrix(Matrix SAFEARRAY(VT_R8)* [in,out]) - a 4x4 buffer,
        # passed as a plain list (a VARIANT wrapper is rejected by pywin32).
        assert call_shape(profile.GetMatrix) == (
            ([0.0] * 16,),
            {},
        )

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
    def test_reports_unsupported_without_calling_com(self, sketch_mgr):
        """Profile.CleanGeometry2d rejects every documented argument form.

        Five, six and seven arguments, with the layer as None, 0 and Missing,
        all return 0x80070057 E_INVALIDARG on Solid Edge 2026. The options
        argument is a CleanProfileOptions value the type library does not
        enumerate, so there is nothing left to try.
        """
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        result = sm.clean_sketch_geometry()
        assert result["unsupported"] is True
        assert "Clean Geometry" in result["error"]
        profile.CleanGeometry2d.assert_not_called()

    def test_no_active_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        sm.active_profile = None

        result = sm.clean_sketch_geometry()
        assert "error" in result
        assert "No active sketch" in result["error"]


# ============================================================================
# PROJECT SILHOUETTE EDGES
# ============================================================================


class TestProjectSilhouetteEdges:
    """ProjectSilhouetteEdges(FaceToProject, Geometry2dCount, Geometry2d):
    the face cannot be selected from here, so the call is refused."""

    def test_unsupported_without_calling_com(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile

        result = sm.project_silhouette_edges()
        assert result["unsupported"] is True
        assert "cannot select" in result["error"]
        profile.ProjectSilhouetteEdges.assert_not_called()

    def test_no_active_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        sm.active_profile = None

        result = sm.project_silhouette_edges()
        assert "error" in result
        assert "No active sketch" in result["error"]
        assert "unsupported" not in result


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

        # IncludeRegionFaces(NumberOfRegionFaces, RegionFaces SAFEARRAY)
        args, kwargs = profile.IncludeRegionFaces.call_args
        assert kwargs == {}
        assert args[0] == 2
        assert list(args[1]) == [face1, face2]

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
        # Profile.ChainLocate(x, y) - tolerance is not part of the API.
        profile.ChainLocate.assert_called_once_with(0.05, 0.05)

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
        assert result["tolerance"] == 0.005  # echoed back, never sent to COM
        profile.ChainLocate.assert_called_once_with(0.1, 0.2)


# ============================================================================
# CONVERT TO CURVE
# ============================================================================


class TestConvertToCurve:
    def test_success(self, sketch_mgr):
        sm, doc = sketch_mgr
        line, circle = MagicMock(), MagicMock()
        profile = _profile_with_geometry(lines=[line], circles=[circle])
        sm.active_profile = profile
        profile.ConvertToCurve.return_value = MagicMock()

        result = sm.convert_to_curve()
        assert result["status"] == "converted"
        assert result["type"] == "curve"
        assert result["element_count"] == 2

        # ConvertToCurve(NumberOfElements, ElementArray SAFEARRAY(DISPATCH))
        args, kwargs = profile.ConvertToCurve.call_args
        assert kwargs == {}
        assert args[0] == 2
        assert list(args[1]) == [line, circle]

    def test_no_active_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        sm.active_profile = None

        result = sm.convert_to_curve()
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_empty_sketch_is_rejected_before_com(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = _profile_with_geometry()
        sm.active_profile = profile

        result = sm.convert_to_curve()
        assert "no 2D geometry" in result["error"]
        profile.ConvertToCurve.assert_not_called()

    def test_com_error(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = _profile_with_geometry(lines=[MagicMock()])
        sm.active_profile = profile
        profile.ConvertToCurve.side_effect = Exception("COM error")

        result = sm.convert_to_curve()
        assert "error" in result
        assert "traceback" not in result  # tracebacks only under SOLIDEDGE_MCP_DEBUG


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
        assert "traceback" not in result  # tracebacks only under SOLIDEDGE_MCP_DEBUG


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

        profile.Lines2d = _collection([elem1])
        profile.Arcs2d = _collection([])
        profile.Circles2d = _collection([elem2])
        profile.Ellipses2d = _collection([])
        profile.EllipticalArcs2d = _collection([])
        profile.BSplineCurves2d = _collection([])
        profile.Conics2d = _collection([])
        profile.OrderedGeometry.return_value = (2, [elem1, elem2])

        result = sm.get_ordered_geometry()
        assert result["status"] == "ok"
        assert result["num_elements"] == 2

        # OrderedGeometry(NumElements [out], Elements SAFEARRAY [in,out]):
        # the count precedes the buffer, so the buffer goes in by keyword.
        args, kwargs = profile.OrderedGeometry.call_args
        assert args == ()
        assert set(kwargs) == {"Elements"}
        assert list(kwargs["Elements"]) == [elem1, elem2]
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


class TestCloseSketchIdempotence:
    """close_sketch() must not queue the same profile twice.

    Loft and sweep consume accumulated_profiles; a duplicated entry silently
    corrupts the next multi-profile feature.
    """

    def _sketch_manager(self):
        from unittest.mock import MagicMock

        from solidedge_mcp.backends.sketching import SketchManager

        doc_mgr = MagicMock()
        sm = SketchManager(doc_mgr)
        profile = MagicMock()
        profile.End.return_value = 0
        sm.active_profile = profile
        sm.active_sketch = MagicMock(Name="Sketch1")
        return sm, profile

    def test_second_close_does_not_duplicate_profile(self):
        sm, profile = self._sketch_manager()

        first = sm.close_sketch()
        assert first["status"] == "closed"
        assert first["accumulated_profiles"] == 1

        second = sm.close_sketch()
        assert second["status"] == "closed"
        assert second["accumulated_profiles"] == 1
        assert sm.get_accumulated_profiles() == [profile]

    def test_distinct_profiles_both_queue(self):
        from unittest.mock import MagicMock

        sm, first_profile = self._sketch_manager()
        sm.close_sketch()

        second_profile = MagicMock()
        second_profile.End.return_value = 0
        sm.active_profile = second_profile
        result = sm.close_sketch()

        assert result["accumulated_profiles"] == 2
        queued = sm.get_accumulated_profiles()
        assert queued[0] is first_profile
        assert queued[1] is second_profile


# ============================================================================
# ACTIVE PLANE INDEX
# ============================================================================


class TestGetActivePlaneIndex:
    def test_none_without_a_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        sm.active_profile = None
        assert sm.get_active_plane_index() is None

    def test_tracked_by_create_sketch(self, sketch_mgr):
        sm, doc = sketch_mgr
        doc.RefPlanes.Item.return_value = MagicMock()

        assert sm.create_sketch("Front")["status"] == "created"
        assert sm.get_active_plane_index() == 3  # Front/XZ

    def test_tracked_by_create_sketch_on_plane_index(self, sketch_mgr):
        sm, doc = sketch_mgr
        doc.RefPlanes.Count = 6

        assert sm.create_sketch_on_plane_index(5)["status"] == "created"
        assert sm.get_active_plane_index() == 5

    def test_cleared_with_the_sketch_state(self, sketch_mgr):
        sm, doc = sketch_mgr
        doc.RefPlanes.Item.return_value = MagicMock()
        sm.create_sketch("Top")

        sm.clear_state()
        assert sm.get_active_plane_index() is None

    def test_falls_back_to_matching_the_plane_name(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        profile.Plane.Name = "Right"
        sm.active_profile = profile
        sm.active_plane_index = None

        planes = [MagicMock(), MagicMock(), MagicMock()]
        planes[0].Name = "Top"
        planes[1].Name = "Right"
        planes[2].Name = "Front"
        doc.RefPlanes.Count = 3
        doc.RefPlanes.Item.side_effect = lambda i: planes[i - 1]

        assert sm.get_active_plane_index() == 2

    def test_none_when_the_plane_cannot_be_resolved(self, sketch_mgr):
        sm, doc = sketch_mgr
        profile = MagicMock()
        sm.active_profile = profile
        sm.active_plane_index = None
        type(profile).Plane = property(lambda self: (_ for _ in ()).throw(Exception("boom")))

        assert sm.get_active_plane_index() is None
