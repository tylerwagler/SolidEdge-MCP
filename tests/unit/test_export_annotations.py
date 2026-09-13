"""
Unit tests for ExportManager annotation backend methods.

Tests dimensions, balloons, notes, text boxes, leaders, center marks,
centerlines, surface finish symbols, weld symbols, geometric tolerances,
and 2D geometry collection access.
Uses unittest.mock to simulate COM objects.
"""

from unittest.mock import MagicMock

import pytest

from solidedge_mcp.backends.constants import DocumentTypeConstants

IG_ASSEMBLY_DOCUMENT = DocumentTypeConstants.igAssemblyDocument
IG_DRAFT_DOCUMENT = DocumentTypeConstants.igDraftDocument
IG_PART_DOCUMENT = DocumentTypeConstants.igPartDocument


@pytest.fixture
def export_mgr():
    """Create ExportManager with mocked dependencies."""
    from solidedge_mcp.backends.export import ExportManager

    dm = MagicMock()
    doc = MagicMock()
    doc.Type = IG_DRAFT_DOCUMENT
    dm.get_active_document.return_value = doc
    return ExportManager(dm), doc


# ============================================================================
# A FAKE DRAFT SHEET
#
# The dimension methods resolve a coordinate to the element under it, which
# means walking real collections and calling real accessors. A MagicMock sheet
# hands back a MagicMock for every attribute, including Count, so it cannot
# stand in. These fakes model what Solid Edge actually exposes: GetStartPoint,
# GetEndPoint and GetCenterPoint are pure [out] and return the pair, and there
# are no StartX/CenterX properties.
# ============================================================================


class FakeCollection:
    """A 1-indexed COM collection."""

    def __init__(self, items):
        self._items = list(items)

    @property
    def Count(self):
        return len(self._items)

    def Item(self, index):
        return self._items[index - 1]


class FakeLine:
    def __init__(self, x1, y1, x2, y2):
        self._start = (x1, y1)
        self._end = (x2, y2)

    def GetStartPoint(self):
        return self._start

    def GetEndPoint(self):
        return self._end


class FakeCircle:
    def __init__(self, cx, cy, radius):
        self._center = (cx, cy)
        self.Radius = radius
        self.Diameter = radius * 2

    def GetCenterPoint(self):
        return self._center


class FakeArc(FakeCircle):
    def __init__(self, cx, cy, radius, start_angle=0.0, sweep=1.5707963267948966):
        super().__init__(cx, cy, radius)
        self.StartAngle = start_angle
        self.SweepAngle = sweep
        self._start = (cx + radius, cy)
        self._end = (cx, cy + radius)

    def GetStartPoint(self):
        return self._start

    def GetEndPoint(self):
        return self._end


class FakeDimensions:
    """Records what was called, and returns the Dimension's value like COM."""

    def __init__(self):
        self.calls = []

    def _record(self, name, args, value):
        self.calls.append((name, args))
        return value

    def AddDistanceBetweenObjects(self, *args):
        return self._record("AddDistanceBetweenObjects", args, 0.1)

    def AddAngleBetweenObjects(self, *args):
        return self._record("AddAngleBetweenObjects", args, 1.5707963267948966)

    def AddRadius(self, *args):
        return self._record("AddRadius", args, 0.02)

    def AddCircularDiameter(self, *args):
        return self._record("AddCircularDiameter", args, 0.04)

    def AddRadialDiameter(self, *args):
        return self._record("AddRadialDiameter", args, 0.06)

    def AddCoordinateOrigin(self, *args):
        return self._record("AddCoordinateOrigin", args, 0.0)

    def AddCoordinate(self, *args):
        return self._record("AddCoordinate", args, 0.1)

    def AddLength(self, *args):
        return self._record("AddLength", args, 0.1)

    def names(self):
        return [name for name, _args in self.calls]


class FakeSheet:
    """A draft sheet holding a 100x70mm rectangle, a circle and an arc."""

    def __init__(self, lines=None, circles=None, arcs=None):
        self.Lines2d = FakeCollection(
            lines
            if lines is not None
            else [
                FakeLine(0.05, 0.05, 0.15, 0.05),
                FakeLine(0.15, 0.05, 0.15, 0.12),
                FakeLine(0.15, 0.12, 0.05, 0.12),
                FakeLine(0.05, 0.12, 0.05, 0.05),
            ]
        )
        self.Circles2d = FakeCollection(
            circles if circles is not None else [FakeCircle(0.10, 0.22, 0.02)]
        )
        self.Arcs2d = FakeCollection(arcs if arcs is not None else [FakeArc(0.22, 0.22, 0.03)])
        self.DrawingViews = FakeCollection([])
        self.Dimensions = FakeDimensions()


@pytest.fixture
def draft(export_mgr):
    """An ExportManager whose active document is a draft with real geometry."""
    em, doc = export_mgr
    sheet = FakeSheet()
    doc.ActiveSheet = sheet
    return em, doc, sheet


# ============================================================================
# ADD DIMENSION
# ============================================================================


class TestAddDimension:
    """Dimensions.AddDistanceBetweenObjects needs the objects at each point.

    There is no AddDistanceBetweenPoints in any Solid Edge type library, so
    the two coordinates are resolved to the nearest element first.
    """

    def test_attaches_to_the_element_under_each_point(self, draft):
        em, _doc, sheet = draft

        result = em.add_dimension(0.05, 0.05, 0.15, 0.05)

        assert result["status"] == "created"
        assert result["value"] == 0.1
        assert sheet.Dimensions.names() == ["AddDistanceBetweenObjects"]
        # keyPoint is True at both ends so the dimension snaps to a vertex.
        args = sheet.Dimensions.calls[0][1]
        assert args[4] is True
        assert args[9] is True
        assert [hit["kind"] for hit in result["attached_to"]] == ["line", "line"]

    def test_empty_sheet_reports_what_is_missing(self, export_mgr):
        em, doc = export_mgr
        doc.ActiveSheet = FakeSheet(lines=[], circles=[], arcs=[])

        result = em.add_dimension(0.05, 0.05, 0.15, 0.05)

        assert "error" in result
        assert "dimension" in result["error"]
        assert result["point"] == [0.05, 0.05]

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.add_dimension(0, 0, 0.1, 0)


# ============================================================================
# ADD ANGULAR DIMENSION
# ============================================================================


class TestAddAngularDimension:
    """The angle comes from the two lines the outer points land on."""

    def test_uses_the_line_at_each_outer_point(self, draft):
        em, _doc, sheet = draft

        result = em.add_angular_dimension(0.10, 0.05, 0.15, 0.05, 0.15, 0.09)

        assert result["status"] == "created"
        assert result["value_radians"] == pytest.approx(1.5707963, rel=1e-5)
        assert sheet.Dimensions.names() == ["AddAngleBetweenObjects"]
        assert result["vertex"] == [0.15, 0.05]

    def test_refuses_when_both_points_land_on_one_line(self, draft):
        em, _doc, sheet = draft

        result = em.add_angular_dimension(0.07, 0.05, 0.10, 0.05, 0.12, 0.05)

        assert "error" in result
        assert "same line" in result["error"]
        assert sheet.Dimensions.calls == []

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.add_angular_dimension(0, 0, 0.05, 0.05, 0.1, 0)


# ============================================================================
# ADD RADIAL DIMENSION
# ============================================================================


class TestAddRadialDimension:
    """Dimensions.AddRadius takes the curve; AddRadial does not exist."""

    def test_finds_the_circle_under_the_point(self, draft):
        em, _doc, sheet = draft

        result = em.add_radial_dimension(0.10, 0.22, 0.12, 0.22)

        assert result["status"] == "created"
        assert result["value"] == 0.02
        assert sheet.Dimensions.names() == ["AddRadius"]
        assert result["attached_to"]["kind"] == "circle"

    def test_falls_back_to_the_centre(self, draft):
        """A caller may know only where the curve is centred."""
        em, _doc, sheet = draft

        result = em.add_radial_dimension(0.22, 0.22, 0.0, 0.0)

        assert result["status"] == "created"
        assert sheet.Dimensions.names() == ["AddRadius"]

    def test_no_curve_on_the_sheet(self, export_mgr):
        em, doc = export_mgr
        doc.ActiveSheet = FakeSheet(circles=[], arcs=[])

        result = em.add_radial_dimension(0.1, 0.1, 0.12, 0.1)

        assert "error" in result
        assert "circle or arc" in result["error"]

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.add_radial_dimension(0.1, 0.1, 0.12, 0.1)


# ============================================================================
# ADD DIAMETER DIMENSION
# ============================================================================


class TestAddDiameterDimension:
    """Circles take AddCircularDiameter; arcs take AddRadialDiameter."""

    def test_circle_uses_the_circular_form(self, draft):
        em, _doc, sheet = draft

        result = em.add_diameter_dimension(0.10, 0.22, 0.12, 0.22)

        assert result["status"] == "created"
        assert result["value"] == 0.04
        assert sheet.Dimensions.names() == ["AddCircularDiameter"]

    def test_arc_uses_the_radial_form(self, draft):
        em, _doc, sheet = draft

        result = em.add_diameter_dimension(0.22, 0.22, 0.25, 0.22)

        assert result["status"] == "created"
        assert sheet.Dimensions.names() == ["AddRadialDiameter"]
        assert result["attached_to"]["kind"] == "arc"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.add_diameter_dimension(0.1, 0.1, 0.12, 0.1)


# ============================================================================
# ADD ORDINATE DIMENSION
# ============================================================================


class TestAddOrdinateDimension:
    """A datum plus a measured point, both attached to real elements."""

    def test_sets_the_origin_then_measures(self, draft):
        em, _doc, sheet = draft

        result = em.add_ordinate_dimension(0.05, 0.05, 0.15, 0.05)

        assert result["status"] == "created"
        assert sheet.Dimensions.names() == ["AddCoordinateOrigin", "AddCoordinate"]
        assert len(result["attached_to"]) == 2

    def test_empty_sheet(self, export_mgr):
        em, doc = export_mgr
        doc.ActiveSheet = FakeSheet(lines=[], circles=[], arcs=[])

        result = em.add_ordinate_dimension(0.05, 0.05, 0.15, 0.05)

        assert "error" in result
        assert "datum" in result["error"]

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.add_ordinate_dimension(0, 0, 0.1, 0)


# ============================================================================
# ADD DISTANCE DIMENSION
# ============================================================================


class TestAddDistanceDimension:
    """The same operation as add_dimension, reached via add_2d_dimension."""

    def test_creates_a_dimension(self, draft):
        em, _doc, sheet = draft

        result = em.add_distance_dimension(0.05, 0.05, 0.15, 0.05)

        assert result["status"] == "created"
        assert sheet.Dimensions.names() == ["AddDistanceBetweenObjects"]
        assert result["point1"] == [0.05, 0.05]
        assert result["point2"] == [0.15, 0.05]

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.add_distance_dimension(0.0, 0.0, 0.1, 0.05)

    def test_com_error(self, draft):
        em, _doc, sheet = draft

        def boom(*args):
            raise Exception("COM error")

        sheet.Dimensions.AddDistanceBetweenObjects = boom

        assert "error" in em.add_distance_dimension(0.05, 0.05, 0.15, 0.05)


# ============================================================================
# ADD LENGTH DIMENSION
# ============================================================================


class TestAddLengthDimension:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        line = MagicMock()
        line.StartX = 0.0
        line.StartY = 0.0
        line.EndX = 0.1
        line.EndY = 0.0

        lines2d = MagicMock()
        lines2d.Count = 2
        lines2d.Item.return_value = line
        sheet.Lines2d = lines2d

        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet

        result = em.add_length_dimension(0)
        assert result["status"] == "added"
        assert result["dimension_type"] == "length"
        assert result["object_index"] == 0
        # AddLength(Object) dimensions the 2D element itself
        dims.AddLength.assert_called_once_with(line)

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        lines2d = MagicMock()
        lines2d.Count = 1
        sheet.Lines2d = lines2d
        doc.ActiveSheet = sheet

        result = em.add_length_dimension(5)
        assert "error" in result
        assert "Invalid line index" in result["error"]

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        result = em.add_length_dimension(0)
        assert "error" in result


# ============================================================================
# ADD RADIUS DIMENSION 2D
# ============================================================================


class TestAddRadiusDimension2d:
    def test_success_circle(self, draft):
        em, _doc, sheet = draft

        result = em.add_radius_dimension_2d(0, "circle")

        assert result["status"] == "added"
        assert result["dimension_type"] == "radius"
        assert result["object_type"] == "circle"
        assert result["value"] == 0.02
        # AddRadialDimension is in no Solid Edge type library; AddRadius is.
        assert sheet.Dimensions.names() == ["AddRadius"]

    def test_success_arc(self, draft):
        em, _doc, sheet = draft

        result = em.add_radius_dimension_2d(0, "arc")

        assert result["status"] == "added"
        assert result["object_type"] == "arc"
        assert sheet.Dimensions.names() == ["AddRadius"]

    def test_index_out_of_range(self, draft):
        em, _doc, _sheet = draft

        result = em.add_radius_dimension_2d(99, "circle")

        assert "error" in result
        assert "Invalid index" in result["error"]

    def test_invalid_object_type(self, draft):
        em, _doc, _sheet = draft

        result = em.add_radius_dimension_2d(0, "polygon")

        assert "error" in result
        assert "Invalid object_type" in result["error"]


# ============================================================================
# ADD ANGLE DIMENSION 2D
# ============================================================================


class TestAddAngleDimension2d:
    """The same operation as add_angular_dimension, via add_2d_dimension."""

    def test_creates_a_dimension(self, draft):
        em, _doc, sheet = draft

        result = em.add_angle_dimension_2d(0.10, 0.05, 0.15, 0.05, 0.15, 0.09)

        assert result["status"] == "created"
        assert sheet.Dimensions.names() == ["AddAngleBetweenObjects"]

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.add_angle_dimension_2d(0.0, 0.0, 0.05, 0.05, 0.1, 0.0)


# ============================================================================
# ADD CENTER MARK
# ============================================================================


class TestAddCenterMark:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        center_marks = MagicMock()
        sheet.CenterMarks = center_marks
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_center_mark(0.1, 0.1)
        assert result["status"] == "added"
        assert result["type"] == "center_mark"
        assert result["position"] == [0.1, 0.1]
        center_marks.Add.assert_called_once_with(0.1, 0.1, 0)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        result = em.add_center_mark(0.1, 0.1)
        assert "error" in result

    def test_exception(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.CenterMarks.Add.side_effect = Exception("COM error")
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_center_mark(0.1, 0.1)
        assert "error" in result


# ============================================================================
# ADD CENTERLINE
# ============================================================================


class TestAddCenterline:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        centerlines = MagicMock()
        sheet.CenterLines = centerlines
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_centerline(0.0, 0.05, 0.1, 0.05)
        assert result["status"] == "added"
        assert result["type"] == "centerline"
        assert result["start"] == [0.0, 0.05]
        assert result["end"] == [0.1, 0.05]
        centerlines.Add.assert_called_once_with(0.0, 0.05, 0, 0.1, 0.05, 0)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        result = em.add_centerline(0, 0, 0.1, 0)
        assert "error" in result

    def test_exception(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.CenterLines.Add.side_effect = Exception("COM error")
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_centerline(0, 0, 0.1, 0)
        assert "error" in result


# ============================================================================
# ADD SURFACE FINISH SYMBOL
# ============================================================================


class TestAddSurfaceFinishSymbol:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sfs = MagicMock()
        sheet.SurfaceFinishSymbols = sfs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_surface_finish_symbol(0.1, 0.1, "machined")
        assert result["status"] == "added"
        assert result["type"] == "surface_finish_symbol"
        assert result["symbol_type"] == "machined"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        result = em.add_surface_finish_symbol(0.1, 0.1)
        assert "error" in result

    def test_invalid_type(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()

        result = em.add_surface_finish_symbol(0.1, 0.1, "invalid_type")
        assert "error" in result


# ============================================================================
# ADD WELD SYMBOL
# ============================================================================


class TestAddWeldSymbol:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        ws = MagicMock()
        sheet.WeldSymbols = ws
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        symbol = MagicMock()
        ws.Add.return_value = symbol

        result = em.add_weld_symbol(0.1, 0.1, "fillet")
        assert result["status"] == "added"
        assert result["type"] == "weld_symbol"
        assert result["weld_type"] == "fillet"
        # WeldSymbols.Add(x1, y1, z1); the type is the TopType property
        ws.Add.assert_called_once_with(0.1, 0.1, 0)
        assert symbol.TopType == 1  # igDimWeldTopFillet

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        result = em.add_weld_symbol(0.1, 0.1)
        assert "error" in result

    def test_invalid_type(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()

        result = em.add_weld_symbol(0.1, 0.1, "invalid_weld")
        assert "error" in result


# ============================================================================
# ADD GEOMETRIC TOLERANCE
# ============================================================================


class TestAddGeometricTolerance:
    """The collection is FeatureControlFrames; sheet.FCFs does not exist."""

    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        frames = MagicMock()
        frame = MagicMock()
        frames.Add.return_value = frame
        frames.Count = 1
        sheet.FeatureControlFrames = frames
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_geometric_tolerance(0.1, 0.1, "0.05 A B")

        assert result["status"] == "added"
        assert result["type"] == "geometric_tolerance"
        assert result["text"] == "0.05 A B"
        frames.Add.assert_called_once_with(0.1, 0.1, 0)
        assert frame.Text == "0.05 A B"

    def test_no_longer_silently_writes_a_text_box(self, export_mgr):
        """The old fallback made every tolerance a plain text box."""
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.FeatureControlFrames.Add.side_effect = Exception("not available")
        text_boxes = MagicMock()
        sheet.TextBoxes = text_boxes
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_geometric_tolerance(0.1, 0.1, "0.05 A")

        assert "error" in result
        text_boxes.Add.assert_not_called()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.add_geometric_tolerance(0.1, 0.1)


# ============================================================================
# ADD TEXT BOX
# ============================================================================


class TestAddTextBox:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        text_boxes = MagicMock()
        text_box = MagicMock()
        text_boxes.Add.return_value = text_box
        sheet.TextBoxes = text_boxes
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_text_box(0.1, 0.1, "Hello")
        assert result["status"] == "added"
        assert result["text"] == "Hello"
        text_boxes.Add.assert_called_once_with(0.1, 0.1, 0)
        assert text_box.Text == "Hello"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        result = em.add_text_box(0.1, 0.1, "Test")
        assert "error" in result


# ============================================================================
# ADD LEADER
# ============================================================================


class TestAddLeader:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        leaders = MagicMock()
        leader = MagicMock()
        leaders.Add.return_value = leader
        sheet.Leaders = leaders
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_leader(0.05, 0.05, 0.15, 0.15, "Note")
        assert result["status"] == "added"
        assert result["text"] == "Note"
        leaders.Add.assert_called_once_with(0.05, 0.05, 0, 0.15, 0.15, 0)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        result = em.add_leader(0.05, 0.05, 0.15, 0.15)
        assert "error" in result


# ============================================================================
# ADD NOTE
# ============================================================================


class TestAddNote:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        text_boxes = MagicMock()
        text_box = MagicMock()
        text_boxes.Add.return_value = text_box
        sheet.TextBoxes = text_boxes
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_note(0.1, 0.1, "Test note")
        assert result["status"] == "added"
        assert result["type"] == "note"
        assert result["text"] == "Test note"
        text_boxes.Add.assert_called_once_with(0.1, 0.1, 0)
        assert text_box.Text == "Test note"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        result = em.add_note(0.1, 0.1, "Test")
        assert "error" in result


# ============================================================================
# ADD BALLOON
# ============================================================================


class TestAddBalloon:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        balloons = MagicMock()
        balloon = MagicMock()
        balloons.Add.return_value = balloon
        sheet.Balloons = balloons
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_balloon(0.1, 0.1, "1", 0.05, 0.05)
        assert result["status"] == "added"
        assert result["type"] == "balloon"
        # Balloons.Add(x1, y1, z1) places the balloon; the leader is a vertex.
        balloons.Add.assert_called_once_with(0.1, 0.1, 0)
        balloon.AddVertex.assert_called_once_with(0.05, 0.05, 0)
        assert balloon.BalloonText == "1"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        result = em.add_balloon(0.1, 0.1)
        assert "error" in result


# ============================================================================
# 2D GEOMETRY COLLECTION ACCESS
# ============================================================================


class TestGetLines2d:
    """Line2d has GetStartPoint/GetEndPoint, not StartX/StartY/EndX/EndY."""

    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Lines2d = FakeCollection(
            [FakeLine(0.0, 0.0, 0.1, 0.05), FakeLine(0.1, 0.05, 0.2, 0.0)]
        )
        doc.ActiveSheet = sheet

        result = em.get_lines2d()

        assert result["count"] == 2
        assert result["lines"][0]["start"] == [0.0, 0.0]
        assert result["lines"][0]["end"] == [0.1, 0.05]
        assert result["lines"][1]["index"] == 1

    def test_coordinates_are_not_silently_dropped(self, export_mgr):
        """Reading StartX returned an index and nothing else."""
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Lines2d = FakeCollection([FakeLine(0.0, 0.0, 0.1, 0.05)])
        doc.ActiveSheet = sheet

        assert set(em.get_lines2d()["lines"][0]) == {"index", "start", "end"}

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Lines2d = FakeCollection([])
        doc.ActiveSheet = sheet

        result = em.get_lines2d()
        assert result["count"] == 0
        assert result["lines"] == []

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.get_lines2d()


class TestGetCircles2d:
    """Circle2d has GetCenterPoint, not CenterX/CenterY."""

    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Circles2d = FakeCollection([FakeCircle(0.05, 0.05, 0.02)])
        doc.ActiveSheet = sheet

        result = em.get_circles2d()

        assert result["count"] == 1
        assert result["circles"][0]["center"] == [0.05, 0.05]
        assert result["circles"][0]["radius"] == 0.02
        assert result["circles"][0]["diameter"] == 0.04

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Circles2d = FakeCollection([])
        doc.ActiveSheet = sheet

        result = em.get_circles2d()
        assert result["count"] == 0
        assert result["circles"] == []

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.get_circles2d()


class TestGetArcs2d:
    """Arc2d reports StartAngle and SweepAngle; there is no EndAngle."""

    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Arcs2d = FakeCollection([FakeArc(0.1, 0.1, 0.03, 0.0, 3.14159)])
        doc.ActiveSheet = sheet

        result = em.get_arcs2d()

        assert result["count"] == 1
        assert result["arcs"][0]["center"] == [0.1, 0.1]
        assert result["arcs"][0]["radius"] == 0.03
        assert result["arcs"][0]["start_angle"] == 0.0
        assert result["arcs"][0]["sweep_angle"] == 3.14159

    def test_end_angle_is_derived_from_the_sweep(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Arcs2d = FakeCollection([FakeArc(0.1, 0.1, 0.03, 0.5, 1.0)])
        doc.ActiveSheet = sheet

        assert em.get_arcs2d()["arcs"][0]["end_angle"] == pytest.approx(1.5)

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Arcs2d = FakeCollection([])
        doc.ActiveSheet = sheet

        result = em.get_arcs2d()
        assert result["count"] == 0
        assert result["arcs"] == []

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.get_arcs2d()


# ============================================================================
# DOCUMENT-TYPE GUARD (replaces the old hasattr probe)
# ============================================================================


class TestDraftDocumentGuard:
    """The guard reads Document.Type instead of probing for Sheets."""

    def test_part_document_is_rejected(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        result = em.add_balloon(0.1, 0.1)
        assert result["error"] == "Active document is not a draft document"

    def test_raising_type_getter_is_rejected(self, export_mgr):
        em, doc = export_mgr
        type(doc).Type = property(lambda self: (_ for _ in ()).throw(Exception("gone")))
        try:
            result = em.add_balloon(0.1, 0.1)
            assert result["error"] == "Active document is not a draft document"
        finally:
            del type(doc).Type

    def test_raising_sheet_getter_surfaces_the_real_error(self, export_mgr):
        em, doc = export_mgr
        type(doc).ActiveSheet = property(
            lambda self: (_ for _ in ()).throw(Exception("ActiveSheet exploded"))
        )
        try:
            result = em.add_balloon(0.1, 0.1)
            # The old hasattr probe swallowed this as "not a draft document".
            assert "error" in result
            assert result["error"] != "Active document is not a draft document"
        finally:
            del type(doc).ActiveSheet


class TestWeldTypeConstants:
    """The weld type map now uses DimWeldTypeConstants values, not 0..4."""

    @pytest.mark.parametrize(
        ("weld_type", "expected"),
        [
            ("fillet", 1),  # igDimWeldTopFillet
            ("spot", 2),  # igDimWeldTopSpot
            ("seam", 3),  # igDimWeldTopSeam
            ("groove", 5),  # igDimWeldTopVGroove
            ("plug", 6),  # igDimWeldTopSlot
        ],
    )
    def test_top_type_value(self, export_mgr, weld_type, expected):
        em, doc = export_mgr
        sheet = MagicMock()
        ws = MagicMock()
        symbol = MagicMock()
        ws.Add.return_value = symbol
        sheet.WeldSymbols = ws
        doc.ActiveSheet = sheet

        result = em.add_weld_symbol(0.1, 0.1, weld_type)

        assert result["status"] == "added"
        ws.Add.assert_called_once_with(0.1, 0.1, 0)
        assert symbol.TopType == expected
        # 0 is igDimWeldTypeNone - never a valid symbol
        assert symbol.TopType != 0


# ============================================================================
# DRAW SHEET GEOMETRY
# ============================================================================


class TestDrawSheetGeometry:
    """The server could read and dimension sheet geometry but not create any."""

    def test_line(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Lines2d.Count = 1
        doc.ActiveSheet = sheet

        result = em.draw_sheet_geometry("line", x1=0.0, y1=0.0, x2=0.1, y2=0.05)

        assert result["status"] == "created"
        assert result["elements_created"] == 1
        sheet.Lines2d.AddBy2Points.assert_called_once_with(0.0, 0.0, 0.1, 0.05)

    def test_rectangle_closes_with_four_lines(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Lines2d.Count = 4
        doc.ActiveSheet = sheet

        result = em.draw_sheet_geometry("rectangle", x1=0.0, y1=0.0, x2=0.1, y2=0.05)

        assert result["elements_created"] == 4
        corners = [call.args for call in sheet.Lines2d.AddBy2Points.call_args_list]
        assert corners == [
            (0.0, 0.0, 0.1, 0.0),
            (0.1, 0.0, 0.1, 0.05),
            (0.1, 0.05, 0.0, 0.05),
            (0.0, 0.05, 0.0, 0.0),
        ]

    def test_circle(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Circles2d.Count = 1
        doc.ActiveSheet = sheet

        result = em.draw_sheet_geometry("circle", center_x=0.1, center_y=0.2, radius=0.03)

        assert result["status"] == "created"
        sheet.Circles2d.AddByCenterRadius.assert_called_once_with(0.1, 0.2, 0.03)

    def test_circle_rejects_a_non_positive_radius(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        doc.ActiveSheet = sheet

        result = em.draw_sheet_geometry("circle", center_x=0.1, center_y=0.2, radius=0.0)

        assert "error" in result
        assert "meters" in result["error"]
        sheet.Circles2d.AddByCenterRadius.assert_not_called()

    def test_circle_3point(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Circles2d.Count = 1
        doc.ActiveSheet = sheet

        em.draw_sheet_geometry("circle_3point", x1=0.0, y1=0.0, x2=0.1, y2=0.1, x3=0.2, y3=0.0)

        sheet.Circles2d.AddBy3Points.assert_called_once_with(0.0, 0.0, 0.1, 0.1, 0.2, 0.0)

    def test_arc(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Arcs2d.Count = 1
        doc.ActiveSheet = sheet

        em.draw_sheet_geometry("arc", center_x=0.1, center_y=0.1, x1=0.13, y1=0.1, x2=0.1, y2=0.13)

        sheet.Arcs2d.AddByCenterStartEnd.assert_called_once_with(0.1, 0.1, 0.13, 0.1, 0.1, 0.13)

    def test_unknown_shape(self, export_mgr):
        em, doc = export_mgr
        doc.ActiveSheet = MagicMock()

        result = em.draw_sheet_geometry("spiral")

        assert "error" in result
        assert "Unknown shape" in result["error"]

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.draw_sheet_geometry("line")
