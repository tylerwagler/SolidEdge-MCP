"""
Unit tests for ExportManager annotation backend methods.

Tests dimensions, balloons, notes, text boxes, leaders, center marks,
centerlines, surface finish symbols, weld symbols, geometric tolerances,
and 2D geometry collection access.
Uses unittest.mock to simulate COM objects.
"""

from unittest.mock import MagicMock

import pytest


@pytest.fixture
def export_mgr():
    """Create ExportManager with mocked dependencies."""
    from solidedge_mcp.backends.export import ExportManager

    dm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc
    return ExportManager(dm), doc


# ============================================================================
# ADD DIMENSION
# ============================================================================


class TestAddDimension:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_dimension(0.0, 0.0, 0.1, 0.0)
        assert result["status"] == "added"
        assert result["type"] == "dimension"
        dims.AddLength.assert_called_once()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_dimension(0, 0, 0.1, 0)
        assert "error" in result


# ============================================================================
# ADD ANGULAR DIMENSION
# ============================================================================


class TestAddAngularDimension:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_angular_dimension(0.0, 0.0, 0.05, 0.05, 0.1, 0.0)
        assert result["status"] == "created"
        assert result["type"] == "angular_dimension"
        assert result["vertex"] == [0.05, 0.05]
        dims.AddAngular.assert_called_once()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_angular_dimension(0, 0, 0.05, 0.05, 0.1, 0)
        assert "error" in result

    def test_exception(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Dimensions.AddAngular.side_effect = Exception("COM error")
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_angular_dimension(0, 0, 0, 0, 0, 0)
        assert "error" in result


# ============================================================================
# ADD RADIAL DIMENSION
# ============================================================================


class TestAddRadialDimension:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_radial_dimension(0.05, 0.05, 0.1, 0.05)
        assert result["status"] == "created"
        assert result["type"] == "radial_dimension"
        assert result["center"] == [0.05, 0.05]
        dims.AddRadial.assert_called_once()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_radial_dimension(0.05, 0.05, 0.1, 0.05)
        assert "error" in result

    def test_custom_text_position(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_radial_dimension(0.05, 0.05, 0.1, 0.05, dim_x=0.2, dim_y=0.2)
        assert result["status"] == "created"
        assert result["text_position"] == [0.2, 0.2]


# ============================================================================
# ADD DIAMETER DIMENSION
# ============================================================================


class TestAddDiameterDimension:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_diameter_dimension(0.05, 0.05, 0.1, 0.05)
        assert result["status"] == "created"
        assert result["type"] == "diameter_dimension"
        assert result["center"] == [0.05, 0.05]
        dims.AddDiameter.assert_called_once()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_diameter_dimension(0.05, 0.05, 0.1, 0.05)
        assert "error" in result

    def test_exception(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Dimensions.AddDiameter.side_effect = Exception("COM error")
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_diameter_dimension(0, 0, 0.1, 0)
        assert "error" in result


# ============================================================================
# ADD ORDINATE DIMENSION
# ============================================================================


class TestAddOrdinateDimension:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_ordinate_dimension(0.0, 0.0, 0.1, 0.0)
        assert result["status"] == "created"
        assert result["type"] == "ordinate_dimension"
        assert result["origin"] == [0.0, 0.0]
        assert result["point"] == [0.1, 0.0]
        dims.AddOrdinate.assert_called_once()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_ordinate_dimension(0, 0, 0.1, 0)
        assert "error" in result

    def test_custom_text_position(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_ordinate_dimension(0.0, 0.0, 0.1, 0.0, dim_x=0.15, dim_y=0.05)
        assert result["status"] == "created"
        assert result["text_position"] == [0.15, 0.05]


# ============================================================================
# ADD DISTANCE DIMENSION
# ============================================================================


class TestAddDistanceDimension:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet

        result = em.add_distance_dimension(0.0, 0.0, 0.1, 0.05)
        assert result["status"] == "added"
        assert result["dimension_type"] == "distance"
        assert result["point1"] == [0.0, 0.0]
        assert result["point2"] == [0.1, 0.05]
        dims.AddDistanceBetweenPoints.assert_called_once()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.add_distance_dimension(0.0, 0.0, 0.1, 0.05)
        assert "error" in result

    def test_com_error(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        dims.AddDistanceBetweenPoints.side_effect = Exception("COM error")
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet

        result = em.add_distance_dimension(0.0, 0.0, 0.1, 0.05)
        assert "error" in result


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
        dims.AddLength.assert_called_once()

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
        del doc.ActiveSheet

        result = em.add_length_dimension(0)
        assert "error" in result


# ============================================================================
# ADD RADIUS DIMENSION 2D
# ============================================================================


class TestAddRadiusDimension2d:
    def test_success_circle(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        circle = MagicMock()
        circle.CenterX = 0.05
        circle.CenterY = 0.05
        circle.Radius = 0.02

        circles2d = MagicMock()
        circles2d.Count = 1
        circles2d.Item.return_value = circle
        sheet.Circles2d = circles2d

        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet

        result = em.add_radius_dimension_2d(0, "circle")
        assert result["status"] == "added"
        assert result["dimension_type"] == "radius"
        assert result["object_type"] == "circle"
        dims.AddRadialDimension.assert_called_once()

    def test_success_arc(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        arc = MagicMock()
        arc.CenterX = 0.1
        arc.CenterY = 0.1
        arc.Radius = 0.03

        arcs2d = MagicMock()
        arcs2d.Count = 1
        arcs2d.Item.return_value = arc
        sheet.Arcs2d = arcs2d

        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet

        result = em.add_radius_dimension_2d(0, "arc")
        assert result["status"] == "added"
        assert result["dimension_type"] == "radius"
        assert result["object_type"] == "arc"

    def test_invalid_object_type(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        doc.ActiveSheet = sheet

        result = em.add_radius_dimension_2d(0, "polygon")
        assert "error" in result
        assert "Invalid object_type" in result["error"]


# ============================================================================
# ADD ANGLE DIMENSION 2D
# ============================================================================


class TestAddAngleDimension2d:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet

        result = em.add_angle_dimension_2d(0.0, 0.0, 0.05, 0.05, 0.1, 0.0)
        assert result["status"] == "added"
        assert result["dimension_type"] == "angle"
        assert result["vertex"] == [0.05, 0.05]
        dims.AddAngle.assert_called_once()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.add_angle_dimension_2d(0.0, 0.0, 0.05, 0.05, 0.1, 0.0)
        assert "error" in result

    def test_com_error(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        dims.AddAngle.side_effect = Exception("COM error")
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet

        result = em.add_angle_dimension_2d(0.0, 0.0, 0.05, 0.05, 0.1, 0.0)
        assert "error" in result


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
        del doc.Sheets

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
        sheet.Centerlines = centerlines
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
        del doc.Sheets

        result = em.add_centerline(0, 0, 0.1, 0)
        assert "error" in result

    def test_exception(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Centerlines.Add.side_effect = Exception("COM error")
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
        del doc.Sheets

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

        result = em.add_weld_symbol(0.1, 0.1, "fillet")
        assert result["status"] == "added"
        assert result["type"] == "weld_symbol"
        assert result["weld_type"] == "fillet"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

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
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        fcfs = MagicMock()
        fcf = MagicMock()
        fcfs.Add.return_value = fcf
        sheet.FCFs = fcfs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_geometric_tolerance(0.1, 0.1, "0.05 A B")
        assert result["status"] == "added"
        assert result["type"] == "geometric_tolerance"
        assert result["text"] == "0.05 A B"
        fcfs.Add.assert_called_once_with(0.1, 0.1, 0)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_geometric_tolerance(0.1, 0.1)
        assert "error" in result

    def test_fallback_to_textbox(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.FCFs.Add.side_effect = Exception("FCFs not available")
        text_boxes = MagicMock()
        text_box = MagicMock()
        text_boxes.Add.return_value = text_box
        sheet.TextBoxes = text_boxes
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_geometric_tolerance(0.1, 0.1, "0.05 A")
        assert result["status"] == "added"
        text_boxes.Add.assert_called_once_with(0.1, 0.1, 0)


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
        del doc.Sheets

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
        del doc.Sheets

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
        del doc.Sheets

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
        balloons.Add.assert_called_once_with(0.05, 0.05, 0, 0.1, 0.1, 0)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_balloon(0.1, 0.1)
        assert "error" in result


# ============================================================================
# 2D GEOMETRY COLLECTION ACCESS
# ============================================================================


class TestGetLines2d:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        line1 = MagicMock()
        line1.StartX = 0.0
        line1.StartY = 0.0
        line1.EndX = 0.1
        line1.EndY = 0.05

        line2 = MagicMock()
        line2.StartX = 0.1
        line2.StartY = 0.05
        line2.EndX = 0.2
        line2.EndY = 0.0

        lines2d = MagicMock()
        lines2d.Count = 2
        lines2d.Item.side_effect = lambda i: {1: line1, 2: line2}[i]
        sheet.Lines2d = lines2d
        doc.ActiveSheet = sheet

        result = em.get_lines2d()
        assert result["count"] == 2
        assert result["lines"][0]["start"] == [0.0, 0.0]
        assert result["lines"][0]["end"] == [0.1, 0.05]
        assert result["lines"][1]["index"] == 1

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        lines2d = MagicMock()
        lines2d.Count = 0
        sheet.Lines2d = lines2d
        doc.ActiveSheet = sheet

        result = em.get_lines2d()
        assert result["count"] == 0
        assert result["lines"] == []

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.get_lines2d()
        assert "error" in result


class TestGetCircles2d:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        circle1 = MagicMock()
        circle1.CenterX = 0.05
        circle1.CenterY = 0.05
        circle1.Radius = 0.02

        circles2d = MagicMock()
        circles2d.Count = 1
        circles2d.Item.return_value = circle1
        sheet.Circles2d = circles2d
        doc.ActiveSheet = sheet

        result = em.get_circles2d()
        assert result["count"] == 1
        assert result["circles"][0]["center"] == [0.05, 0.05]
        assert result["circles"][0]["radius"] == 0.02

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        circles2d = MagicMock()
        circles2d.Count = 0
        sheet.Circles2d = circles2d
        doc.ActiveSheet = sheet

        result = em.get_circles2d()
        assert result["count"] == 0
        assert result["circles"] == []

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.get_circles2d()
        assert "error" in result


class TestGetArcs2d:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        arc1 = MagicMock()
        arc1.CenterX = 0.1
        arc1.CenterY = 0.1
        arc1.Radius = 0.03
        arc1.StartAngle = 0.0
        arc1.EndAngle = 3.14159

        arcs2d = MagicMock()
        arcs2d.Count = 1
        arcs2d.Item.return_value = arc1
        sheet.Arcs2d = arcs2d
        doc.ActiveSheet = sheet

        result = em.get_arcs2d()
        assert result["count"] == 1
        assert result["arcs"][0]["center"] == [0.1, 0.1]
        assert result["arcs"][0]["radius"] == 0.03
        assert result["arcs"][0]["start_angle"] == 0.0
        assert result["arcs"][0]["end_angle"] == 3.14159

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        arcs2d = MagicMock()
        arcs2d.Count = 0
        sheet.Arcs2d = arcs2d
        doc.ActiveSheet = sheet

        result = em.get_arcs2d()
        assert result["count"] == 0
        assert result["arcs"] == []

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.get_arcs2d()
        assert "error" in result
