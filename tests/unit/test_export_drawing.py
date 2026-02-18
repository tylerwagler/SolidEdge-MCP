"""
Unit tests for ExportManager drawing backend methods.

Tests draft sheet management, assembly drawing views, parts lists,
sheet collections, and drawing view counts/scales.
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
# DRAFT SHEET MANAGEMENT
# ============================================================================


class TestAddDraftSheet:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Name = "Sheet 2"

        sheets = MagicMock()
        sheets.Count = 2
        sheets.AddSheet.return_value = sheet
        doc.Sheets = sheets

        result = em.add_draft_sheet()
        assert result["status"] == "added"
        assert result["total_sheets"] == 2
        sheets.AddSheet.assert_called_once()
        sheet.Activate.assert_called_once()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_draft_sheet()
        assert "error" in result


# ============================================================================
# CREATE DRAFT DOCUMENT
# ============================================================================


@pytest.fixture
def doc_mgr():
    """Create DocumentManager with mocked connection."""
    from solidedge_mcp.backends.documents import DocumentManager

    conn = MagicMock()
    app = MagicMock()
    conn.get_application.return_value = app
    dm = DocumentManager(conn)
    return dm, app


class TestCreateDraftDocument:
    def test_success(self, doc_mgr):
        dm, app = doc_mgr
        draft_doc = MagicMock()
        draft_doc.Name = "Draft1.dft"
        app.Documents.Add.return_value = draft_doc

        result = dm.create_draft()
        assert result["status"] == "created"
        app.Documents.Add.assert_called_once()

    def test_with_template(self, doc_mgr):
        dm, app = doc_mgr
        draft_doc = MagicMock()
        draft_doc.Name = "Draft1.dft"
        app.Documents.Add.return_value = draft_doc

        result = dm.create_draft("C:/templates/a3.dft")
        assert result["status"] == "created"


# ============================================================================
# ASSEMBLY DRAWING VIEW
# ============================================================================


class TestAddAssemblyDrawingView:
    def test_success(self, export_mgr):
        em, doc = export_mgr

        model_link = MagicMock()
        model_links = MagicMock()
        model_links.Count = 1
        model_links.Item.return_value = model_link
        doc.ModelLinks = model_links

        sheet = MagicMock()
        dvs = MagicMock()
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_assembly_drawing_view(0.15, 0.15, "Front", 1.0)
        assert result["status"] == "added"
        assert result["orientation"] == "Front"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_assembly_drawing_view()
        assert "error" in result

    def test_no_model_link(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()

        model_links = MagicMock()
        model_links.Count = 0
        doc.ModelLinks = model_links

        result = em.add_assembly_drawing_view()
        assert "error" in result

    def test_invalid_orientation(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()

        model_links = MagicMock()
        model_links.Count = 1
        doc.ModelLinks = model_links

        result = em.add_assembly_drawing_view(orientation="InvalidView")
        assert "error" in result


# ============================================================================
# GET SHEET INFO
# ============================================================================


class TestGetSheetInfo:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Name = "Sheet 1"
        sheets = MagicMock()
        sheets.Count = 2
        doc.ActiveSheet = sheet
        doc.Sheets = sheets

        result = em.get_sheet_info()
        assert result["status"] == "ok"
        assert result["sheet_name"] == "Sheet 1"
        assert result["sheet_count"] == 2

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.get_sheet_info()
        assert "error" in result


# ============================================================================
# ACTIVATE SHEET
# ============================================================================


class TestActivateSheet:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Name = "Sheet 2"
        sheets = MagicMock()
        sheets.Count = 3
        sheets.Item.return_value = sheet
        doc.Sheets = sheets

        result = em.activate_sheet(1)
        assert result["status"] == "activated"
        assert result["sheet_name"] == "Sheet 2"
        sheet.Activate.assert_called_once()

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.activate_sheet(5)
        assert "error" in result

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.activate_sheet(0)
        assert "error" in result


# ============================================================================
# RENAME SHEET
# ============================================================================


class TestRenameSheet:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Name = "Sheet 1"
        sheets = MagicMock()
        sheets.Count = 2
        sheets.Item.return_value = sheet
        doc.Sheets = sheets

        result = em.rename_sheet(0, "Drawing A")
        assert result["status"] == "renamed"
        assert result["new_name"] == "Drawing A"
        assert sheet.Name == "Drawing A"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.rename_sheet(0, "New Name")
        assert "error" in result

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.rename_sheet(5, "New Name")
        assert "error" in result


# ============================================================================
# DELETE SHEET
# ============================================================================


class TestDeleteSheet:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Name = "Sheet 2"
        sheets = MagicMock()
        sheets.Count = 3
        sheets.Item.return_value = sheet
        doc.Sheets = sheets

        result = em.delete_sheet(1)
        assert result["status"] == "deleted"
        sheet.Delete.assert_called_once()

    def test_last_sheet(self, export_mgr):
        em, doc = export_mgr
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.delete_sheet(0)
        assert "error" in result

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.delete_sheet(0)
        assert "error" in result


# ============================================================================
# BATCH 10: DRAFT SHEET COLLECTIONS
# ============================================================================


class TestGetSheetDimensions:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dim1 = MagicMock()
        dim1.Type = 1
        dim1.Value = 0.05
        dim1.Name = "D1"
        dim2 = MagicMock()
        dim2.Type = 2
        dim2.Value = 0.1
        dim2.Name = "D2"

        dims = MagicMock()
        dims.Count = 2
        dims.Item.side_effect = lambda i: [None, dim1, dim2][i]
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet

        result = em.get_sheet_dimensions()
        assert result["count"] == 2
        assert result["dimensions"][0]["value"] == 0.05
        assert result["dimensions"][1]["name"] == "D2"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.get_sheet_dimensions()
        assert "error" in result

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        dims.Count = 0
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet

        result = em.get_sheet_dimensions()
        assert result["count"] == 0
        assert result["dimensions"] == []


class TestGetSheetBalloons:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        b1 = MagicMock()
        b1.BalloonText = "1"
        b1.x = 0.05
        b1.y = 0.1

        balloons = MagicMock()
        balloons.Count = 1
        balloons.Item.return_value = b1
        sheet.Balloons = balloons
        doc.ActiveSheet = sheet

        result = em.get_sheet_balloons()
        assert result["count"] == 1
        assert result["balloons"][0]["text"] == "1"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.get_sheet_balloons()
        assert "error" in result

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        balloons = MagicMock()
        balloons.Count = 0
        sheet.Balloons = balloons
        doc.ActiveSheet = sheet

        result = em.get_sheet_balloons()
        assert result["count"] == 0


class TestGetSheetTextBoxes:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        tb = MagicMock()
        tb.Text = "Hello"
        tb.x = 0.02
        tb.y = 0.03
        tb.Height = 0.005

        text_boxes = MagicMock()
        text_boxes.Count = 1
        text_boxes.Item.return_value = tb
        sheet.TextBoxes = text_boxes
        doc.ActiveSheet = sheet

        result = em.get_sheet_text_boxes()
        assert result["count"] == 1
        assert result["text_boxes"][0]["text"] == "Hello"
        assert result["text_boxes"][0]["height"] == 0.005

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.get_sheet_text_boxes()
        assert "error" in result

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        text_boxes = MagicMock()
        text_boxes.Count = 0
        sheet.TextBoxes = text_boxes
        doc.ActiveSheet = sheet

        result = em.get_sheet_text_boxes()
        assert result["count"] == 0


class TestGetSheetDrawingObjects:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        obj1 = MagicMock()
        obj1.Name = "Obj1"

        drawing_objects = MagicMock()
        drawing_objects.Count = 1
        drawing_objects.Item.return_value = obj1
        sheet.DrawingObjects = drawing_objects
        doc.ActiveSheet = sheet

        result = em.get_sheet_drawing_objects()
        assert result["count"] == 1
        assert result["drawing_objects"][0]["name"] == "Obj1"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.get_sheet_drawing_objects()
        assert "error" in result

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        drawing_objects = MagicMock()
        drawing_objects.Count = 0
        sheet.DrawingObjects = drawing_objects
        doc.ActiveSheet = sheet

        result = em.get_sheet_drawing_objects()
        assert result["count"] == 0


class TestGetSheetSections:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sec = MagicMock()
        sec.Label = "A-A"
        sec.Name = "Section1"
        sec.Type = 1

        sections = MagicMock()
        sections.Count = 1
        sections.Item.return_value = sec
        sheet.Sections = sections
        doc.ActiveSheet = sheet

        result = em.get_sheet_sections()
        assert result["count"] == 1
        assert result["sections"][0]["label"] == "A-A"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.get_sheet_sections()
        assert "error" in result

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sections = MagicMock()
        sections.Count = 0
        sheet.Sections = sections
        doc.ActiveSheet = sheet

        result = em.get_sheet_sections()
        assert result["count"] == 0


# ============================================================================
# TIER 3: CREATE PARTS LIST
# ============================================================================


class TestCreatePartsList:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dv = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        dvs.Item.return_value = dv
        sheet.DrawingViews = dvs

        parts_lists = MagicMock()
        parts_lists.Count = 1
        parts_lists.Add.return_value = MagicMock()
        sheet.PartsLists = parts_lists
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.create_parts_list(auto_balloon=True)
        assert result["status"] == "created"
        assert result["auto_balloon"] is True
        parts_lists.Add.assert_called_once_with(dv, "", 1, 1)

    def test_no_drawing_views(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 0
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.create_parts_list()
        assert "error" in result
        assert "No drawing views" in result["error"]

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.create_parts_list()
        assert "error" in result

    def test_no_auto_balloon(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dv = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        dvs.Item.return_value = dv
        sheet.DrawingViews = dvs

        parts_lists = MagicMock()
        parts_lists.Count = 1
        parts_lists.Add.return_value = MagicMock()
        sheet.PartsLists = parts_lists
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.create_parts_list(auto_balloon=False)
        assert result["status"] == "created"
        assert result["auto_balloon"] is False
        parts_lists.Add.assert_called_once_with(dv, "", 0, 1)


# ============================================================================
# DRAWING VIEW: GET COUNT
# ============================================================================


class TestGetDrawingViewCount:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 3
        # Make _oleobj_ raise so the late-binding branch is skipped
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.get_drawing_view_count()
        assert result["count"] == 3

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.get_drawing_view_count()
        assert "error" in result


# ============================================================================
# DRAWING VIEW: GET SCALE
# ============================================================================


class TestGetDrawingViewScale:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock()
        view.ScaleFactor = 2.0
        dvs = MagicMock()
        dvs.Count = 2
        dvs.Item.return_value = view
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.get_drawing_view_scale(0)
        assert result["view_index"] == 0
        assert result["scale"] == 2.0

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.get_drawing_view_scale(5)
        assert "error" in result

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.get_drawing_view_scale(0)
        assert "error" in result
