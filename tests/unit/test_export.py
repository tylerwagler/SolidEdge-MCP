"""
Unit tests for ExportManager backend methods.

Tests draft sheet management and assembly drawing views.
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
