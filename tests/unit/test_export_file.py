"""
Unit tests for ExportManager file export backend methods.

Tests flat DXF export, PRC export, PLMXML export.
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
# FLAT DXF EXPORT
# ============================================================================


class TestExportFlatDxf:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        flat_models = MagicMock()
        doc.FlatPatternModels = flat_models

        result = em.export_flat_dxf("C:/output/flat.dxf")
        assert result["status"] == "exported"
        assert result["format"] == "Flat DXF"
        flat_models.SaveAsFlatDXFEx.assert_called_once()

    def test_not_sheet_metal(self, export_mgr):
        em, doc = export_mgr
        del doc.FlatPatternModels

        result = em.export_flat_dxf("C:/output/flat.dxf")
        assert "error" in result

    def test_adds_extension(self, export_mgr):
        em, doc = export_mgr
        flat_models = MagicMock()
        doc.FlatPatternModels = flat_models

        result = em.export_flat_dxf("C:/output/flat")
        assert result["path"] == "C:/output/flat.dxf"


class TestExportToPrc:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        doc.Name = "part1.par"

        result = em.export_to_prc("C:/temp/output.prc")
        assert result["status"] == "exported"
        assert result["format"] == "PRC"
        assert result["path"] == "C:/temp/output.prc"
        doc.SaveAsPRC.assert_called_once_with("C:/temp/output.prc")

    def test_auto_extension(self, export_mgr):
        em, doc = export_mgr

        result = em.export_to_prc("C:/temp/output")
        assert result["path"] == "C:/temp/output.prc"
        doc.SaveAsPRC.assert_called_once_with("C:/temp/output.prc")

    def test_com_error(self, export_mgr):
        em, doc = export_mgr
        doc.SaveAsPRC.side_effect = Exception("PRC not supported")

        result = em.export_to_prc("C:/temp/output.prc")
        assert "error" in result


class TestExportToPlmxml:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        doc.Name = "part1.par"

        result = em.export_to_plmxml("C:/temp/output.xml", "C:/config/plmxml.ini")
        assert result["status"] == "exported"
        assert result["format"] == "PLMXML"
        assert result["path"] == "C:/temp/output.xml"
        assert result["ini_file"] == "C:/config/plmxml.ini"
        doc.SaveAsPLMXML.assert_called_once_with("C:/temp/output.xml", "C:/config/plmxml.ini")

    def test_com_error(self, export_mgr):
        em, doc = export_mgr
        doc.SaveAsPLMXML.side_effect = Exception("PLMXML export failed")

        result = em.export_to_plmxml("C:/temp/output.xml", "C:/config/plmxml.ini")
        assert "error" in result
