"""
Unit tests for AssemblyManager backend methods.

Tests BOM, interference check, occurrence bounding box.
Uses unittest.mock to simulate COM objects.
"""

from unittest.mock import MagicMock

import pytest


@pytest.fixture
def asm_mgr():
    """Create AssemblyManager with mocked dependencies."""
    from solidedge_mcp.backends.assembly import AssemblyManager
    dm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc
    return AssemblyManager(dm), doc


# ============================================================================
# OCCURRENCE BOUNDING BOX
# ============================================================================

class TestGetOccurrenceBoundingBox:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()

        def mock_get_range_box(min_pt, max_pt):
            min_pt[0], min_pt[1], min_pt[2] = 0.0, 0.0, 0.0
            max_pt[0], max_pt[1], max_pt[2] = 0.1, 0.2, 0.3

        occ.GetRangeBox = mock_get_range_box

        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.get_occurrence_bounding_box(0)
        assert result["component_index"] == 0
        assert result["min"] == [0.0, 0.0, 0.0]
        assert result["max"] == [0.1, 0.2, 0.3]
        assert result["size"] == pytest.approx([0.1, 0.2, 0.3])

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences  # Remove attribute

        result = am.get_occurrence_bounding_box(0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 2
        doc.Occurrences = occurrences

        result = am.get_occurrence_bounding_box(5)
        assert "error" in result


# ============================================================================
# BOM
# ============================================================================

class TestGetBom:
    def test_basic_bom(self, asm_mgr):
        am, doc = asm_mgr

        occ1 = MagicMock()
        occ1.OccurrenceFileName = "C:/parts/bolt.par"
        occ1.Name = "Bolt_1"
        occ1.IncludeInBom = True
        occ1.IsPatternItem = False

        occ2 = MagicMock()
        occ2.OccurrenceFileName = "C:/parts/bolt.par"
        occ2.Name = "Bolt_2"
        occ2.IncludeInBom = True
        occ2.IsPatternItem = False

        occ3 = MagicMock()
        occ3.OccurrenceFileName = "C:/parts/plate.par"
        occ3.Name = "Plate_1"
        occ3.IncludeInBom = True
        occ3.IsPatternItem = False

        occurrences = MagicMock()
        occurrences.Count = 3
        occurrences.Item.side_effect = lambda i: [None, occ1, occ2, occ3][i]
        doc.Occurrences = occurrences

        result = am.get_bom()
        assert result["total_occurrences"] == 3
        assert result["unique_parts"] == 2

        # Find bolt and plate in BOM
        bolt = next(b for b in result["bom"] if "bolt" in b["file_path"])
        plate = next(b for b in result["bom"] if "plate" in b["file_path"])
        assert bolt["quantity"] == 2
        assert plate["quantity"] == 1

    def test_excludes_non_bom_items(self, asm_mgr):
        am, doc = asm_mgr

        occ1 = MagicMock()
        occ1.OccurrenceFileName = "C:/parts/bolt.par"
        occ1.Name = "Bolt_1"
        occ1.IncludeInBom = True
        occ1.IsPatternItem = False

        occ2 = MagicMock()
        occ2.OccurrenceFileName = "C:/parts/ref.par"
        occ2.Name = "Reference_1"
        occ2.IncludeInBom = False
        occ2.IsPatternItem = False

        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.side_effect = lambda i: [None, occ1, occ2][i]
        doc.Occurrences = occurrences

        result = am.get_bom()
        assert result["unique_parts"] == 1
        assert result["bom"][0]["name"] == "Bolt_1"

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.get_bom()
        assert "error" in result

    def test_empty_assembly(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 0
        doc.Occurrences = occurrences

        result = am.get_bom()
        assert result["total_occurrences"] == 0
        assert result["unique_parts"] == 0
        assert result["bom"] == []


# ============================================================================
# INTERFERENCE CHECK
# ============================================================================

class TestCheckInterference:
    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.check_interference()
        assert "error" in result

    def test_too_few_components(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.check_interference()
        assert result["status"] == "no_interference"

    def test_invalid_component_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 3
        doc.Occurrences = occurrences

        result = am.check_interference(component_index=10)
        assert "error" in result


# ============================================================================
# EXISTING METHODS
# ============================================================================

class TestPlaceComponent:
    def test_at_origin(self, asm_mgr):
        am, doc = asm_mgr
        import os
        # Create a temp file path for the test
        occ = MagicMock()
        occ.Name = "Part_1"
        occ.GetTransform.return_value = [0.0, 0.0, 0.0, 0.0, 0.0, 0.0]

        occurrences = MagicMock()
        occurrences.AddByFilename.return_value = occ
        occurrences.Count = 1
        doc.Occurrences = occurrences

        # Mock os.path.exists
        import unittest.mock
        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.place_component("C:/test/part.par", 0, 0, 0)

        assert result["status"] == "added"
        occurrences.AddByFilename.assert_called_once()

    def test_file_not_found(self, asm_mgr):
        am, doc = asm_mgr
        import unittest.mock
        with unittest.mock.patch("os.path.exists", return_value=False):
            result = am.place_component("nonexistent.par")

        assert "error" in result


class TestSuppressComponent:
    def test_suppress(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.suppress_component(0, suppress=True)
        assert result["status"] == "updated"
        occ.Suppress.assert_called_once()

    def test_unsuppress(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.suppress_component(0, suppress=False)
        assert result["status"] == "updated"
        occ.Unsuppress.assert_called_once()
