"""
Unit tests for AssemblyManager query backend methods.

Tests query mixin: GetOccurrenceBoundingBox, GetBom, GetDocumentTree,
GetComponentDisplayName, GetOccurrenceDocument, GetSubOccurrences,
CheckInterference, IsSubassembly, GetOccurrenceBodies,
GetOccurrenceStyle, GetFaceStyle, GetOccurrence.
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


@pytest.fixture
def asm_mgr_with_sketch():
    """Create AssemblyManager with mocked doc and sketch manager."""
    from solidedge_mcp.backends.assembly import AssemblyManager

    dm = MagicMock()
    sm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc
    return AssemblyManager(dm, sm), doc, sm


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
# DOCUMENT TREE
# ============================================================================


class TestGetDocumentTree:
    def test_flat_assembly(self, asm_mgr):
        am, doc = asm_mgr

        occ1 = MagicMock()
        occ1.Name = "Part_1"
        occ1.OccurrenceFileName = "C:/parts/part1.par"
        occ1.Visible = True
        occ1.SubOccurrences.Count = 0

        occ2 = MagicMock()
        occ2.Name = "Part_2"
        occ2.OccurrenceFileName = "C:/parts/part2.par"
        occ2.Visible = True
        occ2.SubOccurrences.Count = 0

        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.side_effect = lambda i: [None, occ1, occ2][i]
        doc.Occurrences = occurrences
        doc.Name = "Assembly.asm"

        result = am.get_document_tree()
        assert result["top_level_count"] == 2
        assert result["tree"][0]["name"] == "Part_1"
        assert result["tree"][1]["file"] == "C:/parts/part2.par"

    def test_nested_assembly(self, asm_mgr):
        am, doc = asm_mgr

        # Sub-occurrence
        sub_occ = MagicMock()
        sub_occ.Name = "SubPart_1"
        sub_occ.OccurrenceFileName = "C:/parts/sub.par"
        sub_occ.Visible = True
        sub_occs = MagicMock()
        sub_occs.Count = 0
        sub_occ.SubOccurrences = sub_occs

        # Top-level occurrence with children
        occ1 = MagicMock()
        occ1.Name = "SubAsm_1"
        occ1.OccurrenceFileName = "C:/asm/sub.asm"
        occ1.Visible = True
        sub_occs_parent = MagicMock()
        sub_occs_parent.Count = 1
        sub_occs_parent.Item.return_value = sub_occ
        occ1.SubOccurrences = sub_occs_parent

        occurrences = MagicMock()
        occurrences.Count = 1
        occurrences.Item.return_value = occ1
        doc.Occurrences = occurrences
        doc.Name = "TopAsm.asm"

        result = am.get_document_tree()
        assert result["top_level_count"] == 1
        assert result["tree"][0]["name"] == "SubAsm_1"
        assert len(result["tree"][0]["children"]) == 1
        assert result["tree"][0]["children"][0]["name"] == "SubPart_1"

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.get_document_tree()
        assert "error" in result


# ============================================================================
# ASSEMBLY: GET COMPONENT DISPLAY NAME
# ============================================================================


class TestGetComponentDisplayName:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occ.DisplayName = "Bolt M10x30"
        occ.Name = "Bolt_1"
        occ.OccurrenceFileName = "C:/parts/bolt.par"
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.get_component_display_name(0)
        assert result["display_name"] == "Bolt M10x30"
        assert result["name"] == "Bolt_1"
        assert result["file_name"] == "C:/parts/bolt.par"

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.get_component_display_name(5)
        assert "error" in result

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.get_component_display_name(0)
        assert "error" in result


# ============================================================================
# ASSEMBLY: GET OCCURRENCE DOCUMENT
# ============================================================================


class TestGetOccurrenceDocument:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ_doc = MagicMock()
        occ_doc.Name = "bolt.par"
        occ_doc.FullName = "C:/parts/bolt.par"
        occ_doc.Type = 1
        occ_doc.ReadOnly = False

        occ = MagicMock()
        occ.OccurrenceDocument = occ_doc
        occ.OccurrenceFileName = "C:/parts/bolt.par"
        occurrences = MagicMock()
        occurrences.Count = 1
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.get_occurrence_document(0)
        assert result["document_name"] == "bolt.par"
        assert result["full_name"] == "C:/parts/bolt.par"
        assert result["type"] == 1
        assert result["read_only"] is False

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.get_occurrence_document(5)
        assert "error" in result

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.get_occurrence_document(0)
        assert "error" in result


# ============================================================================
# ASSEMBLY: GET SUB-OCCURRENCES
# ============================================================================


class TestGetSubOccurrences:
    def test_with_children(self, asm_mgr):
        am, doc = asm_mgr
        child1 = MagicMock()
        child1.Name = "SubPart_1"
        child1.OccurrenceFileName = "C:/parts/sub1.par"

        child2 = MagicMock()
        child2.Name = "SubPart_2"
        child2.OccurrenceFileName = "C:/parts/sub2.par"

        sub_occs = MagicMock()
        sub_occs.Count = 2
        sub_occs.Item.side_effect = lambda i: {1: child1, 2: child2}[i]

        occ = MagicMock()
        occ.SubOccurrences = sub_occs
        occurrences = MagicMock()
        occurrences.Count = 1
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.get_sub_occurrences(0)
        assert result["count"] == 2
        assert result["sub_occurrences"][0]["name"] == "SubPart_1"
        assert result["sub_occurrences"][1]["file"] == "C:/parts/sub2.par"

    def test_no_children(self, asm_mgr):
        am, doc = asm_mgr
        sub_occs = MagicMock()
        sub_occs.Count = 0

        occ = MagicMock()
        occ.SubOccurrences = sub_occs
        occurrences = MagicMock()
        occurrences.Count = 1
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.get_sub_occurrences(0)
        assert result["count"] == 0
        assert result["sub_occurrences"] == []

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.get_sub_occurrences(5)
        assert "error" in result

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.get_sub_occurrences(0)
        assert "error" in result


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
# ASSEMBLY: IS SUBASSEMBLY
# ============================================================================


class TestIsSubassembly:
    def test_true(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occ.Subassembly = True
        occ.Name = "SubAsm_1"
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.is_subassembly(0)
        assert result["is_subassembly"] is True
        assert result["name"] == "SubAsm_1"

    def test_false(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occ.Subassembly = False
        occ.Name = "Part_1"
        occurrences = MagicMock()
        occurrences.Count = 1
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.is_subassembly(0)
        assert result["is_subassembly"] is False

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.is_subassembly(5)
        assert "error" in result

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.is_subassembly(0)
        assert "error" in result


# ============================================================================
# BATCH 8: GET OCCURRENCE BODIES
# ============================================================================


class TestGetOccurrenceBodies:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        body1 = MagicMock()
        body1.Name = "Body_1"
        body1.Volume = 0.001
        body2 = MagicMock()
        body2.Name = "Body_2"
        body2.Volume = 0.002

        bodies = MagicMock()
        bodies.Count = 2
        bodies.Item.side_effect = lambda i: {1: body1, 2: body2}[i]

        occ = MagicMock()
        occ.Bodies = bodies
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.get_occurrence_bodies(0)
        assert result["body_count"] == 2
        assert result["bodies"][0]["name"] == "Body_1"
        assert result["bodies"][0]["volume"] == 0.001
        assert result["bodies"][1]["name"] == "Body_2"

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.get_occurrence_bodies(0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.get_occurrence_bodies(5)
        assert "error" in result


# ============================================================================
# BATCH 8: GET OCCURRENCE STYLE
# ============================================================================


class TestGetOccurrenceStyle:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occ.Style = "Steel"
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.get_occurrence_style(0)
        assert result["component_index"] == 0
        assert result["style"] == "Steel"

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.get_occurrence_style(0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.get_occurrence_style(5)
        assert "error" in result


# ============================================================================
# BATCH 8: GET FACE STYLE
# ============================================================================


class TestGetFaceStyle:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occ.GetFaceStyle2.return_value = "Aluminum"
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.get_face_style(0)
        assert result["component_index"] == 0
        assert result["face_style"] == "Aluminum"
        occ.GetFaceStyle2.assert_called_once()

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.get_face_style(0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.get_face_style(5)
        assert "error" in result


# ============================================================================
# GET OCCURRENCE BY ID
# ============================================================================


class TestGetOccurrence:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occ.Name = "Part_1"
        occ.OccurrenceFileName = "C:/parts/part1.par"
        occ.GetTransform.return_value = [0.1, 0.2, 0.3, 0.0, 0.0, 0.0]
        occ.Visible = True

        occurrences = MagicMock()
        occurrences.GetOccurrence.return_value = occ
        doc.Occurrences = occurrences

        result = am.get_occurrence(42)
        assert result["internal_id"] == 42
        assert result["name"] == "Part_1"
        assert result["file_path"] == "C:/parts/part1.par"
        assert result["position"] == [0.1, 0.2, 0.3]
        occurrences.GetOccurrence.assert_called_once_with(42)

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.get_occurrence(1)
        assert "error" in result
        assert "not an assembly" in result["error"]

    def test_com_error(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.GetOccurrence.side_effect = Exception("ID not found in occurrences")
        doc.Occurrences = occurrences

        result = am.get_occurrence(99999)
        assert "error" in result
        assert "ID not found" in result["error"]
