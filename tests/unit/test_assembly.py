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


# ============================================================================
# ASSEMBLY RELATIONS
# ============================================================================


class TestGetAssemblyRelations:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr

        rel1 = MagicMock()
        rel1.Type = 0
        rel1.Status = 1
        rel1.Suppressed = False
        rel1.Name = "Ground_1"

        rel2 = MagicMock()
        rel2.Type = 2
        rel2.Status = 1
        rel2.Suppressed = False
        rel2.Name = "Planar_1"

        relations = MagicMock()
        relations.Count = 2
        relations.Item.side_effect = lambda i: [None, rel1, rel2][i]
        doc.Relations3d = relations

        result = am.get_assembly_relations()
        assert result["count"] == 2
        assert result["relations"][0]["type_name"] == "Ground"
        assert result["relations"][1]["type_name"] == "Planar"

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d

        result = am.get_assembly_relations()
        assert "error" in result

    def test_empty(self, asm_mgr):
        am, doc = asm_mgr
        relations = MagicMock()
        relations.Count = 0
        doc.Relations3d = relations

        result = am.get_assembly_relations()
        assert result["count"] == 0
        assert result["relations"] == []


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


# ============================================================================
# TIER 3: OCCURRENCE MOVE
# ============================================================================


class TestOccurrenceMove:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.occurrence_move(0, 0.1, 0.2, 0.3)
        assert result["status"] == "moved"
        assert result["delta"] == [0.1, 0.2, 0.3]
        occ.Move.assert_called_once_with(0.1, 0.2, 0.3)

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.occurrence_move(0, 0.1, 0.0, 0.0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 2
        doc.Occurrences = occurrences

        result = am.occurrence_move(5, 0.1, 0.0, 0.0)
        assert "error" in result
        assert "Invalid" in result["error"]


# ============================================================================
# TIER 3: OCCURRENCE ROTATE
# ============================================================================


class TestOccurrenceRotate:
    def test_success(self, asm_mgr):
        import math

        am, doc = asm_mgr
        occ = MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.occurrence_rotate(0, 0, 0, 0, 0, 0, 1, 90)
        assert result["status"] == "rotated"
        assert result["angle_degrees"] == 90
        # Verify angle was converted to radians
        occ.Rotate.assert_called_once()
        call_args = occ.Rotate.call_args[0]
        assert abs(call_args[6] - math.radians(90)) < 0.001

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.occurrence_rotate(0, 0, 0, 0, 0, 0, 1, 45)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.occurrence_rotate(5, 0, 0, 0, 0, 0, 1, 90)
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
# SET COMPONENT TRANSFORM
# ============================================================================


class TestSetComponentTransform:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.set_component_transform(0, 0.1, 0.2, 0.3, 45.0, 0.0, 90.0)
        assert result["status"] == "updated"
        assert result["component_index"] == 0
        assert result["origin"] == [0.1, 0.2, 0.3]
        assert result["angles_degrees"] == [45.0, 0.0, 90.0]
        occ.PutTransform.assert_called_once()
        # Verify radians conversion
        import math

        args = occ.PutTransform.call_args[0]
        assert args[0] == 0.1
        assert args[3] == pytest.approx(math.radians(45.0))
        assert args[5] == pytest.approx(math.radians(90.0))

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.set_component_transform(0, 0, 0, 0, 0, 0, 0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.set_component_transform(5, 0, 0, 0, 0, 0, 0)
        assert "error" in result


# ============================================================================
# SET COMPONENT ORIGIN
# ============================================================================


class TestSetComponentOrigin:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.set_component_origin(0, 0.1, 0.2, 0.3)
        assert result["status"] == "updated"
        assert result["component_index"] == 0
        assert result["origin"] == [0.1, 0.2, 0.3]
        occ.PutOrigin.assert_called_once_with(0.1, 0.2, 0.3)

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.set_component_origin(0, 0, 0, 0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.set_component_origin(5, 0, 0, 0)
        assert "error" in result


# ============================================================================
# MIRROR COMPONENT
# ============================================================================


class TestMirrorComponent:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        plane = MagicMock()
        ref_planes = MagicMock()
        ref_planes.Count = 3
        ref_planes.Item.return_value = plane
        doc.RefPlanes = ref_planes

        result = am.mirror_component(0, 1)
        assert result["status"] == "mirrored"
        assert result["component_index"] == 0
        assert result["plane_index"] == 1
        occ.Mirror.assert_called_once_with(plane)

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.mirror_component(0, 1)
        assert "error" in result

    def test_invalid_component_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.mirror_component(5, 1)
        assert "error" in result

    def test_invalid_plane_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 2
        doc.Occurrences = occurrences

        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = am.mirror_component(0, 10)
        assert "error" in result


# ============================================================================
# ADD COMPONENT WITH TRANSFORM
# ============================================================================


class TestAddComponentWithTransform:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occ.Name = "Part1:1"
        occurrences = MagicMock()
        occurrences.Count = 1
        occurrences.AddWithTransform.return_value = occ
        doc.Occurrences = occurrences

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_component_with_transform("C:\\test.par", 0.1, 0.2, 0.3, 0, 0, 45)
        assert result["status"] == "added"
        assert result["name"] == "Part1:1"
        occurrences.AddWithTransform.assert_called_once()

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_component_with_transform("C:\\test.par")
        assert "error" in result

    def test_default_values(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occ.Name = "Part1:1"
        occurrences = MagicMock()
        occurrences.Count = 1
        occurrences.AddWithTransform.return_value = occ
        doc.Occurrences = occurrences

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_component_with_transform("C:\\test.par")
        assert result["status"] == "added"


# ============================================================================
# DELETE RELATION
# ============================================================================


class TestDeleteRelation:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        rel = MagicMock()
        rel.Type = 2
        relations = MagicMock()
        relations.Count = 3
        relations.Item.return_value = rel
        doc.Relations3d = relations

        result = am.delete_relation(1)
        assert result["status"] == "deleted"
        assert result["relation_index"] == 1
        rel.Delete.assert_called_once()

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d

        result = am.delete_relation(0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        relations = MagicMock()
        relations.Count = 2
        doc.Relations3d = relations

        result = am.delete_relation(5)
        assert "error" in result


# ============================================================================
# GET RELATION INFO
# ============================================================================


class TestGetRelationInfo:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        rel = MagicMock()
        rel.Type = 2
        rel.Status = 0
        rel.Name = "Planar1"
        rel.Suppressed = False
        relations = MagicMock()
        relations.Count = 3
        relations.Item.return_value = rel
        doc.Relations3d = relations

        result = am.get_relation_info(0)
        assert result["relation_index"] == 0
        assert result["name"] == "Planar1"

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d

        result = am.get_relation_info(0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        relations = MagicMock()
        relations.Count = 1
        doc.Relations3d = relations

        result = am.get_relation_info(5)
        assert "error" in result


# ============================================================================
# ADD PLANAR RELATION
# ============================================================================


class TestAddPlanarRelation:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ1, occ2 = MagicMock(), MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.side_effect = lambda i: {1: occ1, 2: occ2}[i]
        doc.Occurrences = occurrences
        relations = MagicMock()
        doc.Relations3d = relations

        result = am.add_planar_relation(0, 1, 0.01, "Align")
        assert result["status"] == "created"
        assert result["relation_type"] == "Planar"
        assert result["offset"] == 0.01
        relations.AddPlanar.assert_called_once_with(occ1, occ2, 0.01, 1)

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d
        result = am.add_planar_relation(0, 1)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences
        doc.Relations3d = MagicMock()
        result = am.add_planar_relation(0, 5)
        assert "error" in result


# ============================================================================
# ADD AXIAL RELATION
# ============================================================================


class TestAddAxialRelation:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ1, occ2 = MagicMock(), MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.side_effect = lambda i: {1: occ1, 2: occ2}[i]
        doc.Occurrences = occurrences
        relations = MagicMock()
        doc.Relations3d = relations

        result = am.add_axial_relation(0, 1, "Antialign")
        assert result["status"] == "created"
        assert result["relation_type"] == "Axial"
        assert result["orientation"] == "Antialign"
        relations.AddAxial.assert_called_once_with(occ1, occ2, 2)

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d
        result = am.add_axial_relation(0, 1)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences
        doc.Relations3d = MagicMock()
        result = am.add_axial_relation(0, 5)
        assert "error" in result


# ============================================================================
# ADD ANGULAR RELATION
# ============================================================================


class TestAddAngularRelation:
    def test_success(self, asm_mgr):
        import math

        am, doc = asm_mgr
        occ1, occ2 = MagicMock(), MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.side_effect = lambda i: {1: occ1, 2: occ2}[i]
        doc.Occurrences = occurrences
        relations = MagicMock()
        doc.Relations3d = relations

        result = am.add_angular_relation(0, 1, 45.0)
        assert result["status"] == "created"
        assert result["relation_type"] == "Angular"
        assert result["angle_degrees"] == 45.0
        relations.AddAngular.assert_called_once_with(occ1, occ2, pytest.approx(math.radians(45.0)))

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d
        result = am.add_angular_relation(0, 1, 30.0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences
        doc.Relations3d = MagicMock()
        result = am.add_angular_relation(0, 5)
        assert "error" in result


# ============================================================================
# ADD POINT RELATION
# ============================================================================


class TestAddPointRelation:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ1, occ2 = MagicMock(), MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.side_effect = lambda i: {1: occ1, 2: occ2}[i]
        doc.Occurrences = occurrences
        relations = MagicMock()
        doc.Relations3d = relations

        result = am.add_point_relation(0, 1)
        assert result["status"] == "created"
        assert result["relation_type"] == "Point"
        relations.AddPoint.assert_called_once_with(occ1, occ2)

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d
        result = am.add_point_relation(0, 1)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences
        doc.Relations3d = MagicMock()
        result = am.add_point_relation(0, 5)
        assert "error" in result


# ============================================================================
# ADD TANGENT RELATION
# ============================================================================


class TestAddTangentRelation:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ1, occ2 = MagicMock(), MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.side_effect = lambda i: {1: occ1, 2: occ2}[i]
        doc.Occurrences = occurrences
        relations = MagicMock()
        doc.Relations3d = relations

        result = am.add_tangent_relation(0, 1)
        assert result["status"] == "created"
        assert result["relation_type"] == "Tangent"
        relations.AddTangent.assert_called_once_with(occ1, occ2)

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d
        result = am.add_tangent_relation(0, 1)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences
        doc.Relations3d = MagicMock()
        result = am.add_tangent_relation(0, 5)
        assert "error" in result


# ============================================================================
# ADD GEAR RELATION
# ============================================================================


class TestAddGearRelation:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ1, occ2 = MagicMock(), MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.side_effect = lambda i: {1: occ1, 2: occ2}[i]
        doc.Occurrences = occurrences
        relations = MagicMock()
        doc.Relations3d = relations

        result = am.add_gear_relation(0, 1, 2.0, 3.0)
        assert result["status"] == "created"
        assert result["relation_type"] == "Gear"
        assert result["ratio1"] == 2.0
        assert result["ratio2"] == 3.0
        relations.AddGear.assert_called_once_with(occ1, occ2, 2.0, 3.0)

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d
        result = am.add_gear_relation(0, 1)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences
        doc.Relations3d = MagicMock()
        result = am.add_gear_relation(0, 5)
        assert "error" in result


# ============================================================================
# GET RELATION OFFSET
# ============================================================================


class TestGetRelationOffset:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        rel = MagicMock()
        rel.Offset = 0.025
        relations = MagicMock()
        relations.Count = 2
        relations.Item.return_value = rel
        doc.Relations3d = relations

        result = am.get_relation_offset(0)
        assert result["relation_index"] == 0
        assert result["offset"] == 0.025

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d
        result = am.get_relation_offset(0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        relations = MagicMock()
        relations.Count = 1
        doc.Relations3d = relations
        result = am.get_relation_offset(5)
        assert "error" in result


# ============================================================================
# SET RELATION OFFSET
# ============================================================================


class TestSetRelationOffset:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        rel = MagicMock()
        relations = MagicMock()
        relations.Count = 2
        relations.Item.return_value = rel
        doc.Relations3d = relations

        result = am.set_relation_offset(0, 0.05)
        assert result["status"] == "updated"
        assert result["offset"] == 0.05
        assert rel.Offset == 0.05

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d
        result = am.set_relation_offset(0, 0.01)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        relations = MagicMock()
        relations.Count = 1
        doc.Relations3d = relations
        result = am.set_relation_offset(5, 0.01)
        assert "error" in result


# ============================================================================
# GET RELATION ANGLE
# ============================================================================


class TestGetRelationAngle:
    def test_success(self, asm_mgr):
        import math

        am, doc = asm_mgr
        rel = MagicMock()
        rel.Angle = math.radians(90.0)
        relations = MagicMock()
        relations.Count = 2
        relations.Item.return_value = rel
        doc.Relations3d = relations

        result = am.get_relation_angle(0)
        assert result["relation_index"] == 0
        assert result["angle_degrees"] == pytest.approx(90.0)
        assert result["angle_radians"] == pytest.approx(math.radians(90.0))

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d
        result = am.get_relation_angle(0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        relations = MagicMock()
        relations.Count = 1
        doc.Relations3d = relations
        result = am.get_relation_angle(5)
        assert "error" in result


# ============================================================================
# SET RELATION ANGLE
# ============================================================================


class TestSetRelationAngle:
    def test_success(self, asm_mgr):
        import math

        am, doc = asm_mgr
        rel = MagicMock()
        relations = MagicMock()
        relations.Count = 2
        relations.Item.return_value = rel
        doc.Relations3d = relations

        result = am.set_relation_angle(0, 45.0)
        assert result["status"] == "updated"
        assert result["angle_degrees"] == 45.0
        assert rel.Angle == pytest.approx(math.radians(45.0))

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d
        result = am.set_relation_angle(0, 30.0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        relations = MagicMock()
        relations.Count = 1
        doc.Relations3d = relations
        result = am.set_relation_angle(5, 45.0)
        assert "error" in result


# ============================================================================
# GET NORMALS ALIGNED
# ============================================================================


class TestGetNormalsAligned:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        rel = MagicMock()
        rel.NormalsAligned = True
        relations = MagicMock()
        relations.Count = 2
        relations.Item.return_value = rel
        doc.Relations3d = relations

        result = am.get_normals_aligned(0)
        assert result["relation_index"] == 0
        assert result["normals_aligned"] is True

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d
        result = am.get_normals_aligned(0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        relations = MagicMock()
        relations.Count = 1
        doc.Relations3d = relations
        result = am.get_normals_aligned(5)
        assert "error" in result


# ============================================================================
# SET NORMALS ALIGNED
# ============================================================================


class TestSetNormalsAligned:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        rel = MagicMock()
        relations = MagicMock()
        relations.Count = 2
        relations.Item.return_value = rel
        doc.Relations3d = relations

        result = am.set_normals_aligned(0, False)
        assert result["status"] == "updated"
        assert result["normals_aligned"] is False
        assert rel.NormalsAligned is False

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d
        result = am.set_normals_aligned(0, True)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        relations = MagicMock()
        relations.Count = 1
        doc.Relations3d = relations
        result = am.set_normals_aligned(5, True)
        assert "error" in result


# ============================================================================
# SUPPRESS RELATION
# ============================================================================


class TestSuppressRelation:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        rel = MagicMock()
        relations = MagicMock()
        relations.Count = 2
        relations.Item.return_value = rel
        doc.Relations3d = relations

        result = am.suppress_relation(0)
        assert result["status"] == "suppressed"
        assert result["relation_index"] == 0
        assert rel.Suppressed is True

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d
        result = am.suppress_relation(0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        relations = MagicMock()
        relations.Count = 1
        doc.Relations3d = relations
        result = am.suppress_relation(5)
        assert "error" in result


# ============================================================================
# UNSUPPRESS RELATION
# ============================================================================


class TestUnsuppressRelation:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        rel = MagicMock()
        relations = MagicMock()
        relations.Count = 2
        relations.Item.return_value = rel
        doc.Relations3d = relations

        result = am.unsuppress_relation(0)
        assert result["status"] == "unsuppressed"
        assert result["relation_index"] == 0
        assert rel.Suppressed is False

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d
        result = am.unsuppress_relation(0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        relations = MagicMock()
        relations.Count = 1
        doc.Relations3d = relations
        result = am.unsuppress_relation(5)
        assert "error" in result


# ============================================================================
# GET RELATION GEOMETRY
# ============================================================================


class TestGetRelationGeometry:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ1 = MagicMock()
        occ1.Name = "Part_1"
        occ2 = MagicMock()
        occ2.Name = "Part_2"
        rel = MagicMock()
        rel.Type = 2
        rel.Name = "Planar_1"
        rel.OccurrencePart1 = occ1
        rel.OccurrencePart2 = occ2
        rel.Offset = 0.01
        rel.NormalsAligned = True
        rel.Suppressed = False
        relations = MagicMock()
        relations.Count = 2
        relations.Item.return_value = rel
        doc.Relations3d = relations

        result = am.get_relation_geometry(0)
        assert result["relation_index"] == 0
        assert result["occurrence1_name"] == "Part_1"
        assert result["occurrence2_name"] == "Part_2"
        assert result["offset"] == 0.01
        assert result["normals_aligned"] is True

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d
        result = am.get_relation_geometry(0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        relations = MagicMock()
        relations.Count = 1
        doc.Relations3d = relations
        result = am.get_relation_geometry(5)
        assert "error" in result


# ============================================================================
# GET GEAR RATIO
# ============================================================================


class TestGetGearRatio:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        rel = MagicMock()
        rel.RatioValue1 = 2.0
        rel.RatioValue2 = 3.0
        relations = MagicMock()
        relations.Count = 2
        relations.Item.return_value = rel
        doc.Relations3d = relations

        result = am.get_gear_ratio(0)
        assert result["relation_index"] == 0
        assert result["ratio1"] == 2.0
        assert result["ratio2"] == 3.0

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Relations3d
        result = am.get_gear_ratio(0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        relations = MagicMock()
        relations.Count = 1
        doc.Relations3d = relations
        result = am.get_gear_ratio(5)
        assert "error" in result


# ============================================================================
# BATCH 8: ADD FAMILY MEMBER
# ============================================================================


class TestAddFamilyMember:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occ.Name = "Bolt_M10:1"
        occurrences = MagicMock()
        occurrences.Count = 1
        occurrences.AddFamilyByFilename.return_value = occ
        doc.Occurrences = occurrences

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_family_member("C:\\parts\\bolt_family.par", "M10")
        assert result["status"] == "added"
        assert result["family_member"] == "M10"
        assert result["name"] == "Bolt_M10:1"
        occurrences.AddFamilyByFilename.assert_called_once_with("C:\\parts\\bolt_family.par", "M10")

    def test_file_not_found(self, asm_mgr):
        am, doc = asm_mgr
        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=False):
            result = am.add_family_member("C:\\nonexistent.par", "M10")
        assert "error" in result
        assert "File not found" in result["error"]

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences
        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_family_member("C:\\parts\\bolt.par", "M10")
        assert "error" in result


# ============================================================================
# BATCH 8: ADD FAMILY WITH TRANSFORM
# ============================================================================


class TestAddFamilyWithTransform:
    def test_success(self, asm_mgr):
        import math

        am, doc = asm_mgr
        occ = MagicMock()
        occ.Name = "Bolt_M10:1"
        occurrences = MagicMock()
        occurrences.Count = 1
        occurrences.AddFamilyByFilename.return_value = occ
        doc.Occurrences = occurrences

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_family_with_transform(
                "C:\\parts\\bolt.par", "M10", 0.1, 0.2, 0.3, 45, 0, 90
            )
        assert result["status"] == "added"
        assert result["origin"] == [0.1, 0.2, 0.3]
        assert result["angles_degrees"] == [45, 0, 90]
        occ.PutTransform.assert_called_once_with(
            0.1,
            0.2,
            0.3,
            pytest.approx(math.radians(45)),
            pytest.approx(math.radians(0)),
            pytest.approx(math.radians(90)),
        )

    def test_file_not_found(self, asm_mgr):
        am, doc = asm_mgr
        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=False):
            result = am.add_family_with_transform("C:\\nonexistent.par", "M10")
        assert "error" in result
        assert "File not found" in result["error"]

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences
        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_family_with_transform("C:\\parts\\bolt.par", "M10")
        assert "error" in result


# ============================================================================
# BATCH 8: ADD BY TEMPLATE
# ============================================================================


class TestAddByTemplate:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occ.Name = "TemplatedPart:1"
        occurrences = MagicMock()
        occurrences.Count = 1
        occurrences.AddByTemplate.return_value = occ
        doc.Occurrences = occurrences

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_by_template("C:\\parts\\part.par", "MyTemplate")
        assert result["status"] == "added"
        assert result["template_name"] == "MyTemplate"
        assert result["name"] == "TemplatedPart:1"
        occurrences.AddByTemplate.assert_called_once_with("C:\\parts\\part.par", "MyTemplate")

    def test_file_not_found(self, asm_mgr):
        am, doc = asm_mgr
        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=False):
            result = am.add_by_template("C:\\nonexistent.par", "MyTemplate")
        assert "error" in result
        assert "File not found" in result["error"]

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences
        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_by_template("C:\\parts\\part.par", "MyTemplate")
        assert "error" in result


# ============================================================================
# BATCH 8: ADD ADJUSTABLE PART
# ============================================================================


class TestAddAdjustablePart:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occ.Name = "AdjPart:1"
        occurrences = MagicMock()
        occurrences.Count = 1
        occurrences.AddAsAdjustablePart.return_value = occ
        doc.Occurrences = occurrences

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_adjustable_part("C:\\parts\\spring.par")
        assert result["status"] == "added"
        assert result["adjustable"] is True
        assert result["name"] == "AdjPart:1"
        occurrences.AddAsAdjustablePart.assert_called_once_with("C:\\parts\\spring.par")

    def test_file_not_found(self, asm_mgr):
        am, doc = asm_mgr
        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=False):
            result = am.add_adjustable_part("C:\\nonexistent.par")
        assert "error" in result
        assert "File not found" in result["error"]

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences
        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_adjustable_part("C:\\parts\\spring.par")
        assert "error" in result


# ============================================================================
# BATCH 8: REORDER OCCURRENCE
# ============================================================================


class TestReorderOccurrence:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 3
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.reorder_occurrence(0, 2)
        assert result["status"] == "reordered"
        assert result["component_index"] == 0
        assert result["target_index"] == 2
        occurrences.ReorderOccurrence.assert_called_once_with(occ, 3)

    def test_invalid_component_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 2
        doc.Occurrences = occurrences

        result = am.reorder_occurrence(5, 0)
        assert "error" in result
        assert "Invalid component index" in result["error"]

    def test_invalid_target_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 2
        doc.Occurrences = occurrences

        result = am.reorder_occurrence(0, 5)
        assert "error" in result
        assert "Invalid target index" in result["error"]


# ============================================================================
# BATCH 8: PUT TRANSFORM EULER
# ============================================================================


class TestPutTransformEuler:
    def test_success(self, asm_mgr):
        import math

        am, doc = asm_mgr
        occ = MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.put_transform_euler(0, 0.1, 0.2, 0.3, 45.0, 0.0, 90.0)
        assert result["status"] == "updated"
        assert result["position"] == [0.1, 0.2, 0.3]
        assert result["angles_degrees"] == [45.0, 0.0, 90.0]
        occ.PutTransform.assert_called_once()
        args = occ.PutTransform.call_args[0]
        assert args[0] == 0.1
        assert args[3] == pytest.approx(math.radians(45.0))
        assert args[5] == pytest.approx(math.radians(90.0))

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.put_transform_euler(0, 0, 0, 0, 0, 0, 0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.put_transform_euler(5, 0, 0, 0, 0, 0, 0)
        assert "error" in result


# ============================================================================
# BATCH 8: PUT ORIGIN
# ============================================================================


class TestPutOrigin:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.put_origin(0, 0.1, 0.2, 0.3)
        assert result["status"] == "updated"
        assert result["origin"] == [0.1, 0.2, 0.3]
        occ.PutOrigin.assert_called_once_with(0.1, 0.2, 0.3)

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.put_origin(0, 0, 0, 0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.put_origin(5, 0, 0, 0)
        assert "error" in result


# ============================================================================
# BATCH 8: MAKE WRITABLE
# ============================================================================


class TestMakeWritable:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.make_writable(0)
        assert result["status"] == "writable"
        assert result["component_index"] == 0
        occ.MakeWritable.assert_called_once()

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.make_writable(0)
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.make_writable(5)
        assert "error" in result


# ============================================================================
# BATCH 8: SWAP FAMILY MEMBER
# ============================================================================


class TestSwapFamilyMember:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.swap_family_member(0, "M12")
        assert result["status"] == "swapped"
        assert result["new_member_name"] == "M12"
        occ.SwapFamilyMember.assert_called_once_with("M12")

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.swap_family_member(0, "M12")
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.swap_family_member(5, "M12")
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
# BATCH 8: IS TUBE
# ============================================================================


class TestIsTube:
    def test_true(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occ.IsTube = True
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.is_tube(0)
        assert result["component_index"] == 0
        assert result["is_tube"] is True

    def test_false(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occ.IsTube = False
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.is_tube(0)
        assert result["is_tube"] is False

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.is_tube(5)
        assert "error" in result


# ============================================================================
# BATCH 8: GET ADJUSTABLE PART
# ============================================================================


class TestGetAdjustablePart:
    def test_adjustable(self, asm_mgr):
        am, doc = asm_mgr
        adj_part = MagicMock()
        adj_part.Name = "Spring_Adj"
        occ = MagicMock()
        occ.GetAdjustablePart.return_value = adj_part
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.get_adjustable_part(0)
        assert result["is_adjustable"] is True
        assert result["adjustable_name"] == "Spring_Adj"

    def test_not_adjustable(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occ.GetAdjustablePart.return_value = None
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.get_adjustable_part(0)
        assert result["is_adjustable"] is False

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.get_adjustable_part(5)
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
# ADD FAMILY WITH MATRIX
# ============================================================================


class TestAddFamilyWithMatrix:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occ.Name = "FamilyPart_1"

        occurrences = MagicMock()
        occurrences.AddFamilyWithMatrix.return_value = occ
        occurrences.Count = 1
        doc.Occurrences = occurrences

        matrix = [1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0.1, 0.2, 0.3, 1]

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_family_with_matrix("C:/parts/family.par", "MemberA", matrix)

        assert result["status"] == "added"
        assert result["family_member"] == "MemberA"
        assert result["name"] == "FamilyPart_1"
        assert result["position"] == [0.1, 0.2, 0.3]
        occurrences.AddFamilyWithMatrix.assert_called_once_with(
            "C:/parts/family.par", matrix, "MemberA"
        )

    def test_invalid_matrix_length(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        doc.Occurrences = occurrences

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_family_with_matrix("C:/parts/family.par", "MemberA", [1, 0, 0])

        assert "error" in result
        assert "16 elements" in result["error"]

    def test_file_not_found(self, asm_mgr):
        am, doc = asm_mgr

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=False):
            result = am.add_family_with_matrix(
                "C:/nonexistent.par",
                "MemberA",
                [1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1],
            )

        assert "error" in result
        assert "File not found" in result["error"]


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
