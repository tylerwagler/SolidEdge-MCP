"""
Unit tests for AssemblyManager placement backend methods.

Tests placement mixin: PlaceComponent, AddComponentWithTransform,
AddFamilyMember, AddFamilyWithTransform, AddFamilyWithMatrix,
AddByTemplate, AddAdjustablePart, ReorderOccurrence.
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
