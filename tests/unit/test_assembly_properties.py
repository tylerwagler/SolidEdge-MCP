"""
Unit tests for AssemblyManager properties backend methods.

Tests properties mixin: SuppressComponent, IsTube, GetAdjustablePart,
MakeWritable, SwapFamilyMember.
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
