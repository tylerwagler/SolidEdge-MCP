"""
Unit tests for AssemblyManager transforms backend methods.

Tests transforms mixin: SetComponentTransform, SetComponentOrigin,
OccurrenceMove, OccurrenceRotate, MirrorComponent,
PutTransformEuler, PutOrigin.
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
