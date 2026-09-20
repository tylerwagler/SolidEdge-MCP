"""
Unit tests for AssemblyManager transforms backend methods.

Tests transforms mixin: SetComponentTransform, SetComponentOrigin,
OccurrenceMove, OccurrenceRotate, MirrorComponent,
PutTransformEuler, PutOrigin.
Uses unittest.mock to simulate COM objects.
"""

from unittest.mock import MagicMock

import pytest

from solidedge_mcp.backends.constants import DocumentTypeConstants

IG_ASSEMBLY_DOCUMENT = DocumentTypeConstants.igAssemblyDocument
IG_DRAFT_DOCUMENT = DocumentTypeConstants.igDraftDocument
IG_PART_DOCUMENT = DocumentTypeConstants.igPartDocument


@pytest.fixture
def asm_mgr():
    """Create AssemblyManager with mocked dependencies."""
    from solidedge_mcp.backends.assembly import AssemblyManager

    dm = MagicMock()
    doc = MagicMock()
    doc.Type = IG_ASSEMBLY_DOCUMENT
    dm.get_active_document.return_value = doc
    return AssemblyManager(dm), doc


@pytest.fixture
def asm_mgr_with_sketch():
    """Create AssemblyManager with mocked doc and sketch manager."""
    from solidedge_mcp.backends.assembly import AssemblyManager

    dm = MagicMock()
    sm = MagicMock()
    doc = MagicMock()
    doc.Type = IG_ASSEMBLY_DOCUMENT
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
        doc.Type = IG_PART_DOCUMENT

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
        doc.Type = IG_PART_DOCUMENT

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
        doc.Type = IG_PART_DOCUMENT

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
        doc.Type = IG_PART_DOCUMENT

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
        doc.Type = IG_PART_DOCUMENT

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
        doc.Type = IG_PART_DOCUMENT

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
        doc.Type = IG_PART_DOCUMENT

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
# UPDATE COMPONENT POSITION (Occurrence.GetMatrix / PutMatrix)
# ============================================================================


#: A recognisable 4x4 row-major transform: identity rotation, translated to
#: (9, 8, 7). Returned by the mocked GetMatrix so the test can prove the
#: rotation block survives and only the translation row is rewritten.
CURRENT_MATRIX = [
    1.0, 0.0, 0.0, 0.0,
    0.0, 1.0, 0.0, 0.0,
    0.0, 0.0, 1.0, 0.0,
    9.0, 8.0, 7.0, 1.0,
]  # fmt: skip


class TestUpdateComponentPosition:
    """GetMatrix takes an [in, out] array; PutMatrix takes (Matrix, Replace)."""

    @pytest.fixture
    def occ_mgr(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        # Verified on SE 2026: Occurrence.GetMatrix takes a plain 16-element
        # list for its [in,out] SAFEARRAY(VT_R8) parameter and returns the
        # filled matrix; the list passed in is NOT updated in place, so the
        # backend reads the return value and the mock must supply a real one.
        occ.GetMatrix.return_value = list(CURRENT_MATRIX)
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences
        return am, doc, occ

    def test_get_matrix_is_given_its_inout_array(self, occ_mgr):
        am, doc, occ = occ_mgr

        result = am.update_component_position(0, 0.1, 0.2, 0.3)

        assert result["status"] == "position_updated"
        # GetMatrix(Matrix as SAFEARRAY(VT_R8)*) is [in, out]: calling it with
        # no argument raises "Parameter not optional" (0x8002000F). The buffer
        # is a plain list of 16 floats -- a VARIANT wrapper is rejected with
        # "Objects for SAFEARRAYS must be sequences".
        assert occ.GetMatrix.call_count == 1
        assert not occ.GetMatrix.call_args.kwargs
        assert occ.GetMatrix.call_args.args == ([0.0] * 16,)
        (buffer,) = occ.GetMatrix.call_args.args
        assert isinstance(buffer, list)
        assert all(isinstance(v, float) for v in buffer)

    def test_put_matrix_gets_matrix_and_replace_flag(self, occ_mgr):
        am, doc, occ = occ_mgr

        am.update_component_position(0, 0.1, 0.2, 0.3)

        # PutMatrix(Matrix as SAFEARRAY(VT_R8)*, Replace as VT_BOOL).
        # The matrix GetMatrix returned flows through with only the row-major
        # translation slots (12, 13, 14) rewritten.
        expected = list(CURRENT_MATRIX)
        expected[12], expected[13], expected[14] = 0.1, 0.2, 0.3
        occ.PutMatrix.assert_called_once_with(expected, True)

    def test_com_failure_is_reported(self, occ_mgr):
        am, doc, occ = occ_mgr
        occ.PutMatrix.side_effect = Exception("grounded")

        result = am.update_component_position(0, 0.1, 0.2, 0.3)

        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.update_component_position(5, 0, 0, 0)
        assert "error" in result
