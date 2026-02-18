"""
Unit tests for AssemblyManager relations backend methods.

Tests relations mixin: GetAssemblyRelations, DeleteRelation, GetRelationInfo,
AddPlanarRelation, AddAxialRelation, AddAngularRelation,
AddPointRelation, AddTangentRelation, AddGearRelation,
GetRelationOffset, SetRelationOffset, GetRelationAngle, SetRelationAngle,
GetNormalsAligned, SetNormalsAligned,
SuppressRelation, UnsuppressRelation,
GetRelationGeometry, GetGearRatio.
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
