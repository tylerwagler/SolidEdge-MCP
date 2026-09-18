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
        doc.Type = IG_PART_DOCUMENT

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
        doc.Type = IG_PART_DOCUMENT

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
        doc.Type = IG_PART_DOCUMENT

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


def _assembly_with_faces(doc, n_occurrences=2, radius=None):
    """Occurrences whose part documents expose one face each, with a range for the midpoint."""
    doc.Type = 3  # igAssemblyDocument
    doc.Occurrences.Count = n_occurrences
    occurrences, faces, refs = [], [], {}
    for i in range(n_occurrences):
        occ = MagicMock(name=f"occ{i}")
        face = MagicMock(name=f"face{i}")
        face.Geometry.Radius = radius
        face.GetRange.return_value = (0.0, 0.0, float(i), 0.08, 0.048, float(i))
        faces_coll = occ.OccurrenceDocument.Models.Item.return_value.Body.Faces.return_value
        faces_coll.Count = 1
        faces_coll.Item.return_value = face
        occ.OccurrenceDocument.Models.Count = 1
        occurrences.append(occ)
        faces.append(face)
    doc.Occurrences.Item.side_effect = lambda i: occurrences[i - 1]
    doc.CreateReference.side_effect = lambda occ, face: refs.setdefault(
        id(face), MagicMock(name="ref")
    )
    doc.Relations3d.Count = 1
    return occurrences, faces, refs


class TestAddPlanarRelation:
    """AddPlanar on References from CreateReference, with on-face constraining points."""

    def test_mates_two_faces_through_references(self, asm_mgr):
        am, doc = asm_mgr
        occs, faces, refs = _assembly_with_faces(doc)

        result = am.add_planar_relation(0, 1, orientation="Antialign")

        assert result["status"] == "created"
        assert result["mated"] is True
        doc.CreateReference.assert_any_call(occs[0], faces[0])
        doc.CreateReference.assert_any_call(occs[1], faces[1])
        doc.Relations3d.AddPlanar.assert_called_once_with(
            refs[id(faces[0])], refs[id(faces[1])], True, [0.04, 0.024, 0.0], [0.04, 0.024, 1.0]
        )

    def test_align_passes_normals_not_aligned(self, asm_mgr):
        am, doc = asm_mgr
        _assembly_with_faces(doc)
        am.add_planar_relation(0, 1, orientation="Align")
        assert doc.Relations3d.AddPlanar.call_args.args[2] is False

    def test_a_curved_face_or_offset_is_refused(self, asm_mgr):
        am, doc = asm_mgr
        _assembly_with_faces(doc, radius=0.01)
        assert "planar" in am.add_planar_relation(0, 1)["error"]
        _assembly_with_faces(doc)
        assert "offset" in am.add_planar_relation(0, 1, offset=0.01)["error"]
        doc.Relations3d.AddPlanar.assert_not_called()

    def test_bad_indices_are_refused(self, asm_mgr):
        am, doc = asm_mgr
        _assembly_with_faces(doc)
        assert "Invalid occurrence index" in am.add_planar_relation(0, 5)["error"]
        assert "Invalid face index" in am.add_planar_relation(0, 1, face2_index=3)["error"]
        doc.Relations3d.AddPlanar.assert_not_called()


# ============================================================================
# ADD AXIAL RELATION
# ============================================================================


class TestAddAxialRelation:
    """AddAxial on References to cylindrical faces."""

    def test_makes_two_cylinders_coaxial(self, asm_mgr):
        am, doc = asm_mgr
        occs, faces, refs = _assembly_with_faces(doc, radius=0.01)

        result = am.add_axial_relation(0, 1, orientation="Antialign")

        assert result["status"] == "created"
        assert result["relation_type"] == "Axial"
        doc.Relations3d.AddAxial.assert_called_once_with(
            refs[id(faces[0])], refs[id(faces[1])], False
        )

    def test_a_planar_face_is_refused(self, asm_mgr):
        am, doc = asm_mgr
        _assembly_with_faces(doc, radius=None)
        assert "cylindrical" in am.add_axial_relation(0, 1)["error"]
        doc.Relations3d.AddAxial.assert_not_called()


class TestConstraintsDelegate:
    def test_mate_is_a_planar_relation_with_normals_aligned(self, asm_mgr):
        am, doc = asm_mgr
        _assembly_with_faces(doc)
        result = am.create_mate("Mate", 0, 1, face1_index=0, face2_index=0)
        assert result["status"] == "created" and result["mate_type"] == "Mate"
        assert doc.Relations3d.AddPlanar.call_args.args[2] is True

    def test_align_constraints_are_planar_aligned(self, asm_mgr):
        am, doc = asm_mgr
        _assembly_with_faces(doc)
        assert am.add_align_constraint(0, 1)["status"] == "created"
        assert am.add_planar_align_constraint(0, 1)["status"] == "created"
        assert all(c.args[2] is False for c in doc.Relations3d.AddPlanar.call_args_list)

    def test_axial_align_is_an_axial_relation(self, asm_mgr):
        am, doc = asm_mgr
        _assembly_with_faces(doc, radius=0.01)
        assert am.add_axial_align_constraint(0, 1)["status"] == "created"
        doc.Relations3d.AddAxial.assert_called_once()


# ============================================================================
# ADD ANGULAR RELATION
# ============================================================================


class TestAddAngularRelation:
    """AddAngular needs Face/Edge geometry this server cannot select."""

    def test_unsupported_does_not_call_com(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 2
        doc.Occurrences = occurrences
        relations = MagicMock()
        doc.Relations3d = relations

        result = am.add_angular_relation(0, 1, 45.0)

        assert result["unsupported"] is True
        assert "measurement elements" in result["error"].lower()
        relations.AddAngular.assert_not_called()
        occurrences.Item.assert_not_called()

    def test_unsupported_even_without_assembly(self, asm_mgr):
        am, doc = asm_mgr
        doc.Type = IG_PART_DOCUMENT

        result = am.add_angular_relation(0, 1, 45.0)

        assert result["unsupported"] is True
        doc.Relations3d.AddAngular.assert_not_called()


# ============================================================================
# ADD POINT RELATION
# ============================================================================


class TestAddPointRelation:
    """AddPoint needs Face/Edge geometry this server cannot select."""

    def test_unsupported_does_not_call_com(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 2
        doc.Occurrences = occurrences
        relations = MagicMock()
        doc.Relations3d = relations

        result = am.add_point_relation(0, 1)

        assert result["unsupported"] is True
        assert "keypoint" in result["error"].lower()
        relations.AddPoint.assert_not_called()
        occurrences.Item.assert_not_called()

    def test_unsupported_even_without_assembly(self, asm_mgr):
        am, doc = asm_mgr
        doc.Type = IG_PART_DOCUMENT

        result = am.add_point_relation(0, 1)

        assert result["unsupported"] is True
        doc.Relations3d.AddPoint.assert_not_called()


# ============================================================================
# ADD TANGENT RELATION
# ============================================================================


class TestAddTangentRelation:
    """AddTangent needs Face/Edge geometry this server cannot select."""

    def test_unsupported_does_not_call_com(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 2
        doc.Occurrences = occurrences
        relations = MagicMock()
        doc.Relations3d = relations

        result = am.add_tangent_relation(0, 1)

        assert result["unsupported"] is True
        assert "tangent" in result["error"].lower()
        relations.AddTangent.assert_not_called()
        occurrences.Item.assert_not_called()

    def test_unsupported_even_without_assembly(self, asm_mgr):
        am, doc = asm_mgr
        doc.Type = IG_PART_DOCUMENT

        result = am.add_tangent_relation(0, 1)

        assert result["unsupported"] is True
        doc.Relations3d.AddTangent.assert_not_called()


# ============================================================================
# ADD GEAR RELATION
# ============================================================================


class TestAddGearRelation:
    """AddGear needs Face/Edge geometry this server cannot select."""

    def test_unsupported_does_not_call_com(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 2
        doc.Occurrences = occurrences
        relations = MagicMock()
        doc.Relations3d = relations

        result = am.add_gear_relation(0, 1, 2.0, 1.0)

        assert result["unsupported"] is True
        assert "geartype" in result["error"].lower()
        relations.AddGear.assert_not_called()
        occurrences.Item.assert_not_called()

    def test_unsupported_even_without_assembly(self, asm_mgr):
        am, doc = asm_mgr
        doc.Type = IG_PART_DOCUMENT

        result = am.add_gear_relation(0, 1, 2.0, 1.0)

        assert result["unsupported"] is True
        doc.Relations3d.AddGear.assert_not_called()


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
        doc.Type = IG_PART_DOCUMENT
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
        doc.Type = IG_PART_DOCUMENT
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
        doc.Type = IG_PART_DOCUMENT
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
        doc.Type = IG_PART_DOCUMENT
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
        doc.Type = IG_PART_DOCUMENT
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
        doc.Type = IG_PART_DOCUMENT
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
        doc.Type = IG_PART_DOCUMENT
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
        doc.Type = IG_PART_DOCUMENT
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
        rel.Occurrence1 = occ1
        rel.Occurrence2 = occ2
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
        doc.Type = IG_PART_DOCUMENT
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
        doc.Type = IG_PART_DOCUMENT
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
# DOCUMENT-TYPE GUARD (replaces the old hasattr probe)
# ============================================================================


class TestAssemblyDocumentGuard:
    """The guard reads Document.Type instead of probing for a member.

    hasattr on a late-bound COM proxy also returns False when the member
    exists but its getter raises, which reported unrelated COM failures as
    "Active document is not an assembly".
    """

    def test_weldment_assembly_is_accepted(self, asm_mgr):
        am, doc = asm_mgr
        doc.Type = DocumentTypeConstants.igWeldmentAssemblyDocument
        relations = MagicMock()
        relations.Count = 0
        doc.Relations3d = relations

        result = am.get_assembly_relations()
        assert result["count"] == 0

    def test_part_document_is_rejected(self, asm_mgr):
        am, doc = asm_mgr
        doc.Type = IG_PART_DOCUMENT

        result = am.get_assembly_relations()
        assert result["error"] == "Active document is not an assembly"

    def test_raising_type_getter_is_rejected(self, asm_mgr):
        am, doc = asm_mgr
        type(doc).Type = property(lambda self: (_ for _ in ()).throw(Exception("gone")))
        try:
            result = am.get_assembly_relations()
            assert result["error"] == "Active document is not an assembly"
        finally:
            del type(doc).Type

    def test_raising_relations_getter_surfaces_the_real_error(self, asm_mgr):
        am, doc = asm_mgr
        type(doc).Relations3d = property(
            lambda self: (_ for _ in ()).throw(Exception("Relations3d exploded"))
        )
        try:
            result = am.get_assembly_relations()
            # The old hasattr probe swallowed this as "not an assembly".
            assert "error" in result
            assert result["error"] != "Active document is not an assembly"
        finally:
            del type(doc).Relations3d


class TestAddAxialRelationOrientation:
    """AddAxial's third argument is NormalsAligned (VT_BOOL), not an enum."""

    def test_align_is_true(self, asm_mgr):
        am, doc = asm_mgr
        _assembly_with_faces(doc, radius=0.01)
        am.add_axial_relation(0, 1, "Align")
        assert doc.Relations3d.AddAxial.call_args.args[2] is True

    def test_antialign_is_false(self, asm_mgr):
        am, doc = asm_mgr
        _assembly_with_faces(doc, radius=0.01)
        am.add_axial_relation(0, 1, "Antialign")
        assert doc.Relations3d.AddAxial.call_args.args[2] is False
