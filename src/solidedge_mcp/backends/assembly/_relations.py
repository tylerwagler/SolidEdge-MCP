"""Assembly relation (constraint) operations."""

import contextlib
import math
from typing import Any

from solidedge_mcp.backends.errors import error_result

from ..constants import FaceQueryConstants
from ..features._base import verifies_collection_growth
from ..logging import get_logger
from ._base import com_get

_logger = get_logger(__name__)


class RelationsMixin:
    """Mixin providing assembly relation/constraint methods."""

    def _face_reference(
        self, doc: Any, occurrence_index: int, face_index: int, cylindrical: bool = False
    ) -> tuple[Any, list[float]] | dict[str, Any]:
        """A Reference to a face of an occurrence, plus a point on that face.

        Relations take Reference objects, not faces: ``Relations3d.AddPlanar``
        answers 0x80040225 to faces read from the occurrence's part document
        (and to its RefPlanes and the assembly's own), and ``Occurrence.Body``
        raises. ``AssemblyDocument.CreateReference(Occurrence, Entity)`` wraps
        the part-document face with the path to it -- the "Working with
        References" route in Siemens' reference -- and with that both
        ``AddPlanar`` and ``AddAxial`` build (Solid Edge 2026). The constraining
        points of a planar relation must lie on the faces; the face range's
        midpoint, in the part's coordinates, is what worked.
        """
        occurrences = doc.Occurrences
        count = com_get(occurrences, "Count", 0)
        if occurrence_index < 0 or occurrence_index >= count:
            return {"error": f"Invalid occurrence index: {occurrence_index}. Count: {count}"}
        occurrence = occurrences.Item(occurrence_index + 1)
        part = com_get(occurrence, "OccurrenceDocument")
        models = com_get(part, "Models")
        if models is None or not com_get(models, "Count", 0):
            return {
                "error": f"Occurrence {occurrence_index} has no solid body to take a face from."
            }
        faces = models.Item(1).Body.Faces(FaceQueryConstants.igQueryAll)
        n_faces = com_get(faces, "Count", 0)
        if face_index < 0 or face_index >= n_faces:
            return {
                "error": (
                    f"Invalid face index {face_index} for occurrence {occurrence_index}: "
                    f"its body has {n_faces} faces (0-based)."
                )
            }
        face = faces.Item(face_index + 1)
        radius = com_get(com_get(face, "Geometry"), "Radius")
        if cylindrical and radius is None:
            return {
                "error": (
                    f"Face {face_index} of occurrence {occurrence_index} is not cylindrical; "
                    "an axial relation needs a cylinder or cone on each part."
                )
            }
        if not cylindrical and radius is not None:
            return {
                "error": (
                    f"Face {face_index} of occurrence {occurrence_index} is curved; a planar "
                    "relation needs a planar face on each part."
                )
            }
        reference = doc.CreateReference(occurrence, face)
        r = face.GetRange([0.0] * 6)
        flat: list[float] = []
        for item in r:
            flat.extend(item if isinstance(item, tuple) else (item,))
        midpoint = [(flat[0] + flat[3]) / 2, (flat[1] + flat[4]) / 2, (flat[2] + flat[5]) / 2]
        return reference, midpoint

    @verifies_collection_growth("Relations3d")
    def create_mate(
        self,
        mate_type: str,
        component1_index: int,
        component2_index: int,
        face1_index: int = 0,
        face2_index: int = 0,
    ) -> dict[str, Any]:
        """Mate a planar face of one component to a planar face of another.

        A mate is a planar relation with NormalsAligned True; see
        ``add_planar_relation``. mate_type is echoed.
        """
        result = self.add_planar_relation(
            component1_index,
            component2_index,
            orientation="Antialign",
            face1_index=face1_index,
            face2_index=face2_index,
        )
        result.setdefault("mate_type", mate_type)
        return result

    def add_align_constraint(
        self,
        component1_index: int,
        component2_index: int,
        face1_index: int = 0,
        face2_index: int = 0,
    ) -> dict[str, Any]:
        """Align a planar face of one component with one of another (see add_planar_relation)."""
        return self.add_planar_relation(
            component1_index,
            component2_index,
            orientation="Align",
            face1_index=face1_index,
            face2_index=face2_index,
        )

    def add_angle_constraint(
        self, component1_index: int, component2_index: int, angle: float
    ) -> dict[str, Any]:
        """Add an angle constraint between two components (requires UI for face selection)"""
        return {
            "error": "Constraint creation requires face/edge selection. Use Solid Edge UI.",
            "unsupported": True,
            "constraint_type": "angle",
            "component1": component1_index,
            "component2": component2_index,
            "angle": angle,
        }

    def add_planar_align_constraint(
        self,
        component1_index: int,
        component2_index: int,
        face1_index: int = 0,
        face2_index: int = 0,
    ) -> dict[str, Any]:
        """Same as add_align_constraint: a planar relation with the normals aligned."""
        return self.add_planar_relation(
            component1_index,
            component2_index,
            orientation="Align",
            face1_index=face1_index,
            face2_index=face2_index,
        )

    def add_axial_align_constraint(
        self,
        component1_index: int,
        component2_index: int,
        face1_index: int = 0,
        face2_index: int = 0,
    ) -> dict[str, Any]:
        """Make two components' cylindrical faces coaxial (see add_axial_relation)."""
        return self.add_axial_relation(
            component1_index,
            component2_index,
            orientation="Align",
            face1_index=face1_index,
            face2_index=face2_index,
        )

    def _validate_occurrences(
        self, doc: Any, occurrence1_index: int, occurrence2_index: int
    ) -> tuple[Any, Any, dict[str, Any] | None]:
        """Validate two occurrence indices and return occurrence objects.

        Returns (occ1, occ2, error_dict). If error_dict is not None, caller should return it.
        """
        err = self._require_assembly(doc)
        if err:
            return None, None, err

        occurrences = doc.Occurrences

        if occurrence1_index < 0 or occurrence1_index >= occurrences.Count:
            return (
                None,
                None,
                {
                    "error": f"Invalid occurrence1 index: "
                    f"{occurrence1_index}. Count: {occurrences.Count}"
                },
            )
        if occurrence2_index < 0 or occurrence2_index >= occurrences.Count:
            return (
                None,
                None,
                {
                    "error": f"Invalid occurrence2 index: "
                    f"{occurrence2_index}. Count: {occurrences.Count}"
                },
            )

        occ1 = occurrences.Item(occurrence1_index + 1)
        occ2 = occurrences.Item(occurrence2_index + 1)
        return occ1, occ2, None

    def _validate_relation_index(
        self, doc: Any, relation_index: int
    ) -> tuple[Any, dict[str, Any] | None]:
        """Validate a relation index and return the relation object.

        Returns (relation, error_dict). If error_dict is not None, caller should return it.
        """
        err = self._require_assembly(doc)
        if err:
            return None, err

        relations = doc.Relations3d

        if relation_index < 0 or relation_index >= relations.Count:
            return None, {
                "error": f"Invalid relation index: {relation_index}. Count: {relations.Count}"
            }

        return relations.Item(relation_index + 1), None

    def delete_relation(self, relation_index: int) -> dict[str, Any]:
        """
        Delete an assembly relation (constraint) by index.

        Args:
            relation_index: 0-based index of the relation

        Returns:
            Dict with status
        """
        try:
            _logger.info(f"Deleting relation at index {relation_index}")
            doc = self.doc_manager.get_active_document()

            err = self._require_assembly(doc)
            if err:
                return err

            relations = doc.Relations3d

            if relation_index < 0 or relation_index >= relations.Count:
                return {
                    "error": f"Invalid relation index: {relation_index}. Count: {relations.Count}"
                }

            rel = relations.Item(relation_index + 1)
            name = ""
            with contextlib.suppress(Exception):
                name = rel.Name

            rel.Delete()

            return {"status": "deleted", "relation_index": relation_index, "name": name}
        except Exception as e:
            _logger.error(f"Failed to delete relation: {e}")
            return error_result(e)

    def get_relation_info(self, relation_index: int) -> dict[str, Any]:
        """
        Get detailed information about a specific assembly relation.

        Args:
            relation_index: 0-based index of the relation

        Returns:
            Dict with relation type, status, offset, and connected elements
        """
        try:
            _logger.info(f"Getting relation info at index {relation_index}")
            doc = self.doc_manager.get_active_document()

            err = self._require_assembly(doc)
            if err:
                return err

            relations = doc.Relations3d

            if relation_index < 0 or relation_index >= relations.Count:
                return {
                    "error": f"Invalid relation index: {relation_index}. Count: {relations.Count}"
                }

            rel = relations.Item(relation_index + 1)

            type_names = {
                0: "Ground",
                1: "Axial",
                2: "Planar",
                3: "Connect",
                4: "Angle",
                5: "Tangent",
                6: "Cam",
                7: "Gear",
                8: "ParallelAxis",
                9: "Center",
            }

            info: dict[str, Any] = {"relation_index": relation_index}

            with contextlib.suppress(Exception):
                info["type"] = rel.Type
                info["type_name"] = type_names.get(rel.Type, f"Unknown({rel.Type})")
            with contextlib.suppress(Exception):
                info["status"] = rel.Status
            with contextlib.suppress(Exception):
                info["name"] = rel.Name
            with contextlib.suppress(Exception):
                info["suppressed"] = rel.Suppressed
            with contextlib.suppress(Exception):
                info["offset"] = rel.Offset
            with contextlib.suppress(Exception):
                info["normals_aligned"] = rel.NormalsAligned

            return info
        except Exception as e:
            _logger.error(f"Failed to get relation info: {e}")
            return error_result(e)

    @verifies_collection_growth("Relations3d")
    def add_planar_relation(
        self,
        occurrence1_index: int,
        occurrence2_index: int,
        offset: float = 0.0,
        orientation: str = "Align",
        face1_index: int = 0,
        face2_index: int = 0,
    ) -> dict[str, Any]:
        """Relate a planar face of one occurrence to a planar face of another.

        ``Relations3d.AddPlanar(Plane1, Plane2, NormalsAligned, ConstrainingPoint1,
        ConstrainingPoint2)`` on References from ``_face_reference``.
        NormalsAligned True mates the faces (they touch, normals opposed as
        Solid Edge counts it), False aligns them; ``orientation='Antialign'``
        mates, anything else aligns. Verified on Solid Edge 2026 between two
        placed boxes: Relations3d grows by one PlanarRelation3d.

        Args:
            occurrence1_index, occurrence2_index: 0-based occurrences.
            offset: not carried by AddPlanar; must be 0.
            orientation: 'Align' or 'Antialign' (mate).
            face1_index, face2_index: 0-based faces of each occurrence's body.
        """
        if offset:
            return {
                "error": (
                    "AddPlanar takes no offset; a planar relation is made at zero offset "
                    "(set one afterwards in the Solid Edge UI)."
                ),
                "offset": offset,
            }
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_assembly(doc)
            if err:
                return err
            first = self._face_reference(doc, occurrence1_index, face1_index)
            if isinstance(first, dict):
                return first
            second = self._face_reference(doc, occurrence2_index, face2_index)
            if isinstance(second, dict):
                return second
            (ref1, point1), (ref2, point2) = first, second
            mated = orientation == "Antialign"
            doc.Relations3d.AddPlanar(ref1, ref2, mated, point1, point2)
            return {
                "status": "created",
                "relation_type": "Planar",
                "orientation": orientation,
                "mated": mated,
                "occurrence1_index": occurrence1_index,
                "occurrence2_index": occurrence2_index,
                "face1_index": face1_index,
                "face2_index": face2_index,
                "relations": com_get(doc.Relations3d, "Count"),
            }
        except Exception as e:
            return error_result(e)

    @verifies_collection_growth("Relations3d")
    def add_axial_relation(
        self,
        occurrence1_index: int,
        occurrence2_index: int,
        orientation: str = "Align",
        face1_index: int = 0,
        face2_index: int = 0,
    ) -> dict[str, Any]:
        """Make two occurrences' cylindrical faces coaxial.

        ``Relations3d.AddAxial(Axis1, Axis2, NormalsAligned)`` on References from
        ``_face_reference`` (cylindrical faces). Verified on Solid Edge 2026
        between two placed cylinders: Relations3d grows by one AxialRelation3d.
        Handing it occurrences, as this did, answered E_FAIL.

        Args:
            occurrence1_index, occurrence2_index: 0-based occurrences.
            orientation: 'Align' or 'Antialign' (normals opposed).
            face1_index, face2_index: 0-based cylindrical faces of each body.
        """
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_assembly(doc)
            if err:
                return err
            first = self._face_reference(doc, occurrence1_index, face1_index, cylindrical=True)
            if isinstance(first, dict):
                return first
            second = self._face_reference(doc, occurrence2_index, face2_index, cylindrical=True)
            if isinstance(second, dict):
                return second
            aligned = orientation != "Antialign"
            doc.Relations3d.AddAxial(first[0], second[0], aligned)
            return {
                "status": "created",
                "relation_type": "Axial",
                "orientation": orientation,
                "occurrence1_index": occurrence1_index,
                "occurrence2_index": occurrence2_index,
                "face1_index": face1_index,
                "face2_index": face2_index,
                "relations": com_get(doc.Relations3d, "Count"),
            }
        except Exception as e:
            return error_result(e)

    def add_angular_relation(
        self,
        occurrence1_index: int,
        occurrence2_index: int,
        angle: float = 0.0,
    ) -> dict[str, Any]:
        """
        Add an angular relation between two assembly components.

        NOT AVAILABLE via COM automation. The real signature is
        ``Relations3d.AddAngular(Element1, Element2, ReverseElement1Direction,
        ReverseElement2Direction, MeasureElement1, MeasureElement2, Angle,
        MeasureToPositiveSide, MeasureFromPositiveSide, MeasureCCW)`` — ten
        arguments, four of them Faces or Edges on the two parts (the elements
        being constrained plus the two elements the angle is measured
        against). This server cannot select geometry, so it returns an
        ``unsupported`` error dict without touching COM. The signature is kept
        so tool dispatch keeps working.

        Args:
            occurrence1_index: 0-based index of first component
            occurrence2_index: 0-based index of second component
            angle: Angle in degrees

        Returns:
            Dict with an ``unsupported`` error
        """
        _logger.warning("add_angular_relation is not available via COM automation")
        del occurrence1_index, occurrence2_index, angle
        return {
            "error": (
                "Angular relations need four Face/Edge elements (the two constrained "
                "elements plus the two measurement elements Relations3d.AddAngular "
                "requires), which this server cannot select. Use the Solid Edge UI."
            ),
            "unsupported": True,
        }

    def add_point_relation(
        self,
        occurrence1_index: int,
        occurrence2_index: int,
    ) -> dict[str, Any]:
        """
        Add a point (connect) relation between two assembly components.

        NOT AVAILABLE via COM automation. The real signature is
        ``Relations3d.AddPoint(PointGeometry, PointKeyPoint, ConnectGeometry,
        [ConnectKeyPoint])``: the two geometry arguments are Faces, Edges or
        Vertices on the parts and the keypoint arguments say which point of
        that geometry to connect (``Relation3dGeometryConstants``). This
        server cannot select geometry, so it returns an ``unsupported`` error
        dict without touching COM. The signature is kept so tool dispatch
        keeps working.

        Args:
            occurrence1_index: 0-based index of first component
            occurrence2_index: 0-based index of second component

        Returns:
            Dict with an ``unsupported`` error
        """
        _logger.warning("add_point_relation is not available via COM automation")
        del occurrence1_index, occurrence2_index
        return {
            "error": (
                "Point (connect) relations need a geometry element and a keypoint on "
                "each part (Relations3d.AddPoint takes PointGeometry, PointKeyPoint, "
                "ConnectGeometry, ConnectKeyPoint), which this server cannot select. "
                "Use the Solid Edge UI."
            ),
            "unsupported": True,
        }

    def add_tangent_relation(
        self,
        occurrence1_index: int,
        occurrence2_index: int,
    ) -> dict[str, Any]:
        """
        Add a tangent relation between two assembly components.

        NOT AVAILABLE via COM automation. The real signature is
        ``Relations3d.AddTangent(Element1, Element2, ConstrainingPoint1,
        ConstrainingPoint2, Offset, IsHalfSpacePositive)``: the elements are
        the two Faces being made tangent and the constraining points are
        3-element arrays picked on them. This server cannot select a face, so
        it returns an ``unsupported`` error dict without touching COM. The
        signature is kept so tool dispatch keeps working.

        Args:
            occurrence1_index: 0-based index of first component
            occurrence2_index: 0-based index of second component

        Returns:
            Dict with an ``unsupported`` error
        """
        _logger.warning("add_tangent_relation is not available via COM automation")
        del occurrence1_index, occurrence2_index
        return {
            "error": (
                "Tangent relations need the two Faces being made tangent plus a "
                "constraining point on each (Relations3d.AddTangent takes Element1, "
                "Element2, ConstrainingPoint1, ConstrainingPoint2, Offset, "
                "IsHalfSpacePositive), which this server cannot select. "
                "Use the Solid Edge UI."
            ),
            "unsupported": True,
        }

    def add_gear_relation(
        self,
        occurrence1_index: int,
        occurrence2_index: int,
        ratio1: float = 1.0,
        ratio2: float = 1.0,
    ) -> dict[str, Any]:
        """
        Add a gear relation between two assembly components.

        NOT AVAILABLE via COM automation. The real signature is
        ``Relations3d.AddGear(Element1, Element2, GearType, RatioType,
        GearRatio1, GearRatio2, Flip)``: ``Element1``/``Element2`` are the
        rotational Faces or axes the gear couples, not the occurrences. This
        server cannot select geometry, so it returns an ``unsupported`` error
        dict without touching COM. The signature is kept so tool dispatch
        keeps working.

        Args:
            occurrence1_index: 0-based index of first component
            occurrence2_index: 0-based index of second component
            ratio1: Gear ratio value for first component (default 1.0)
            ratio2: Gear ratio value for second component (default 1.0)

        Returns:
            Dict with an ``unsupported`` error
        """
        _logger.warning("add_gear_relation is not available via COM automation")
        del occurrence1_index, occurrence2_index, ratio1, ratio2
        return {
            "error": (
                "Gear relations need the two rotational Face/axis elements being "
                "coupled (Relations3d.AddGear takes Element1, Element2, GearType, "
                "RatioType, GearRatio1, GearRatio2, Flip), which this server cannot "
                "select. Use the Solid Edge UI."
            ),
            "unsupported": True,
        }

    def get_relation_offset(self, relation_index: int) -> dict[str, Any]:
        """
        Get the offset value from a planar relation.

        Args:
            relation_index: 0-based index into Relations3d collection

        Returns:
            Dict with offset value (meters)
        """
        try:
            doc = self.doc_manager.get_active_document()
            rel, err = self._validate_relation_index(doc, relation_index)
            if err:
                return err

            offset = rel.Offset

            return {
                "relation_index": relation_index,
                "offset": offset,
            }
        except Exception as e:
            _logger.error(f"Failed to get relation offset: {e}")
            return error_result(e)

    def set_relation_offset(self, relation_index: int, offset: float) -> dict[str, Any]:
        """
        Set the offset value on a planar relation.

        Args:
            relation_index: 0-based index into Relations3d collection
            offset: New offset value in meters

        Returns:
            Dict with status
        """
        try:
            _logger.info(f"Setting relation offset: index={relation_index}, offset={offset}")
            doc = self.doc_manager.get_active_document()
            rel, err = self._validate_relation_index(doc, relation_index)
            if err:
                return err

            rel.Offset = offset

            return {
                "status": "updated",
                "relation_index": relation_index,
                "offset": offset,
            }
        except Exception as e:
            _logger.error(f"Failed to set relation offset: {e}")
            return error_result(e)

    def get_relation_angle(self, relation_index: int) -> dict[str, Any]:
        """
        Get the angle value from an angular relation.

        The COM API stores angles in radians; this returns degrees.

        Args:
            relation_index: 0-based index into Relations3d collection

        Returns:
            Dict with angle in degrees
        """
        try:
            doc = self.doc_manager.get_active_document()
            rel, err = self._validate_relation_index(doc, relation_index)
            if err:
                return err

            angle_rad = rel.Angle
            angle_deg = math.degrees(angle_rad)

            return {
                "relation_index": relation_index,
                "angle_degrees": angle_deg,
                "angle_radians": angle_rad,
            }
        except Exception as e:
            _logger.error(f"Failed to get relation angle: {e}")
            return error_result(e)

    def set_relation_angle(self, relation_index: int, angle: float) -> dict[str, Any]:
        """
        Set the angle value on an angular relation.

        Args:
            relation_index: 0-based index into Relations3d collection
            angle: New angle in degrees (converted to radians for COM)

        Returns:
            Dict with status
        """
        try:
            _logger.info(f"Setting relation angle: index={relation_index}, angle={angle}")
            doc = self.doc_manager.get_active_document()
            rel, err = self._validate_relation_index(doc, relation_index)
            if err:
                return err

            angle_rad = math.radians(angle)
            rel.Angle = angle_rad

            return {
                "status": "updated",
                "relation_index": relation_index,
                "angle_degrees": angle,
            }
        except Exception as e:
            _logger.error(f"Failed to set relation angle: {e}")
            return error_result(e)

    def get_normals_aligned(self, relation_index: int) -> dict[str, Any]:
        """
        Get the NormalsAligned property from a relation.

        Args:
            relation_index: 0-based index into Relations3d collection

        Returns:
            Dict with normals_aligned boolean
        """
        try:
            doc = self.doc_manager.get_active_document()
            rel, err = self._validate_relation_index(doc, relation_index)
            if err:
                return err

            aligned = rel.NormalsAligned

            return {
                "relation_index": relation_index,
                "normals_aligned": aligned,
            }
        except Exception as e:
            _logger.error(f"Failed to get normals aligned: {e}")
            return error_result(e)

    def set_normals_aligned(self, relation_index: int, aligned: bool) -> dict[str, Any]:
        """
        Set the NormalsAligned property on a relation.

        Args:
            relation_index: 0-based index into Relations3d collection
            aligned: True to align normals, False otherwise

        Returns:
            Dict with status
        """
        try:
            _logger.info(f"Setting normals aligned: index={relation_index}, aligned={aligned}")
            doc = self.doc_manager.get_active_document()
            rel, err = self._validate_relation_index(doc, relation_index)
            if err:
                return err

            rel.NormalsAligned = aligned

            return {
                "status": "updated",
                "relation_index": relation_index,
                "normals_aligned": aligned,
            }
        except Exception as e:
            _logger.error(f"Failed to set normals aligned: {e}")
            return error_result(e)

    def suppress_relation(self, relation_index: int) -> dict[str, Any]:
        """
        Suppress an assembly relation.

        Args:
            relation_index: 0-based index into Relations3d collection

        Returns:
            Dict with status
        """
        try:
            _logger.info(f"Suppressing relation at index {relation_index}")
            doc = self.doc_manager.get_active_document()
            rel, err = self._validate_relation_index(doc, relation_index)
            if err:
                return err

            rel.Suppressed = True

            return {
                "status": "suppressed",
                "relation_index": relation_index,
            }
        except Exception as e:
            _logger.error(f"Failed to suppress relation: {e}")
            return error_result(e)

    def unsuppress_relation(self, relation_index: int) -> dict[str, Any]:
        """
        Unsuppress an assembly relation.

        Args:
            relation_index: 0-based index into Relations3d collection

        Returns:
            Dict with status
        """
        try:
            _logger.info(f"Unsuppressing relation at index {relation_index}")
            doc = self.doc_manager.get_active_document()
            rel, err = self._validate_relation_index(doc, relation_index)
            if err:
                return err

            rel.Suppressed = False

            return {
                "status": "unsuppressed",
                "relation_index": relation_index,
            }
        except Exception as e:
            _logger.error(f"Failed to unsuppress relation: {e}")
            return error_result(e)

    def get_relation_geometry(self, relation_index: int) -> dict[str, Any]:
        """
        Get geometry info from a relation (connected occurrence references).

        Attempts to read OccurrencePart1, OccurrencePart2, and other geometry
        properties from the relation. Not all properties are available on all
        relation types.

        Args:
            relation_index: 0-based index into Relations3d collection

        Returns:
            Dict with available geometry/occurrence info
        """
        try:
            _logger.info(f"Getting relation geometry at index {relation_index}")
            doc = self.doc_manager.get_active_document()
            rel, err = self._validate_relation_index(doc, relation_index)
            if err:
                return err

            info: dict[str, Any] = {"relation_index": relation_index}

            with contextlib.suppress(Exception):
                info["type"] = rel.Type

            with contextlib.suppress(Exception):
                info["name"] = rel.Name

            with contextlib.suppress(Exception):
                occ1 = rel.Occurrence1
                info["occurrence1_name"] = com_get(occ1, "Name", str(occ1))

            with contextlib.suppress(Exception):
                occ2 = rel.Occurrence2
                info["occurrence2_name"] = com_get(occ2, "Name", str(occ2))

            with contextlib.suppress(Exception):
                info["offset"] = rel.Offset

            with contextlib.suppress(Exception):
                info["normals_aligned"] = rel.NormalsAligned

            with contextlib.suppress(Exception):
                info["suppressed"] = rel.Suppressed

            return info
        except Exception as e:
            _logger.error(f"Failed to get relation geometry: {e}")
            return error_result(e)

    def get_gear_ratio(self, relation_index: int) -> dict[str, Any]:
        """
        Get gear ratio values from a gear relation.

        Reads RatioValue1 and RatioValue2 from the relation.

        Args:
            relation_index: 0-based index into Relations3d collection

        Returns:
            Dict with ratio1 and ratio2 values
        """
        try:
            doc = self.doc_manager.get_active_document()
            rel, err = self._validate_relation_index(doc, relation_index)
            if err:
                return err

            info: dict[str, Any] = {"relation_index": relation_index}

            try:
                info["ratio1"] = rel.RatioValue1
            except Exception:
                info["ratio1"] = None
                info["ratio1_error"] = "RatioValue1 not available on this relation"

            try:
                info["ratio2"] = rel.RatioValue2
            except Exception:
                info["ratio2"] = None
                info["ratio2_error"] = "RatioValue2 not available on this relation"

            return info
        except Exception as e:
            _logger.error(f"Failed to get gear ratio: {e}")
            return error_result(e)
