"""Assembly relation (constraint) operations."""

import contextlib
import math
import traceback
from typing import Any


class RelationsMixin:
    """Mixin providing assembly relation/constraint methods."""

    def create_mate(
        self, mate_type: str, component1_index: int, component2_index: int
    ) -> dict[str, Any]:
        """
        Create a mate/assembly relationship between components.

        Note: Actual mate creation requires face/edge selection which cannot
        be done programmatically without specific geometry references.

        Args:
            mate_type: Type of mate - 'Planar', 'Axial', 'Insert', 'Match', 'Parallel', 'Angle'
            component1_index: Index of first component
            component2_index: Index of second component

        Returns:
            Dict with status and mate info
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Relations3d"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component1_index >= occurrences.Count or component2_index >= occurrences.Count:
                return {"error": "Invalid component index"}

            # Mate creation requires face/edge selection
            return {
                "error": "Mate creation requires face/edge "
                "selection which is not available via "
                "COM automation. Use Solid Edge UI to "
                "create mates.",
                "mate_type": mate_type,
                "component1": component1_index,
                "component2": component2_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_align_constraint(self, component1_index: int, component2_index: int) -> dict[str, Any]:
        """Add an align constraint between two components (requires UI for face selection)"""
        return {
            "error": "Constraint creation requires face/edge selection. Use Solid Edge UI.",
            "constraint_type": "align",
            "component1": component1_index,
            "component2": component2_index,
        }

    def add_angle_constraint(
        self, component1_index: int, component2_index: int, angle: float
    ) -> dict[str, Any]:
        """Add an angle constraint between two components (requires UI for face selection)"""
        return {
            "error": "Constraint creation requires face/edge selection. Use Solid Edge UI.",
            "constraint_type": "angle",
            "component1": component1_index,
            "component2": component2_index,
            "angle": angle,
        }

    def add_planar_align_constraint(
        self, component1_index: int, component2_index: int
    ) -> dict[str, Any]:
        """Add a planar align constraint (requires UI for face selection)"""
        return {
            "error": "Constraint creation requires face/edge selection. Use Solid Edge UI.",
            "constraint_type": "planar_align",
            "component1": component1_index,
            "component2": component2_index,
        }

    def add_axial_align_constraint(
        self, component1_index: int, component2_index: int
    ) -> dict[str, Any]:
        """Add an axial align constraint (requires UI for face selection)"""
        return {
            "error": "Constraint creation requires face/edge selection. Use Solid Edge UI.",
            "constraint_type": "axial_align",
            "component1": component1_index,
            "component2": component2_index,
        }

    def _validate_occurrences(
        self, doc, occurrence1_index: int, occurrence2_index: int
    ) -> tuple[Any, Any, dict[str, Any] | None]:
        """Validate two occurrence indices and return occurrence objects.

        Returns (occ1, occ2, error_dict). If error_dict is not None, caller should return it.
        """
        if not hasattr(doc, "Relations3d"):
            return None, None, {"error": "Active document is not an assembly"}

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
        self, doc, relation_index: int
    ) -> tuple[Any, dict[str, Any] | None]:
        """Validate a relation index and return the relation object.

        Returns (relation, error_dict). If error_dict is not None, caller should return it.
        """
        if not hasattr(doc, "Relations3d"):
            return None, {"error": "Active document is not an assembly"}

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
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Relations3d"):
                return {"error": "Active document is not an assembly"}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_relation_info(self, relation_index: int) -> dict[str, Any]:
        """
        Get detailed information about a specific assembly relation.

        Args:
            relation_index: 0-based index of the relation

        Returns:
            Dict with relation type, status, offset, and connected elements
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Relations3d"):
                return {"error": "Active document is not an assembly"}

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

            info = {"relation_index": relation_index}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_planar_relation(
        self,
        occurrence1_index: int,
        occurrence2_index: int,
        offset: float = 0.0,
        orientation: str = "Align",
    ) -> dict[str, Any]:
        """
        Add a planar relation between two assembly components.

        Uses Relations3d.AddPlanar(Occurrence1, Occurrence2, Offset, OrientationType).

        Args:
            occurrence1_index: 0-based index of first component
            occurrence2_index: 0-based index of second component
            offset: Offset distance in meters (default 0.0)
            orientation: "Align" (1), "Antialign" (2), or "NotSpecified" (0)

        Returns:
            Dict with status and relation info
        """
        try:
            doc = self.doc_manager.get_active_document()
            occ1, occ2, err = self._validate_occurrences(doc, occurrence1_index, occurrence2_index)
            if err:
                return err

            orient_map = {"Align": 1, "Antialign": 2, "NotSpecified": 0}
            orient_val = orient_map.get(orientation, 0)

            relations = doc.Relations3d
            relations.AddPlanar(occ1, occ2, offset, orient_val)

            return {
                "status": "created",
                "relation_type": "Planar",
                "occurrence1_index": occurrence1_index,
                "occurrence2_index": occurrence2_index,
                "offset": offset,
                "orientation": orientation,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_axial_relation(
        self,
        occurrence1_index: int,
        occurrence2_index: int,
        orientation: str = "Align",
    ) -> dict[str, Any]:
        """
        Add an axial relation between two assembly components.

        Uses Relations3d.AddAxial(Occurrence1, Occurrence2, OrientationType).

        Args:
            occurrence1_index: 0-based index of first component
            occurrence2_index: 0-based index of second component
            orientation: "Align" (1), "Antialign" (2), or "NotSpecified" (0)

        Returns:
            Dict with status and relation info
        """
        try:
            doc = self.doc_manager.get_active_document()
            occ1, occ2, err = self._validate_occurrences(doc, occurrence1_index, occurrence2_index)
            if err:
                return err

            orient_map = {"Align": 1, "Antialign": 2, "NotSpecified": 0}
            orient_val = orient_map.get(orientation, 0)

            relations = doc.Relations3d
            relations.AddAxial(occ1, occ2, orient_val)

            return {
                "status": "created",
                "relation_type": "Axial",
                "occurrence1_index": occurrence1_index,
                "occurrence2_index": occurrence2_index,
                "orientation": orientation,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_angular_relation(
        self,
        occurrence1_index: int,
        occurrence2_index: int,
        angle: float = 0.0,
    ) -> dict[str, Any]:
        """
        Add an angular relation between two assembly components.

        Uses Relations3d.AddAngular(Occurrence1, Occurrence2, AngleInRadians).

        Args:
            occurrence1_index: 0-based index of first component
            occurrence2_index: 0-based index of second component
            angle: Angle in degrees (converted to radians for COM)

        Returns:
            Dict with status and relation info
        """
        try:
            doc = self.doc_manager.get_active_document()
            occ1, occ2, err = self._validate_occurrences(doc, occurrence1_index, occurrence2_index)
            if err:
                return err

            angle_rad = math.radians(angle)

            relations = doc.Relations3d
            relations.AddAngular(occ1, occ2, angle_rad)

            return {
                "status": "created",
                "relation_type": "Angular",
                "occurrence1_index": occurrence1_index,
                "occurrence2_index": occurrence2_index,
                "angle_degrees": angle,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_point_relation(
        self,
        occurrence1_index: int,
        occurrence2_index: int,
    ) -> dict[str, Any]:
        """
        Add a point (connect) relation between two assembly components.

        Uses Relations3d.AddPoint(Occurrence1, Occurrence2).

        Args:
            occurrence1_index: 0-based index of first component
            occurrence2_index: 0-based index of second component

        Returns:
            Dict with status and relation info
        """
        try:
            doc = self.doc_manager.get_active_document()
            occ1, occ2, err = self._validate_occurrences(doc, occurrence1_index, occurrence2_index)
            if err:
                return err

            relations = doc.Relations3d
            relations.AddPoint(occ1, occ2)

            return {
                "status": "created",
                "relation_type": "Point",
                "occurrence1_index": occurrence1_index,
                "occurrence2_index": occurrence2_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_tangent_relation(
        self,
        occurrence1_index: int,
        occurrence2_index: int,
    ) -> dict[str, Any]:
        """
        Add a tangent relation between two assembly components.

        Uses Relations3d.AddTangent(Occurrence1, Occurrence2).

        Args:
            occurrence1_index: 0-based index of first component
            occurrence2_index: 0-based index of second component

        Returns:
            Dict with status and relation info
        """
        try:
            doc = self.doc_manager.get_active_document()
            occ1, occ2, err = self._validate_occurrences(doc, occurrence1_index, occurrence2_index)
            if err:
                return err

            relations = doc.Relations3d
            relations.AddTangent(occ1, occ2)

            return {
                "status": "created",
                "relation_type": "Tangent",
                "occurrence1_index": occurrence1_index,
                "occurrence2_index": occurrence2_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_gear_relation(
        self,
        occurrence1_index: int,
        occurrence2_index: int,
        ratio1: float = 1.0,
        ratio2: float = 1.0,
    ) -> dict[str, Any]:
        """
        Add a gear relation between two assembly components.

        Uses Relations3d.AddGear(Occurrence1, Occurrence2, Ratio1, Ratio2).

        Args:
            occurrence1_index: 0-based index of first component
            occurrence2_index: 0-based index of second component
            ratio1: Gear ratio value for first component (default 1.0)
            ratio2: Gear ratio value for second component (default 1.0)

        Returns:
            Dict with status and relation info
        """
        try:
            doc = self.doc_manager.get_active_document()
            occ1, occ2, err = self._validate_occurrences(doc, occurrence1_index, occurrence2_index)
            if err:
                return err

            relations = doc.Relations3d
            relations.AddGear(occ1, occ2, ratio1, ratio2)

            return {
                "status": "created",
                "relation_type": "Gear",
                "occurrence1_index": occurrence1_index,
                "occurrence2_index": occurrence2_index,
                "ratio1": ratio1,
                "ratio2": ratio2,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

    def suppress_relation(self, relation_index: int) -> dict[str, Any]:
        """
        Suppress an assembly relation.

        Args:
            relation_index: 0-based index into Relations3d collection

        Returns:
            Dict with status
        """
        try:
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
            return {"error": str(e), "traceback": traceback.format_exc()}

    def unsuppress_relation(self, relation_index: int) -> dict[str, Any]:
        """
        Unsuppress an assembly relation.

        Args:
            relation_index: 0-based index into Relations3d collection

        Returns:
            Dict with status
        """
        try:
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
                occ1 = rel.OccurrencePart1
                info["occurrence1_name"] = occ1.Name if hasattr(occ1, "Name") else str(occ1)

            with contextlib.suppress(Exception):
                occ2 = rel.OccurrencePart2
                info["occurrence2_name"] = occ2.Name if hasattr(occ2, "Name") else str(occ2)

            with contextlib.suppress(Exception):
                info["offset"] = rel.Offset

            with contextlib.suppress(Exception):
                info["normals_aligned"] = rel.NormalsAligned

            with contextlib.suppress(Exception):
                info["suppressed"] = rel.Suppressed

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}
