"""Placement operations for assembly components."""

import math
import os
import traceback
from typing import Any

from ..logging import get_logger

_logger = get_logger(__name__)


class PlacementMixin:
    """Mixin providing component placement methods."""

    def add_component(
        self, file_path: str, x: float = 0, y: float = 0, z: float = 0
    ) -> dict[str, Any]:
        """
        Add a component (part) to the active assembly.

        Uses Occurrences.AddByFilename for origin placement, or
        Occurrences.AddWithMatrix for positioned placement.

        Args:
            file_path: Path to the part file (.par or .asm)
            x, y, z: Position coordinates in meters

        Returns:
            Dict with status and component info
        """
        try:
            _logger.info(f"Adding component: {file_path}")
            if not os.path.exists(file_path):
                return {"error": f"File not found: {file_path}"}

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if x == 0 and y == 0 and z == 0:
                # Place at origin
                occurrence = occurrences.AddByFilename(file_path)
            else:
                # Place with transformation matrix (identity rotation + translation)
                matrix = [1.0, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0, x, y, z, 1.0]
                occurrence = occurrences.AddWithMatrix(file_path, matrix)

            # Get actual position from transform
            try:
                transform = occurrence.GetTransform()
                position = [transform[0], transform[1], transform[2]]
            except Exception:
                position = [x, y, z]

            return {
                "status": "added",
                "file_path": file_path,
                "name": (
                    occurrence.Name if hasattr(occurrence, "Name") else os.path.basename(file_path)
                ),
                "position": position,
                "index": occurrences.Count - 1,
            }
        except Exception as e:
            _logger.error(f"Failed to add component: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    # Alias for MCP tool compatibility
    place_component = add_component

    def add_component_with_transform(
        self,
        file_path: str,
        origin_x: float = 0,
        origin_y: float = 0,
        origin_z: float = 0,
        angle_x: float = 0,
        angle_y: float = 0,
        angle_z: float = 0,
    ) -> dict[str, Any]:
        """
        Add a component with position and Euler rotation angles.

        Uses Occurrences.AddWithTransform(filename, ox, oy, oz, ax, ay, az)
        where angles are in radians.

        Args:
            file_path: Path to the part or assembly file
            origin_x, origin_y, origin_z: Position in meters
            angle_x, angle_y, angle_z: Rotation angles in degrees

        Returns:
            Dict with status and component info
        """
        try:
            _logger.info(f"Adding component with transform: {file_path}")
            if not os.path.exists(file_path):
                return {"error": f"File not found: {file_path}"}

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            ax_rad = math.radians(angle_x)
            ay_rad = math.radians(angle_y)
            az_rad = math.radians(angle_z)

            occurrence = occurrences.AddWithTransform(
                file_path, origin_x, origin_y, origin_z, ax_rad, ay_rad, az_rad
            )

            return {
                "status": "added",
                "file_path": file_path,
                "name": occurrence.Name
                if hasattr(occurrence, "Name")
                else os.path.basename(file_path),
                "origin": [origin_x, origin_y, origin_z],
                "angles_degrees": [angle_x, angle_y, angle_z],
                "index": occurrences.Count - 1,
            }
        except Exception as e:
            _logger.error(f"Failed to add component with transform: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_family_member(
        self,
        file_path: str,
        family_member_name: str,
        x: float = 0,
        y: float = 0,
        z: float = 0,
    ) -> dict[str, Any]:
        """
        Add a Family of Parts member to the assembly.

        Uses Occurrences.AddFamilyByFilename to place a specific family member.

        Args:
            file_path: Path to the Family of Parts file (.par)
            family_member_name: Name of the family member to place
            x: X position in meters (unused, placement is at origin)
            y: Y position in meters (unused, placement is at origin)
            z: Z position in meters (unused, placement is at origin)

        Returns:
            Dict with status and component info
        """
        try:
            _logger.info(f"Adding family member: {family_member_name} from {file_path}")
            if not os.path.exists(file_path):
                return {"error": f"File not found: {file_path}"}

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            occ = occurrences.AddFamilyByFilename(file_path, family_member_name)

            return {
                "status": "added",
                "file_path": file_path,
                "family_member": family_member_name,
                "name": occ.Name if hasattr(occ, "Name") else "Unknown",
                "index": occurrences.Count - 1,
            }
        except Exception as e:
            _logger.error(f"Failed to add family member: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_family_with_transform(
        self,
        file_path: str,
        family_member_name: str,
        origin_x: float = 0,
        origin_y: float = 0,
        origin_z: float = 0,
        angle_x: float = 0,
        angle_y: float = 0,
        angle_z: float = 0,
    ) -> dict[str, Any]:
        """
        Add a Family of Parts member with position and rotation.

        Places the family member, then applies a transform via PutTransform.
        Angles are in degrees (converted to radians internally).

        Args:
            file_path: Path to the Family of Parts file (.par)
            family_member_name: Name of the family member to place
            origin_x, origin_y, origin_z: Position in meters
            angle_x, angle_y, angle_z: Rotation angles in degrees

        Returns:
            Dict with status and component info
        """
        try:
            _logger.info(
                "Adding family member with transform: %s from %s",
                family_member_name, file_path,
            )
            if not os.path.exists(file_path):
                return {"error": f"File not found: {file_path}"}

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            occ = occurrences.AddFamilyByFilename(file_path, family_member_name)

            # Apply transform
            ax_rad = math.radians(angle_x)
            ay_rad = math.radians(angle_y)
            az_rad = math.radians(angle_z)
            occ.PutTransform(origin_x, origin_y, origin_z, ax_rad, ay_rad, az_rad)

            return {
                "status": "added",
                "file_path": file_path,
                "family_member": family_member_name,
                "name": occ.Name if hasattr(occ, "Name") else "Unknown",
                "origin": [origin_x, origin_y, origin_z],
                "angles_degrees": [angle_x, angle_y, angle_z],
                "index": occurrences.Count - 1,
            }
        except Exception as e:
            _logger.error(f"Failed to add family member with transform: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_family_with_matrix(
        self,
        family_file_path: str,
        member_name: str,
        matrix: list[float],
    ) -> dict[str, Any]:
        """
        Add a Family of Parts member with a 4x4 transformation matrix.

        Uses Occurrences.AddFamilyWithMatrix(OccurrenceFileName, Matrix, MemberName)
        to place a specific family member at the position/orientation defined by the matrix.

        Args:
            family_file_path: Path to the Family of Parts file (.par)
            member_name: Name of the family member to place
            matrix: 16-element list of floats representing a 4x4 transformation matrix

        Returns:
            Dict with status and component info
        """
        try:
            _logger.info(f"Adding family member with matrix: {member_name} from {family_file_path}")
            if not os.path.exists(family_file_path):
                return {"error": f"File not found: {family_file_path}"}

            if len(matrix) != 16:
                return {"error": f"Matrix must have exactly 16 elements, got {len(matrix)}"}

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            occ = occurrences.AddFamilyWithMatrix(family_file_path, matrix, member_name)

            # Extract position from transform
            position = [matrix[12], matrix[13], matrix[14]]

            return {
                "status": "added",
                "file_path": family_file_path,
                "family_member": member_name,
                "name": occ.Name if hasattr(occ, "Name") else "Unknown",
                "position": position,
                "matrix": matrix,
                "index": occurrences.Count - 1,
            }
        except Exception as e:
            _logger.error(f"Failed to add family member with matrix: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_by_template(
        self,
        file_path: str,
        template_name: str,
    ) -> dict[str, Any]:
        """
        Add a component to the assembly using a template.

        Uses Occurrences.AddByTemplate(filename, templateName).

        Args:
            file_path: Path to the part or assembly file
            template_name: Name of the template to use

        Returns:
            Dict with status and component info
        """
        try:
            _logger.info(f"Adding component by template: {file_path}, template={template_name}")
            if not os.path.exists(file_path):
                return {"error": f"File not found: {file_path}"}

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            occ = occurrences.AddByTemplate(file_path, template_name)

            return {
                "status": "added",
                "file_path": file_path,
                "template_name": template_name,
                "name": occ.Name if hasattr(occ, "Name") else "Unknown",
                "index": occurrences.Count - 1,
            }
        except Exception as e:
            _logger.error(f"Failed to add component by template: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_adjustable_part(
        self,
        file_path: str,
        x: float = 0,
        y: float = 0,
        z: float = 0,
    ) -> dict[str, Any]:
        """
        Add a part as an adjustable part to the assembly.

        Uses Occurrences.AddAsAdjustablePart(filename).

        Args:
            file_path: Path to the part file (.par)
            x: X position in meters (unused, placement is at origin)
            y: Y position in meters (unused, placement is at origin)
            z: Z position in meters (unused, placement is at origin)

        Returns:
            Dict with status and component info
        """
        try:
            _logger.info(f"Adding adjustable part: {file_path}")
            if not os.path.exists(file_path):
                return {"error": f"File not found: {file_path}"}

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            occ = occurrences.AddAsAdjustablePart(file_path)

            return {
                "status": "added",
                "file_path": file_path,
                "adjustable": True,
                "name": occ.Name if hasattr(occ, "Name") else "Unknown",
                "index": occurrences.Count - 1,
            }
        except Exception as e:
            _logger.error(f"Failed to add adjustable part: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def reorder_occurrence(
        self,
        component_index: int,
        target_index: int,
    ) -> dict[str, Any]:
        """
        Reorder an occurrence in the assembly tree.

        Uses Occurrences.ReorderOccurrence(occurrence, targetIndex).
        Both indices are 0-based (converted to 1-based for COM).

        Args:
            component_index: 0-based index of the component to move
            target_index: 0-based target position

        Returns:
            Dict with status
        """
        try:
            _logger.info(f"Reordering occurrence {component_index} to {target_index}")
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. Count: {occurrences.Count}"
                }

            if target_index < 0 or target_index >= occurrences.Count:
                return {
                    "error": f"Invalid target index: {target_index}. Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)
            occurrences.ReorderOccurrence(occurrence, target_index + 1)

            return {
                "status": "reordered",
                "component_index": component_index,
                "target_index": target_index,
            }
        except Exception as e:
            _logger.error(f"Failed to reorder occurrence: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}
