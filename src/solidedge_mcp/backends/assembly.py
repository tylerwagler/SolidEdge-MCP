"""
Solid Edge Assembly Operations

Handles assembly creation and component management.
"""

from typing import Dict, Any, Optional, List
import os
import traceback
from .constants import MateTypeConstants


class AssemblyManager:
    """Manages assembly operations"""

    def __init__(self, document_manager):
        self.doc_manager = document_manager

    def add_component(self, file_path: str, x: float = 0, y: float = 0, z: float = 0) -> Dict[str, Any]:
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
            if not os.path.exists(file_path):
                return {"error": f"File not found: {file_path}"}

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if x == 0 and y == 0 and z == 0:
                # Place at origin
                occurrence = occurrences.AddByFilename(file_path)
            else:
                # Place with transformation matrix (identity rotation + translation)
                matrix = [1.0, 0.0, 0.0, 0.0,
                          0.0, 1.0, 0.0, 0.0,
                          0.0, 0.0, 1.0, 0.0,
                          x,   y,   z,   1.0]
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
                "name": occurrence.Name if hasattr(occurrence, 'Name') else os.path.basename(file_path),
                "position": position,
                "index": occurrences.Count - 1
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # Alias for MCP tool compatibility
    place_component = add_component

    def list_components(self) -> Dict[str, Any]:
        """
        List all components in the active assembly.

        Uses Occurrence.GetTransform() for position/rotation and
        OccurrenceFileName for file path.

        Returns:
            Dict with list of components and their properties
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            components = []

            for i in range(1, occurrences.Count + 1):
                occurrence = occurrences.Item(i)
                comp = {
                    "index": i - 1,
                    "name": occurrence.Name if hasattr(occurrence, 'Name') else f"Component {i}",
                }

                # Get file path
                try:
                    comp["file_path"] = occurrence.OccurrenceFileName
                except Exception:
                    comp["file_path"] = "Unknown"

                # Get transform (originX, originY, originZ, angleX, angleY, angleZ)
                try:
                    transform = occurrence.GetTransform()
                    comp["position"] = [transform[0], transform[1], transform[2]]
                    comp["rotation"] = [transform[3], transform[4], transform[5]]
                except Exception:
                    comp["position"] = [0, 0, 0]
                    comp["rotation"] = [0, 0, 0]

                # Visibility/suppression
                try:
                    comp["visible"] = occurrence.Visible
                except Exception:
                    comp["visible"] = True

                components.append(comp)

            return {
                "components": components,
                "count": len(components)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_mate(self, mate_type: str, component1_index: int, component2_index: int) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Relations3d'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component1_index >= occurrences.Count or component2_index >= occurrences.Count:
                return {"error": "Invalid component index"}

            # Mate creation requires face/edge selection
            return {
                "error": "Mate creation requires face/edge selection which is not available via COM automation. Use Solid Edge UI to create mates.",
                "mate_type": mate_type,
                "component1": component1_index,
                "component2": component2_index
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_component_info(self, component_index: int) -> Dict[str, Any]:
        """
        Get detailed information about a specific component.

        Uses GetTransform for position/rotation and GetMatrix for the full 4x4 matrix.

        Args:
            component_index: 0-based index of the component

        Returns:
            Dict with component details
        """
        try:
            doc = self.doc_manager.get_active_document()
            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}. Count: {occurrences.Count}"}

            occurrence = occurrences.Item(component_index + 1)

            info = {
                "index": component_index,
                "name": occurrence.Name if hasattr(occurrence, 'Name') else "Unknown",
            }

            # File path
            try:
                info["file_path"] = occurrence.OccurrenceFileName
            except Exception:
                info["file_path"] = "Unknown"

            # Transform (position + rotation)
            try:
                transform = occurrence.GetTransform()
                info["position"] = [transform[0], transform[1], transform[2]]
                info["rotation_rad"] = [transform[3], transform[4], transform[5]]
            except Exception:
                pass

            # Full 4x4 matrix
            try:
                matrix = occurrence.GetMatrix()
                info["matrix"] = list(matrix)
            except Exception:
                pass

            # Visibility
            try:
                info["visible"] = occurrence.Visible
            except Exception:
                pass

            # Occurrence document info
            try:
                occ_doc = occurrence.OccurrenceDocument
                info["document_name"] = occ_doc.Name
            except Exception:
                pass

            return info
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def update_component_position(self, component_index: int, x: float, y: float, z: float) -> Dict[str, Any]:
        """
        Update a component's position using a transformation matrix.

        Args:
            component_index: 0-based index of the component
            x, y, z: New position coordinates (meters)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}"}

            occurrence = occurrences.Item(component_index + 1)

            # Get current matrix to preserve rotation
            try:
                current = list(occurrence.GetMatrix())
                # Update translation (indices 12, 13, 14 in row-major 4x4)
                current[12] = x
                current[13] = y
                current[14] = z
                occurrence.SetMatrix(current)
                return {
                    "status": "position_updated",
                    "component": component_index,
                    "position": [x, y, z]
                }
            except Exception as e:
                return {
                    "error": f"Could not update position: {e}",
                    "note": "Position update may not be available for grounded components"
                }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def add_align_constraint(self, component1_index: int, component2_index: int) -> Dict[str, Any]:
        """Add an align constraint between two components (requires UI for face selection)"""
        return {
            "error": "Constraint creation requires face/edge selection. Use Solid Edge UI.",
            "constraint_type": "align",
            "component1": component1_index,
            "component2": component2_index
        }

    def add_angle_constraint(self, component1_index: int, component2_index: int, angle: float) -> Dict[str, Any]:
        """Add an angle constraint between two components (requires UI for face selection)"""
        return {
            "error": "Constraint creation requires face/edge selection. Use Solid Edge UI.",
            "constraint_type": "angle",
            "component1": component1_index,
            "component2": component2_index,
            "angle": angle
        }

    def add_planar_align_constraint(self, component1_index: int, component2_index: int) -> Dict[str, Any]:
        """Add a planar align constraint (requires UI for face selection)"""
        return {
            "error": "Constraint creation requires face/edge selection. Use Solid Edge UI.",
            "constraint_type": "planar_align",
            "component1": component1_index,
            "component2": component2_index
        }

    def add_axial_align_constraint(self, component1_index: int, component2_index: int) -> Dict[str, Any]:
        """Add an axial align constraint (requires UI for face selection)"""
        return {
            "error": "Constraint creation requires face/edge selection. Use Solid Edge UI.",
            "constraint_type": "axial_align",
            "component1": component1_index,
            "component2": component2_index
        }

    def pattern_component(self, component_index: int, count: int, spacing: float, direction: str = "X") -> Dict[str, Any]:
        """
        Create a linear pattern of a component by placing copies with offset.

        Args:
            component_index: 0-based index of the source component
            count: Number of total instances (including original)
            spacing: Distance between instances (meters)
            direction: Pattern direction - 'X', 'Y', or 'Z'

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}"}

            source = occurrences.Item(component_index + 1)

            # Get the source file path
            try:
                file_path = source.OccurrenceFileName
            except Exception:
                return {"error": "Cannot determine source component file path"}

            # Get source position
            try:
                base_matrix = list(source.GetMatrix())
            except Exception:
                base_matrix = [1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1]

            dir_map = {"X": 12, "Y": 13, "Z": 14}
            dir_idx = dir_map.get(direction, 12)

            placed = []
            for i in range(1, count):
                matrix = list(base_matrix)
                matrix[dir_idx] = base_matrix[dir_idx] + (spacing * i)
                occ = occurrences.AddWithMatrix(file_path, matrix)
                placed.append(occ.Name if hasattr(occ, 'Name') else f"copy_{i}")

            return {
                "status": "pattern_created",
                "source_component": component_index,
                "count": count,
                "spacing": spacing,
                "direction": direction,
                "placed_names": placed
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def suppress_component(self, component_index: int, suppress: bool = True) -> Dict[str, Any]:
        """Suppress or unsuppress a component"""
        try:
            doc = self.doc_manager.get_active_document()
            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}"}

            occurrence = occurrences.Item(component_index + 1)

            if hasattr(occurrence, 'Suppress') and suppress:
                occurrence.Suppress()
            elif hasattr(occurrence, 'Unsuppress') and not suppress:
                occurrence.Unsuppress()
            else:
                return {"error": "Suppress/Unsuppress not available on this occurrence"}

            return {
                "status": "updated",
                "component": component_index,
                "suppressed": suppress
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
