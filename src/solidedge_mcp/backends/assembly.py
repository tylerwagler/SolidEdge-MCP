"""
Solid Edge Assembly Operations

Handles assembly creation and component management.
"""

import contextlib
import math
import os
import traceback
from typing import Any


class AssemblyManager:
    """Manages assembly operations"""

    def __init__(self, document_manager):
        self.doc_manager = document_manager

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
            return {"error": str(e), "traceback": traceback.format_exc()}

    # Alias for MCP tool compatibility
    place_component = add_component

    def list_components(self) -> dict[str, Any]:
        """
        List all components in the active assembly.

        Uses Occurrence.GetTransform() for position/rotation and
        OccurrenceFileName for file path.

        Returns:
            Dict with list of components and their properties
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            components = []

            for i in range(1, occurrences.Count + 1):
                occurrence = occurrences.Item(i)
                comp = {
                    "index": i - 1,
                    "name": occurrence.Name if hasattr(occurrence, "Name") else f"Component {i}",
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

            return {"components": components, "count": len(components)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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

    def get_component_info(self, component_index: int) -> dict[str, Any]:
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
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)

            info = {
                "index": component_index,
                "name": occurrence.Name if hasattr(occurrence, "Name") else "Unknown",
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
            with contextlib.suppress(Exception):
                info["visible"] = occurrence.Visible

            # Occurrence document info
            try:
                occ_doc = occurrence.OccurrenceDocument
                info["document_name"] = occ_doc.Name
            except Exception:
                pass

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def update_component_position(
        self, component_index: int, x: float, y: float, z: float
    ) -> dict[str, Any]:
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
                    "position": [x, y, z],
                }
            except Exception as e:
                return {
                    "error": f"Could not update position: {e}",
                    "note": "Position update may not be available for grounded components",
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

    def pattern_component(
        self, component_index: int, count: int, spacing: float, direction: str = "X"
    ) -> dict[str, Any]:
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
                placed.append(occ.Name if hasattr(occ, "Name") else f"copy_{i}")

            return {
                "status": "pattern_created",
                "source_component": component_index,
                "count": count,
                "spacing": spacing,
                "direction": direction,
                "placed_names": placed,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def suppress_component(self, component_index: int, suppress: bool = True) -> dict[str, Any]:
        """Suppress or unsuppress a component"""
        try:
            doc = self.doc_manager.get_active_document()
            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}"}

            occurrence = occurrences.Item(component_index + 1)

            if hasattr(occurrence, "Suppress") and suppress:
                occurrence.Suppress()
            elif hasattr(occurrence, "Unsuppress") and not suppress:
                occurrence.Unsuppress()
            else:
                return {"error": "Suppress/Unsuppress not available on this occurrence"}

            return {"status": "updated", "component": component_index, "suppressed": suppress}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_occurrence_bounding_box(self, component_index: int) -> dict[str, Any]:
        """
        Get the bounding box of a specific component (occurrence) in the assembly.

        Uses Occurrence.GetRangeBox() which returns min/max 3D points.

        Args:
            component_index: 0-based index of the component

        Returns:
            Dict with min/max coordinates
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)

            # GetRangeBox returns two arrays via out params
            import array

            min_point = array.array("d", [0.0, 0.0, 0.0])
            max_point = array.array("d", [0.0, 0.0, 0.0])

            occurrence.GetRangeBox(min_point, max_point)

            return {
                "component_index": component_index,
                "min": [min_point[0], min_point[1], min_point[2]],
                "max": [max_point[0], max_point[1], max_point[2]],
                "size": [
                    max_point[0] - min_point[0],
                    max_point[1] - min_point[1],
                    max_point[2] - min_point[2],
                ],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_bom(self) -> dict[str, Any]:
        """
        Get Bill of Materials from the active assembly.

        Recursively traverses all occurrences, deduplicates by file path,
        and returns a flat BOM with quantities.

        Returns:
            Dict with BOM items (file, name, quantity)
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            bom_counts: dict[str, dict[str, Any]] = {}

            for i in range(1, occurrences.Count + 1):
                occurrence = occurrences.Item(i)

                # Skip items excluded from BOM
                try:
                    if hasattr(occurrence, "IncludeInBom") and not occurrence.IncludeInBom:
                        continue
                except Exception:
                    pass

                # Skip pattern items (counted as part of pattern source)
                try:
                    if hasattr(occurrence, "IsPatternItem") and occurrence.IsPatternItem:
                        continue
                except Exception:
                    pass

                # Get file path as key
                try:
                    file_path = occurrence.OccurrenceFileName
                except Exception:
                    file_path = f"Unknown_{i}"

                name = (
                    occurrence.Name if hasattr(occurrence, "Name") else os.path.basename(file_path)
                )

                if file_path in bom_counts:
                    bom_counts[file_path]["quantity"] += 1
                else:
                    bom_counts[file_path] = {"name": name, "file_path": file_path, "quantity": 1}

            bom_items = list(bom_counts.values())

            return {
                "total_occurrences": occurrences.Count,
                "unique_parts": len(bom_items),
                "bom": bom_items,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_assembly_relations(self) -> dict[str, Any]:
        """
        Get all assembly relations (constraints) in the active assembly.

        Iterates the Relations3d collection to report constraint types,
        status, and connected occurrences.

        Returns:
            Dict with list of relations and their properties
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Relations3d"):
                return {"error": "Active document is not an assembly"}

            relations = doc.Relations3d
            relation_list = []

            # Relation type constants
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

            for i in range(1, relations.Count + 1):
                try:
                    rel = relations.Item(i)
                    rel_info = {"index": i - 1}

                    try:
                        rel_info["type"] = rel.Type
                        rel_info["type_name"] = type_names.get(rel.Type, f"Unknown({rel.Type})")
                    except Exception:
                        pass

                    with contextlib.suppress(Exception):
                        rel_info["status"] = rel.Status

                    with contextlib.suppress(Exception):
                        rel_info["suppressed"] = rel.Suppressed

                    with contextlib.suppress(Exception):
                        rel_info["name"] = rel.Name

                    relation_list.append(rel_info)
                except Exception:
                    relation_list.append({"index": i - 1})

            return {"relations": relation_list, "count": len(relation_list)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_document_tree(self) -> dict[str, Any]:
        """
        Get the hierarchical document tree of the active assembly.

        Recursively traverses occurrences and sub-occurrences to build
        a nested tree structure showing the full assembly hierarchy.

        Returns:
            Dict with nested tree of components
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            def traverse_occurrence(occ, depth=0):
                """Recursively build tree from an occurrence."""
                node = {}

                try:
                    node["name"] = occ.Name
                except Exception:
                    node["name"] = "Unknown"

                try:
                    node["file"] = occ.OccurrenceFileName
                except Exception:
                    node["file"] = "Unknown"

                with contextlib.suppress(Exception):
                    node["visible"] = occ.Visible

                with contextlib.suppress(Exception):
                    node["suppressed"] = occ.IsSuppressed if hasattr(occ, "IsSuppressed") else False

                # Recurse into sub-occurrences
                children = []
                try:
                    sub_occs = occ.SubOccurrences
                    if sub_occs and hasattr(sub_occs, "Count"):
                        for j in range(1, sub_occs.Count + 1):
                            try:
                                child = sub_occs.Item(j)
                                children.append(traverse_occurrence(child, depth + 1))
                            except Exception:
                                children.append({"name": f"SubOcc_{j}", "error": "could not read"})
                except Exception:
                    pass

                if children:
                    node["children"] = children

                return node

            occurrences = doc.Occurrences
            tree = []

            for i in range(1, occurrences.Count + 1):
                try:
                    occ = occurrences.Item(i)
                    tree.append(traverse_occurrence(occ))
                except Exception:
                    tree.append({"name": f"Occurrence_{i}", "error": "could not read"})

            return {
                "tree": tree,
                "top_level_count": len(tree),
                "document": doc.Name if hasattr(doc, "Name") else "Unknown",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_component_visibility(self, component_index: int, visible: bool) -> dict[str, Any]:
        """
        Set the visibility of a component in the assembly.

        Args:
            component_index: 0-based index of the component
            visible: True to show, False to hide

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)
            occurrence.Visible = visible

            return {"status": "updated", "component_index": component_index, "visible": visible}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def is_subassembly(self, component_index: int) -> dict[str, Any]:
        """
        Check if a component is a subassembly (vs a part).

        Args:
            component_index: 0-based index of the component

        Returns:
            Dict with is_subassembly boolean
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)

            result = {"component_index": component_index}

            try:
                result["is_subassembly"] = occurrence.Subassembly
            except Exception:
                # Fallback: check if it has SubOccurrences
                try:
                    sub_occs = occurrence.SubOccurrences
                    result["is_subassembly"] = (
                        sub_occs.Count > 0 if hasattr(sub_occs, "Count") else False
                    )
                except Exception:
                    result["is_subassembly"] = False

            with contextlib.suppress(Exception):
                result["name"] = occurrence.Name

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_component_display_name(self, component_index: int) -> dict[str, Any]:
        """
        Get the display name of a component.

        The display name is the user-visible label in the assembly tree,
        which may differ from the internal Name or OccurrenceFileName.

        Args:
            component_index: 0-based index of the component

        Returns:
            Dict with display_name and other name info
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)

            result = {"component_index": component_index}

            try:
                result["display_name"] = occurrence.DisplayName
            except Exception:
                result["display_name"] = None

            with contextlib.suppress(Exception):
                result["name"] = occurrence.Name

            with contextlib.suppress(Exception):
                result["file_name"] = occurrence.OccurrenceFileName

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_occurrence_document(self, component_index: int) -> dict[str, Any]:
        """
        Get document info for a component's source file.

        Args:
            component_index: 0-based index of the component

        Returns:
            Dict with document name, path, and type
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)

            result = {"component_index": component_index}

            try:
                occ_doc = occurrence.OccurrenceDocument
                with contextlib.suppress(Exception):
                    result["document_name"] = occ_doc.Name
                with contextlib.suppress(Exception):
                    result["full_name"] = occ_doc.FullName
                with contextlib.suppress(Exception):
                    result["type"] = occ_doc.Type
                with contextlib.suppress(Exception):
                    result["read_only"] = occ_doc.ReadOnly
            except Exception:
                result["error_note"] = "Could not access OccurrenceDocument"

            with contextlib.suppress(Exception):
                result["file_name"] = occurrence.OccurrenceFileName

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_sub_occurrences(self, component_index: int) -> dict[str, Any]:
        """
        Get sub-occurrences (children) of a component.

        For subassemblies, this returns the list of nested components.
        For parts, this returns an empty list.

        Args:
            component_index: 0-based index of the component

        Returns:
            Dict with list of sub-occurrence names and count
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)

            children = []
            try:
                sub_occs = occurrence.SubOccurrences
                if sub_occs and hasattr(sub_occs, "Count"):
                    for j in range(1, sub_occs.Count + 1):
                        try:
                            child = sub_occs.Item(j)
                            child_info = {"index": j - 1}
                            try:
                                child_info["name"] = child.Name
                            except Exception:
                                child_info["name"] = f"SubOcc_{j}"
                            with contextlib.suppress(Exception):
                                child_info["file"] = child.OccurrenceFileName
                            children.append(child_info)
                        except Exception:
                            children.append({"index": j - 1, "name": f"SubOcc_{j}"})
            except Exception:
                pass

            return {
                "component_index": component_index,
                "sub_occurrences": children,
                "count": len(children),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def delete_component(self, component_index: int) -> dict[str, Any]:
        """
        Delete/remove a component from the assembly.

        Args:
            component_index: 0-based index of the component to remove

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)
            name = (
                occurrence.Name if hasattr(occurrence, "Name") else f"Component_{component_index}"
            )
            occurrence.Delete()

            return {"status": "deleted", "component_index": component_index, "name": name}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def ground_component(self, component_index: int, ground: bool = True) -> dict[str, Any]:
        """
        Ground (fix in place) or unground a component in the assembly.

        Args:
            component_index: 0-based index of the component
            ground: True to ground, False to unground

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)

            if ground:
                # Add a ground constraint
                relations = doc.Relations3d
                relations.AddGround(occurrence)
                return {"status": "grounded", "component_index": component_index}
            else:
                # Find and delete ground relation for this occurrence
                relations = doc.Relations3d
                for i in range(relations.Count, 0, -1):
                    try:
                        rel = relations.Item(i)
                        # Ground relations have Type = 0
                        if hasattr(rel, "Type") and rel.Type == 0:
                            rel.Delete()
                            return {"status": "ungrounded", "component_index": component_index}
                    except Exception:
                        continue

                return {"error": "No ground relation found for this component"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def check_interference(self, component_index: int | None = None) -> dict[str, Any]:
        """
        Run interference check on the active assembly.

        If component_index is provided, checks that component against all others.
        If not provided, checks all components against each other.

        Args:
            component_index: Optional 0-based index of a specific component to check

        Returns:
            Dict with interference status and details
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if occurrences.Count < 2:
                return {
                    "status": "no_interference",
                    "message": "Need at least 2 components for interference check",
                }

            import ctypes

            # Build set1 - single component or all
            if component_index is not None:
                if component_index < 0 or component_index >= occurrences.Count:
                    return {"error": f"Invalid component index: {component_index}"}
                set1 = [occurrences.Item(component_index + 1)]
            else:
                set1 = [occurrences.Item(i) for i in range(1, occurrences.Count + 1)]

            # Call CheckInterference
            # seInterferenceComparisonSet1vsAllOther = 1
            comparison_method = 1

            # Prepare out parameters
            interference_status = ctypes.c_int(0)
            num_interferences = ctypes.c_int(0)

            try:
                doc.CheckInterference(
                    NumElementsSet1=len(set1),
                    Set1=set1,
                    Status=interference_status,
                    ComparisonMethod=comparison_method,
                    NumElementsSet2=0,
                    AddInterferenceAsOccurrence=False,
                    NumInterferences=num_interferences,
                )

                return {
                    "status": "checked",
                    "interference_found": interference_status.value != 0,
                    "num_interferences": num_interferences.value,
                    "component_checked": component_index,
                }
            except Exception as e:
                # CheckInterference has complex COM signature; report what we can
                return {
                    "error": f"Interference check failed: {e}",
                    "note": "CheckInterference COM signature "
                    "is complex. Use Solid Edge UI for "
                    "reliable results.",
                    "traceback": traceback.format_exc(),
                }

        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def replace_component(self, component_index: int, new_file_path: str) -> dict[str, Any]:
        """
        Replace a component in the assembly with a different part/assembly file.

        Preserves position and attempts to maintain assembly relations.

        Args:
            component_index: 0-based index of the component to replace
            new_file_path: Path to the replacement file (.par or .asm)

        Returns:
            Dict with replacement status
        """
        try:
            import os

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            if not os.path.exists(new_file_path):
                return {"error": f"File not found: {new_file_path}"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)
            old_name = occurrence.Name

            try:
                occurrence.Replace(new_file_path)
            except Exception:
                # Try alternative method
                occurrence.OccurrenceFileName = new_file_path

            return {
                "status": "replaced",
                "component_index": component_index,
                "old_name": old_name,
                "new_file": new_file_path,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_component_transform(self, component_index: int) -> dict[str, Any]:
        """
        Get the full transformation matrix of a component.

        Returns the 4x4 homogeneous transformation matrix and
        decomposed origin + rotation.

        Args:
            component_index: 0-based index of the component

        Returns:
            Dict with matrix, origin, and rotation
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)

            result = {
                "component_index": component_index,
                "name": occurrence.Name,
            }

            # Try GetTransform (origin + angles)
            try:
                transform = occurrence.GetTransform()
                result["origin"] = [transform[0], transform[1], transform[2]]
                result["rotation_angles"] = [transform[3], transform[4], transform[5]]
            except Exception:
                pass

            # Try GetMatrix (full 4x4)
            try:
                matrix = occurrence.GetMatrix()
                result["matrix"] = list(matrix)
            except Exception:
                pass

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_structured_bom(self) -> dict[str, Any]:
        """
        Get a hierarchical Bill of Materials with subassembly structure.

        Unlike get_bom() which returns a flat list, this preserves the
        parent-child hierarchy of subassemblies.

        Returns:
            Dict with structured BOM tree
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            def build_bom_item(occ, depth=0):
                item = {}
                try:
                    item["name"] = occ.Name
                except Exception:
                    item["name"] = "Unknown"

                try:
                    item["file"] = occ.OccurrenceFileName
                except Exception:
                    item["file"] = "Unknown"

                with contextlib.suppress(Exception):
                    item["visible"] = occ.Visible

                try:
                    item["suppressed"] = occ.IsSuppressed
                except Exception:
                    item["suppressed"] = False

                # Check for sub-occurrences (subassembly)
                children = []
                try:
                    sub_occs = occ.SubOccurrences
                    if sub_occs and hasattr(sub_occs, "Count") and sub_occs.Count > 0:
                        item["type"] = "assembly"
                        for j in range(1, sub_occs.Count + 1):
                            try:
                                children.append(build_bom_item(sub_occs.Item(j), depth + 1))
                            except Exception:
                                children.append({"name": f"SubItem_{j}", "error": "unreadable"})
                    else:
                        item["type"] = "part"
                except Exception:
                    item["type"] = "part"

                if children:
                    item["children"] = children
                    item["child_count"] = len(children)

                return item

            bom = []
            for i in range(1, occurrences.Count + 1):
                try:
                    bom.append(build_bom_item(occurrences.Item(i)))
                except Exception:
                    bom.append({"name": f"Component_{i}", "error": "unreadable"})

            return {
                "bom": bom,
                "top_level_count": len(bom),
                "document": doc.Name if hasattr(doc, "Name") else "Unknown",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_component_color(
        self, component_index: int, red: int, green: int, blue: int
    ) -> dict[str, Any]:
        """
        Set the color of a component in the assembly.

        Args:
            component_index: 0-based index of the component
            red: Red component (0-255)
            green: Green component (0-255)
            blue: Blue component (0-255)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)

            # OLE color: BGR format packed into integer
            ole_color = red | (green << 8) | (blue << 16)

            try:
                occurrence.SetColor(red, green, blue)
            except Exception:
                try:
                    occurrence.Color = ole_color
                except Exception:
                    # Try style-based approach
                    occurrence.UseOccurrenceColor = True
                    occurrence.OccurrenceColor = ole_color

            return {
                "status": "updated",
                "component_index": component_index,
                "color": [red, green, blue],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def occurrence_move(
        self, component_index: int, dx: float, dy: float, dz: float
    ) -> dict[str, Any]:
        """
        Move a component by a relative delta.

        Uses Occurrence.Move(DeltaX, DeltaY, DeltaZ) for relative translation.

        Args:
            component_index: 0-based index of the component
            dx: X translation in meters
            dy: Y translation in meters
            dz: Z translation in meters

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)
            occurrence.Move(dx, dy, dz)

            return {"status": "moved", "component_index": component_index, "delta": [dx, dy, dz]}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def occurrence_rotate(
        self,
        component_index: int,
        axis_x1: float,
        axis_y1: float,
        axis_z1: float,
        axis_x2: float,
        axis_y2: float,
        axis_z2: float,
        angle: float,
    ) -> dict[str, Any]:
        """
        Rotate a component around an axis.

        Uses Occurrence.Rotate(AxisX1, AxisY1, AxisZ1, AxisX2, AxisY2, AxisZ2, Angle).
        The axis is defined by two 3D points. Angle is in degrees (converted to radians internally).

        Args:
            component_index: 0-based index of the component
            axis_x1, axis_y1, axis_z1: First point of rotation axis (meters)
            axis_x2, axis_y2, axis_z2: Second point of rotation axis (meters)
            angle: Rotation angle in degrees

        Returns:
            Dict with status
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)
            angle_rad = math.radians(angle)
            occurrence.Rotate(axis_x1, axis_y1, axis_z1, axis_x2, axis_y2, axis_z2, angle_rad)

            return {
                "status": "rotated",
                "component_index": component_index,
                "axis": [[axis_x1, axis_y1, axis_z1], [axis_x2, axis_y2, axis_z2]],
                "angle_degrees": angle,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_component_transform(
        self,
        component_index: int,
        origin_x: float,
        origin_y: float,
        origin_z: float,
        angle_x: float,
        angle_y: float,
        angle_z: float,
    ) -> dict[str, Any]:
        """
        Set a component's full transform (position + rotation).

        Uses Occurrence.PutTransform(OriginX, OriginY, OriginZ, AngleX, AngleY, AngleZ).
        Angles are in degrees (converted to radians internally).

        Args:
            component_index: 0-based index of the component
            origin_x: X position in meters
            origin_y: Y position in meters
            origin_z: Z position in meters
            angle_x: Rotation around X axis in degrees
            angle_y: Rotation around Y axis in degrees
            angle_z: Rotation around Z axis in degrees

        Returns:
            Dict with status
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)
            ax_rad = math.radians(angle_x)
            ay_rad = math.radians(angle_y)
            az_rad = math.radians(angle_z)
            occurrence.PutTransform(origin_x, origin_y, origin_z, ax_rad, ay_rad, az_rad)

            return {
                "status": "updated",
                "component_index": component_index,
                "origin": [origin_x, origin_y, origin_z],
                "angles_degrees": [angle_x, angle_y, angle_z],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_component_origin(
        self, component_index: int, x: float, y: float, z: float
    ) -> dict[str, Any]:
        """
        Set a component's origin (position only, no rotation change).

        Uses Occurrence.PutOrigin(OriginX, OriginY, OriginZ).

        Args:
            component_index: 0-based index of the component
            x: X position in meters
            y: Y position in meters
            z: Z position in meters

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)
            occurrence.PutOrigin(x, y, z)

            return {"status": "updated", "component_index": component_index, "origin": [x, y, z]}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def mirror_component(self, component_index: int, plane_index: int) -> dict[str, Any]:
        """
        Mirror a component across a reference plane.

        Uses Occurrence.Mirror(pPlane).

        Args:
            component_index: 0-based index of the component
            plane_index: 1-based index of the reference plane

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            ref_planes = doc.RefPlanes
            if plane_index < 1 or plane_index > ref_planes.Count:
                return {"error": f"Invalid plane_index: {plane_index}. Count: {ref_planes.Count}"}

            occurrence = occurrences.Item(component_index + 1)
            plane = ref_planes.Item(plane_index)
            occurrence.Mirror(plane)

            return {
                "status": "mirrored",
                "component_index": component_index,
                "plane_index": plane_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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

    def get_occurrence_count(self) -> dict[str, Any]:
        """
        Get the count of top-level occurrences in the assembly.

        Returns:
            Dict with occurrence count
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            return {"count": doc.Occurrences.Count}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # ========================================================================
    # ASSEMBLY RELATION TOOLS (Batch 1)
    # ========================================================================

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

    # ========================================================================
    # BATCH 8: ASSEMBLY OCCURRENCES & PROPERTIES
    # ========================================================================

    def _validate_occurrence_index(
        self, doc, component_index: int
    ) -> tuple[Any, Any, dict[str, Any] | None]:
        """Validate a single occurrence index and return (occurrences, occurrence, error_dict).

        If error_dict is not None, caller should return it.
        """
        if not hasattr(doc, "Occurrences"):
            return None, None, {"error": "Active document is not an assembly"}

        occurrences = doc.Occurrences

        if component_index < 0 or component_index >= occurrences.Count:
            return (
                None,
                None,
                {
                    "error": f"Invalid component index: "
                    f"{component_index}. Count: {occurrences.Count}"
                },
            )

        occurrence = occurrences.Item(component_index + 1)
        return occurrences, occurrence, None

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
            return {"error": str(e), "traceback": traceback.format_exc()}

    def put_transform_euler(
        self,
        component_index: int,
        x: float,
        y: float,
        z: float,
        rx: float,
        ry: float,
        rz: float,
    ) -> dict[str, Any]:
        """
        Set a component's transform using Euler angles.

        Uses occurrence.PutTransform(x, y, z, rx_rad, ry_rad, rz_rad).
        Angles are in degrees (converted to radians internally).

        Args:
            component_index: 0-based index of the component
            x: X position in meters
            y: Y position in meters
            z: Z position in meters
            rx: Rotation around X axis in degrees
            ry: Rotation around Y axis in degrees
            rz: Rotation around Z axis in degrees

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            occurrences, occurrence, err = self._validate_occurrence_index(doc, component_index)
            if err:
                return err

            rx_rad = math.radians(rx)
            ry_rad = math.radians(ry)
            rz_rad = math.radians(rz)
            occurrence.PutTransform(x, y, z, rx_rad, ry_rad, rz_rad)

            return {
                "status": "updated",
                "component_index": component_index,
                "position": [x, y, z],
                "angles_degrees": [rx, ry, rz],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def put_origin(
        self,
        component_index: int,
        x: float,
        y: float,
        z: float,
    ) -> dict[str, Any]:
        """
        Set a component's origin (position only, no rotation change).

        Uses occurrence.PutOrigin(x, y, z).

        Args:
            component_index: 0-based index of the component
            x: X position in meters
            y: Y position in meters
            z: Z position in meters

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            occurrences, occurrence, err = self._validate_occurrence_index(doc, component_index)
            if err:
                return err

            occurrence.PutOrigin(x, y, z)

            return {
                "status": "updated",
                "component_index": component_index,
                "origin": [x, y, z],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def make_writable(self, component_index: int) -> dict[str, Any]:
        """
        Make a component writable (editable) in the assembly.

        Uses occurrence.MakeWritable() to allow editing of the component.

        Args:
            component_index: 0-based index of the component

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            occurrences, occurrence, err = self._validate_occurrence_index(doc, component_index)
            if err:
                return err

            occurrence.MakeWritable()

            return {
                "status": "writable",
                "component_index": component_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def swap_family_member(
        self,
        component_index: int,
        new_member_name: str,
    ) -> dict[str, Any]:
        """
        Swap a Family of Parts occurrence for a different family member.

        Uses occurrence.SwapFamilyMember(newMemberName).

        Args:
            component_index: 0-based index of the component
            new_member_name: Name of the new family member

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            occurrences, occurrence, err = self._validate_occurrence_index(doc, component_index)
            if err:
                return err

            occurrence.SwapFamilyMember(new_member_name)

            return {
                "status": "swapped",
                "component_index": component_index,
                "new_member_name": new_member_name,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_occurrence_bodies(self, component_index: int) -> dict[str, Any]:
        """
        Get body information from a specific component occurrence.

        Reads occurrence.Bodies property to enumerate solid bodies.

        Args:
            component_index: 0-based index of the component

        Returns:
            Dict with body count and body info
        """
        try:
            doc = self.doc_manager.get_active_document()
            occurrences, occurrence, err = self._validate_occurrence_index(doc, component_index)
            if err:
                return err

            bodies_info = []
            try:
                bodies = occurrence.Bodies
                body_count = bodies.Count if hasattr(bodies, "Count") else 0

                for i in range(1, body_count + 1):
                    body = bodies.Item(i)
                    body_info: dict[str, Any] = {"index": i - 1}

                    with contextlib.suppress(Exception):
                        body_info["name"] = body.Name

                    with contextlib.suppress(Exception):
                        body_info["volume"] = body.Volume

                    bodies_info.append(body_info)
            except Exception:
                body_count = 0

            return {
                "component_index": component_index,
                "body_count": len(bodies_info),
                "bodies": bodies_info,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_occurrence_style(self, component_index: int) -> dict[str, Any]:
        """
        Get the style (appearance) of a component occurrence.

        Reads occurrence.Style property.

        Args:
            component_index: 0-based index of the component

        Returns:
            Dict with style info
        """
        try:
            doc = self.doc_manager.get_active_document()
            occurrences, occurrence, err = self._validate_occurrence_index(doc, component_index)
            if err:
                return err

            result: dict[str, Any] = {"component_index": component_index}

            try:
                style = occurrence.Style
                result["style"] = str(style) if style is not None else None
            except Exception:
                result["style"] = None
                result["style_note"] = "Style property not available on this occurrence"

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def is_tube(self, component_index: int) -> dict[str, Any]:
        """
        Check if a component occurrence is a tube.

        Reads occurrence.IsTube property.

        Args:
            component_index: 0-based index of the component

        Returns:
            Dict with is_tube boolean
        """
        try:
            doc = self.doc_manager.get_active_document()
            occurrences, occurrence, err = self._validate_occurrence_index(doc, component_index)
            if err:
                return err

            result: dict[str, Any] = {"component_index": component_index}

            try:
                result["is_tube"] = bool(occurrence.IsTube)
            except Exception:
                result["is_tube"] = False
                result["is_tube_note"] = "IsTube property not available on this occurrence"

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_adjustable_part(self, component_index: int) -> dict[str, Any]:
        """
        Get adjustable part info from a component occurrence.

        Reads occurrence.GetAdjustablePart() to check if the component
        is adjustable and retrieve its adjustable part object info.

        Args:
            component_index: 0-based index of the component

        Returns:
            Dict with adjustable part info
        """
        try:
            doc = self.doc_manager.get_active_document()
            occurrences, occurrence, err = self._validate_occurrence_index(doc, component_index)
            if err:
                return err

            result: dict[str, Any] = {"component_index": component_index}

            try:
                adj_part = occurrence.GetAdjustablePart()
                result["is_adjustable"] = adj_part is not None
                if adj_part is not None:
                    with contextlib.suppress(Exception):
                        result["adjustable_name"] = adj_part.Name
            except Exception:
                result["is_adjustable"] = False
                result["adjustable_note"] = "GetAdjustablePart not available on this occurrence"

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_face_style(self, component_index: int) -> dict[str, Any]:
        """
        Get the face style of a component occurrence.

        Reads occurrence.GetFaceStyle2() for face style information.

        Args:
            component_index: 0-based index of the component

        Returns:
            Dict with face style info
        """
        try:
            doc = self.doc_manager.get_active_document()
            occurrences, occurrence, err = self._validate_occurrence_index(doc, component_index)
            if err:
                return err

            result: dict[str, Any] = {"component_index": component_index}

            try:
                face_style = occurrence.GetFaceStyle2()
                result["face_style"] = str(face_style) if face_style is not None else None
            except Exception:
                result["face_style"] = None
                result["face_style_note"] = "GetFaceStyle2 not available on this occurrence"

            return result
        except Exception as e:
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
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_occurrence(self, internal_id: int) -> dict[str, Any]:
        """
        Get an occurrence by its internal ID.

        Uses Occurrences.GetOccurrence(ID) to retrieve a specific occurrence
        by its internal ID (not by index). This is useful when you know the
        internal identifier assigned by Solid Edge.

        Args:
            internal_id: Internal ID of the occurrence (integer)

        Returns:
            Dict with occurrence info (name, file path, transform, etc.)
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            occurrence = occurrences.GetOccurrence(internal_id)

            if occurrence is None:
                return {"error": f"No occurrence found with ID: {internal_id}"}

            info: dict[str, Any] = {"internal_id": internal_id}

            # Name
            try:
                info["name"] = occurrence.Name
            except Exception:
                info["name"] = "Unknown"

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
                mat = occurrence.GetMatrix()
                info["matrix"] = list(mat)
            except Exception:
                pass

            # Visibility
            with contextlib.suppress(Exception):
                info["visible"] = occurrence.Visible

            # Occurrence document info
            try:
                occ_doc = occurrence.OccurrenceDocument
                info["document_name"] = occ_doc.Name
            except Exception:
                pass

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
