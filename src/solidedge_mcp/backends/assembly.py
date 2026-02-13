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

    def get_occurrence_bounding_box(self, component_index: int) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}. Count: {occurrences.Count}"}

            occurrence = occurrences.Item(component_index + 1)

            # GetRangeBox returns two arrays via out params
            import array
            min_point = array.array('d', [0.0, 0.0, 0.0])
            max_point = array.array('d', [0.0, 0.0, 0.0])

            occurrence.GetRangeBox(min_point, max_point)

            return {
                "component_index": component_index,
                "min": [min_point[0], min_point[1], min_point[2]],
                "max": [max_point[0], max_point[1], max_point[2]],
                "size": [
                    max_point[0] - min_point[0],
                    max_point[1] - min_point[1],
                    max_point[2] - min_point[2]
                ]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_bom(self) -> Dict[str, Any]:
        """
        Get Bill of Materials from the active assembly.

        Recursively traverses all occurrences, deduplicates by file path,
        and returns a flat BOM with quantities.

        Returns:
            Dict with BOM items (file, name, quantity)
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            bom_counts: Dict[str, Dict[str, Any]] = {}

            for i in range(1, occurrences.Count + 1):
                occurrence = occurrences.Item(i)

                # Skip items excluded from BOM
                try:
                    if hasattr(occurrence, 'IncludeInBom') and not occurrence.IncludeInBom:
                        continue
                except Exception:
                    pass

                # Skip pattern items (counted as part of pattern source)
                try:
                    if hasattr(occurrence, 'IsPatternItem') and occurrence.IsPatternItem:
                        continue
                except Exception:
                    pass

                # Get file path as key
                try:
                    file_path = occurrence.OccurrenceFileName
                except Exception:
                    file_path = f"Unknown_{i}"

                name = occurrence.Name if hasattr(occurrence, 'Name') else os.path.basename(file_path)

                if file_path in bom_counts:
                    bom_counts[file_path]["quantity"] += 1
                else:
                    bom_counts[file_path] = {
                        "name": name,
                        "file_path": file_path,
                        "quantity": 1
                    }

            bom_items = list(bom_counts.values())

            return {
                "total_occurrences": occurrences.Count,
                "unique_parts": len(bom_items),
                "bom": bom_items
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_assembly_relations(self) -> Dict[str, Any]:
        """
        Get all assembly relations (constraints) in the active assembly.

        Iterates the Relations3d collection to report constraint types,
        status, and connected occurrences.

        Returns:
            Dict with list of relations and their properties
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Relations3d'):
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

                    try:
                        rel_info["status"] = rel.Status
                    except Exception:
                        pass

                    try:
                        rel_info["suppressed"] = rel.Suppressed
                    except Exception:
                        pass

                    try:
                        rel_info["name"] = rel.Name
                    except Exception:
                        pass

                    relation_list.append(rel_info)
                except Exception:
                    relation_list.append({"index": i - 1})

            return {
                "relations": relation_list,
                "count": len(relation_list)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_document_tree(self) -> Dict[str, Any]:
        """
        Get the hierarchical document tree of the active assembly.

        Recursively traverses occurrences and sub-occurrences to build
        a nested tree structure showing the full assembly hierarchy.

        Returns:
            Dict with nested tree of components
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Occurrences'):
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

                try:
                    node["visible"] = occ.Visible
                except Exception:
                    pass

                try:
                    node["suppressed"] = occ.IsSuppressed if hasattr(occ, 'IsSuppressed') else False
                except Exception:
                    pass

                # Recurse into sub-occurrences
                children = []
                try:
                    sub_occs = occ.SubOccurrences
                    if sub_occs and hasattr(sub_occs, 'Count'):
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
                "document": doc.Name if hasattr(doc, 'Name') else "Unknown"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def set_component_visibility(self, component_index: int, visible: bool) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}. Count: {occurrences.Count}"}

            occurrence = occurrences.Item(component_index + 1)
            occurrence.Visible = visible

            return {
                "status": "updated",
                "component_index": component_index,
                "visible": visible
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def is_subassembly(self, component_index: int) -> Dict[str, Any]:
        """
        Check if a component is a subassembly (vs a part).

        Args:
            component_index: 0-based index of the component

        Returns:
            Dict with is_subassembly boolean
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}. Count: {occurrences.Count}"}

            occurrence = occurrences.Item(component_index + 1)

            result = {"component_index": component_index}

            try:
                result["is_subassembly"] = occurrence.Subassembly
            except Exception:
                # Fallback: check if it has SubOccurrences
                try:
                    sub_occs = occurrence.SubOccurrences
                    result["is_subassembly"] = sub_occs.Count > 0 if hasattr(sub_occs, 'Count') else False
                except Exception:
                    result["is_subassembly"] = False

            try:
                result["name"] = occurrence.Name
            except Exception:
                pass

            return result
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_component_display_name(self, component_index: int) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}. Count: {occurrences.Count}"}

            occurrence = occurrences.Item(component_index + 1)

            result = {"component_index": component_index}

            try:
                result["display_name"] = occurrence.DisplayName
            except Exception:
                result["display_name"] = None

            try:
                result["name"] = occurrence.Name
            except Exception:
                pass

            try:
                result["file_name"] = occurrence.OccurrenceFileName
            except Exception:
                pass

            return result
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_occurrence_document(self, component_index: int) -> Dict[str, Any]:
        """
        Get document info for a component's source file.

        Args:
            component_index: 0-based index of the component

        Returns:
            Dict with document name, path, and type
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}. Count: {occurrences.Count}"}

            occurrence = occurrences.Item(component_index + 1)

            result = {"component_index": component_index}

            try:
                occ_doc = occurrence.OccurrenceDocument
                try:
                    result["document_name"] = occ_doc.Name
                except Exception:
                    pass
                try:
                    result["full_name"] = occ_doc.FullName
                except Exception:
                    pass
                try:
                    result["type"] = occ_doc.Type
                except Exception:
                    pass
                try:
                    result["read_only"] = occ_doc.ReadOnly
                except Exception:
                    pass
            except Exception:
                result["error_note"] = "Could not access OccurrenceDocument"

            try:
                result["file_name"] = occurrence.OccurrenceFileName
            except Exception:
                pass

            return result
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_sub_occurrences(self, component_index: int) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences
            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}. Count: {occurrences.Count}"}

            occurrence = occurrences.Item(component_index + 1)

            children = []
            try:
                sub_occs = occurrence.SubOccurrences
                if sub_occs and hasattr(sub_occs, 'Count'):
                    for j in range(1, sub_occs.Count + 1):
                        try:
                            child = sub_occs.Item(j)
                            child_info = {"index": j - 1}
                            try:
                                child_info["name"] = child.Name
                            except Exception:
                                child_info["name"] = f"SubOcc_{j}"
                            try:
                                child_info["file"] = child.OccurrenceFileName
                            except Exception:
                                pass
                            children.append(child_info)
                        except Exception:
                            children.append({"index": j - 1, "name": f"SubOcc_{j}"})
            except Exception:
                pass

            return {
                "component_index": component_index,
                "sub_occurrences": children,
                "count": len(children)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def delete_component(self, component_index: int) -> Dict[str, Any]:
        """
        Delete/remove a component from the assembly.

        Args:
            component_index: 0-based index of the component to remove

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}. Count: {occurrences.Count}"}

            occurrence = occurrences.Item(component_index + 1)
            name = occurrence.Name if hasattr(occurrence, 'Name') else f"Component_{component_index}"
            occurrence.Delete()

            return {
                "status": "deleted",
                "component_index": component_index,
                "name": name
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def ground_component(self, component_index: int, ground: bool = True) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}. Count: {occurrences.Count}"}

            occurrence = occurrences.Item(component_index + 1)

            if ground:
                # Add a ground constraint
                relations = doc.Relations3d
                relations.AddGround(occurrence)
                return {
                    "status": "grounded",
                    "component_index": component_index
                }
            else:
                # Find and delete ground relation for this occurrence
                relations = doc.Relations3d
                for i in range(relations.Count, 0, -1):
                    try:
                        rel = relations.Item(i)
                        # Ground relations have Type = 0
                        if hasattr(rel, 'Type') and rel.Type == 0:
                            rel.Delete()
                            return {
                                "status": "ungrounded",
                                "component_index": component_index
                            }
                    except Exception:
                        continue

                return {"error": "No ground relation found for this component"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def check_interference(self, component_index: Optional[int] = None) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if occurrences.Count < 2:
                return {"status": "no_interference", "message": "Need at least 2 components for interference check"}

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
                    NumInterferences=num_interferences
                )

                return {
                    "status": "checked",
                    "interference_found": interference_status.value != 0,
                    "num_interferences": num_interferences.value,
                    "component_checked": component_index
                }
            except Exception as e:
                # CheckInterference has complex COM signature; report what we can
                return {
                    "error": f"Interference check failed: {e}",
                    "note": "CheckInterference COM signature is complex. Use Solid Edge UI for reliable results.",
                    "traceback": traceback.format_exc()
                }

        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def replace_component(self, component_index: int, new_file_path: str) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            if not os.path.exists(new_file_path):
                return {"error": f"File not found: {new_file_path}"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}. Count: {occurrences.Count}"}

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
                "new_file": new_file_path
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_component_transform(self, component_index: int) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}. Count: {occurrences.Count}"}

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
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_structured_bom(self) -> Dict[str, Any]:
        """
        Get a hierarchical Bill of Materials with subassembly structure.

        Unlike get_bom() which returns a flat list, this preserves the
        parent-child hierarchy of subassemblies.

        Returns:
            Dict with structured BOM tree
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Occurrences'):
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

                try:
                    item["visible"] = occ.Visible
                except Exception:
                    pass

                try:
                    item["suppressed"] = occ.IsSuppressed
                except Exception:
                    item["suppressed"] = False

                # Check for sub-occurrences (subassembly)
                children = []
                try:
                    sub_occs = occ.SubOccurrences
                    if sub_occs and hasattr(sub_occs, 'Count') and sub_occs.Count > 0:
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
                "document": doc.Name if hasattr(doc, 'Name') else "Unknown"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def set_component_color(self, component_index: int, red: int, green: int, blue: int) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}. Count: {occurrences.Count}"}

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
                "color": [red, green, blue]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def occurrence_move(self, component_index: int, dx: float, dy: float, dz: float) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}. Count: {occurrences.Count}"}

            occurrence = occurrences.Item(component_index + 1)
            occurrence.Move(dx, dy, dz)

            return {
                "status": "moved",
                "component_index": component_index,
                "delta": [dx, dy, dz]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def occurrence_rotate(self, component_index: int,
                          axis_x1: float, axis_y1: float, axis_z1: float,
                          axis_x2: float, axis_y2: float, axis_z2: float,
                          angle: float) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}. Count: {occurrences.Count}"}

            occurrence = occurrences.Item(component_index + 1)
            angle_rad = math.radians(angle)
            occurrence.Rotate(axis_x1, axis_y1, axis_z1,
                              axis_x2, axis_y2, axis_z2,
                              angle_rad)

            return {
                "status": "rotated",
                "component_index": component_index,
                "axis": [[axis_x1, axis_y1, axis_z1], [axis_x2, axis_y2, axis_z2]],
                "angle_degrees": angle
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_occurrence_count(self) -> Dict[str, Any]:
        """
        Get the count of top-level occurrences in the assembly.

        Returns:
            Dict with occurrence count
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Occurrences'):
                return {"error": "Active document is not an assembly"}

            return {"count": doc.Occurrences.Count}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }
