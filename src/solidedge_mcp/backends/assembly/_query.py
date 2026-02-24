"""Query operations for assembly components."""

import contextlib
import os
import traceback
from typing import Any

from ..logging import get_logger

_logger = get_logger(__name__)


class QueryMixin:
    """Mixin providing assembly query/interrogation methods."""

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

            result: dict[str, Any] = {"component_index": component_index}

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

            result: dict[str, Any] = {"component_index": component_index}

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
                            child_info: dict[str, Any] = {"index": j - 1}
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

            def build_bom_item(occ: Any, depth: int = 0) -> dict[str, Any]:
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
                    rel_info: dict[str, Any] = {"index": i - 1}

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

            def traverse_occurrence(occ: Any, depth: int = 0) -> dict[str, Any]:
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
