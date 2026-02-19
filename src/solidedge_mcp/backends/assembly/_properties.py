"""Property operations for assembly components."""

import contextlib
import traceback
from typing import Any

from ..logging import get_logger

_logger = get_logger(__name__)


class PropertiesMixin:
    """Mixin providing component property modification methods."""

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
