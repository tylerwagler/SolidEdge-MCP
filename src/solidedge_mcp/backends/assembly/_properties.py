"""Property operations for assembly components."""

import contextlib
from typing import Any

from solidedge_mcp.backends.errors import error_result

from ..logging import get_logger
from ._base import com_get

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

            # Occurrence exposes no Suppress member (verified against
            # assembly.tlb). Suppression goes through the document, which hands
            # back a SuppressComponent object; that object owns UnSuppress.
            if suppress:
                doc.SetSuppressComponent(occurrence)
                return {"status": "updated", "component": component_index, "suppressed": True}

            return {
                "error": (
                    "Unsuppressing a component is not reachable through COM automation. "
                    "AssemblyDocument.SetSuppressComponent returns the SuppressComponent "
                    "object that owns UnSuppress, and Solid Edge offers no way to look that "
                    "object up again for an already-suppressed occurrence. Unsuppress the "
                    "component in the Solid Edge UI."
                ),
                "unsupported": True,
                "component": component_index,
            }
        except Exception as e:
            return error_result(e)

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

            err = self._require_assembly(doc)
            if err:
                return err

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
            return error_result(e)

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

            err = self._require_assembly(doc)
            if err:
                return err

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)
            name = com_get(occurrence, "Name", f"Component_{component_index}")
            occurrence.Delete()

            return {"status": "deleted", "component_index": component_index, "name": name}
        except Exception as e:
            return error_result(e)

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

            err = self._require_assembly(doc)
            if err:
                return err

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
                        if com_get(rel, "Type") == 0:
                            rel.Delete()
                            return {"status": "ungrounded", "component_index": component_index}
                    except Exception:
                        continue

                return {"error": "No ground relation found for this component"}
        except Exception as e:
            return error_result(e)

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

            err = self._require_assembly(doc)
            if err:
                return err

            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {
                    "error": f"Invalid component index: "
                    f"{component_index}. "
                    f"Count: {occurrences.Count}"
                }

            occurrence = occurrences.Item(component_index + 1)

            red = max(0, min(255, red))
            green = max(0, min(255, green))
            blue = max(0, min(255, blue))

            # Occurrence has no SetColor, no Color, and no
            # UseOccurrenceColor/OccurrenceColor pair, so all three attempts
            # raised. FaceStyle is a get/put property, and a component is
            # coloured by handing it a style whose diffuse colour is wanted,
            # the same way a part body is.
            styles = com_get(doc, "FaceStyles")
            if styles is None:
                return {
                    "error": (
                        "This assembly has no FaceStyles collection, so a component "
                        "colour cannot be set."
                    )
                }

            name = f"MCP {red:02X}{green:02X}{blue:02X}"
            style = None
            with contextlib.suppress(Exception):
                style = styles.Item(name)
            if style is None:
                style = styles.Add(name, "")
            style.SetDiffuse(red / 255.0, green / 255.0, blue / 255.0)
            occurrence.FaceStyle = style

            return {
                "status": "updated",
                "component_index": component_index,
                "color": [red, green, blue],
                "hex": f"#{red:02x}{green:02x}{blue:02x}",
                "style": name,
            }
        except Exception as e:
            return error_result(e)

    def replace_component(
        self, component_index: int, new_file_path: str, replace_all: bool = False
    ) -> dict[str, Any]:
        """Replace one component with a different part or assembly file.

        ``Occurrence.Replace(NewOccurrenceFileName, ReplaceAll,
        [NewFamilyMemberName])`` takes two required arguments. This passed one,
        so the call raised; the fallback then assigned to
        ``Occurrence.OccurrenceFileName``, which the type library marks
        read-only, so that raised too and the caller got an error naming the
        wrong thing. Neither path could ever have replaced anything.

        Args:
            component_index: 0-based index of the component to replace.
            new_file_path: Path to the replacement file (.par or .asm).
            replace_all: Replace every occurrence of the same file, not just
                this one.

        Returns:
            Dict with replacement status
        """
        try:
            import os

            doc = self.doc_manager.get_active_document()

            err = self._require_assembly(doc)
            if err:
                return err

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

            occurrence.Replace(new_file_path, replace_all)

            return {
                "status": "replaced",
                "component_index": component_index,
                "old_name": old_name,
                "new_file": new_file_path,
                "replace_all": replace_all,
            }
        except Exception as e:
            return error_result(e)

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
            return error_result(e)

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
            return error_result(e)

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
            return error_result(e)

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

            # SwapFamilyMember(MemberName as VT_BSTR,
            #     SwapAllOccurrences as VT_BOOL)
            occurrence.SwapFamilyMember(new_member_name, False)

            return {
                "status": "swapped",
                "component_index": component_index,
                "new_member_name": new_member_name,
            }
        except Exception as e:
            return error_result(e)
