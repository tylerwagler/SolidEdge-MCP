"""Variable management and custom properties."""

import contextlib
from typing import Any

from solidedge_mcp.backends.errors import describe_exception, error_result

from ..comutil import com_get
from ..constants import VariableNameBy, seVariableTypeConstants
from ..logging import get_logger

_logger = get_logger(__name__)


class VariablesMixin:
    """Mixin providing variable and custom property methods."""

    doc_manager: Any

    def get_variables(self) -> dict[str, Any]:
        """
        Get all variables from the active document.

        Queries the Variables collection using Query() to list all
        variable names, values, and formulas.

        Returns:
            Dict with list of variables
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            var_list = []
            for i in range(1, variables.Count + 1):
                try:
                    var = variables.Item(i)
                    var_info = {
                        "index": i - 1,
                        "name": com_get(var, "DisplayName", f"Var_{i}"),
                    }
                    with contextlib.suppress(Exception):
                        var_info["value"] = var.Value
                    with contextlib.suppress(Exception):
                        var_info["formula"] = var.Formula
                    with contextlib.suppress(Exception):
                        var_info["units"] = var.Units
                    var_list.append(var_info)
                except Exception:
                    var_list.append({"index": i - 1, "name": f"Var_{i}"})

            return {"variables": var_list, "count": len(var_list)}
        except Exception as e:
            return error_result(e)

    @staticmethod
    def _find_variable(variables: Any, name: str) -> Any:
        """The variable whose display name is ``name``, or None.

        Only the lookup is guarded. Every method here used to search and act
        inside one ``try ... except Exception: continue``, so a failure while
        acting on the variable it had just found skipped past it and fell
        through to "not found" -- which is exactly what rename_variable
        reported for as long as it assigned to DisplayName, a property the
        type library marks read-only. The caller now acts outside the loop,
        where a failure surfaces as itself.
        """
        for i in range(1, com_get(variables, "Count", 0) + 1):
            with contextlib.suppress(Exception):
                var = variables.Item(i)
                if com_get(var, "DisplayName", "") == name:
                    return var
        return None

    def get_variable(self, name: str) -> dict[str, Any]:
        """
        Get a specific variable by name.

        Args:
            name: Variable display name (e.g., 'V1', 'Mass', 'Volume')

        Returns:
            Dict with variable value and info
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            var = self._find_variable(variables, name)
            if var is None:
                return {"error": f"Variable '{name}' not found"}

            result: dict[str, Any] = {"name": name}
            with contextlib.suppress(Exception):
                result["value"] = var.Value
            with contextlib.suppress(Exception):
                result["formula"] = var.Formula
            with contextlib.suppress(Exception):
                result["units"] = var.Units
            return result
        except Exception as e:
            return error_result(e)

    def set_variable(self, name: str, value: float) -> dict[str, Any]:
        """
        Set a variable's value by name.

        Args:
            name: Variable display name
            value: New value to set

        Returns:
            Dict with status and updated value
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            var = self._find_variable(variables, name)
            if var is None:
                return {"error": f"Variable '{name}' not found"}

            old_value = var.Value
            var.Value = value
            return {
                "status": "updated",
                "name": name,
                "old_value": old_value,
                "new_value": com_get(var, "Value", value),
            }
        except Exception as e:
            return error_result(e)

    def add_variable(
        self, name: str, formula: str, units_type: str | None = None
    ) -> dict[str, Any]:
        """
        Create a new user variable in the active document.

        Uses Variables.Add(pName, pFormula, [UnitsType]).
        Type library: Add(pName: VT_BSTR, pFormula: VT_BSTR, [UnitsType: VT_VARIANT]) -> variable*.

        A bare number in a formula is read in the *document's* units, not in
        meters: on an inch template "0.025" is 0.025 inch, and the variable
        ends up holding 0.000635. ``Value`` is always meters, so
        ``set_variable`` and a formula disagree unless the formula names its
        unit. Write "25 mm" to mean 25 mm. Verified on Solid Edge 2026.

        Solid Edge keeps a constant formula as a value and reports ``Formula``
        as empty; only an expression such as "AAA * 2" reads back. The
        ``value`` in the result is what Solid Edge actually computed.

        Args:
            name: Variable name (e.g., 'MyWidth', 'BoltDiameter')
            formula: Formula or value as a string. Give the unit -- "25 mm",
                not "0.025" -- unless you mean document units.
            units_type: Optional units type string

        Returns:
            Dict with status and variable info
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            if units_type is not None:
                var = variables.Add(name, formula, units_type)
            else:
                var = variables.Add(name, formula)

            result = {"status": "created", "name": name, "formula": formula}
            with contextlib.suppress(Exception):
                result["value"] = var.Value
            with contextlib.suppress(Exception):
                result["display_name"] = var.DisplayName

            return result
        except Exception as e:
            return error_result(e)

    def set_variable_formula(self, name: str, formula: str) -> dict[str, Any]:
        """
        Set the formula of an existing variable by display name.

        A bare number in a formula is read in the *document's* units, not in
        meters: on an inch template "0.025" is 0.025 inch, and the variable
        ends up holding 0.000635. ``Value`` is always meters, so
        ``set_variable`` and a formula disagree unless the formula names its
        unit. Write "25 mm" to mean 25 mm. Verified on Solid Edge 2026.

        Solid Edge keeps a constant formula as a value and reports ``Formula``
        as empty; only an expression such as "AAA * 2" reads back. The
        ``value`` in the result is what Solid Edge actually computed.

        Args:
            name: Variable display name
            formula: Formula or value as a string. Give the unit -- "25 mm",
                not "0.025" -- unless you mean document units.

        Returns:
            Dict with old/new formula and the value Solid Edge computed
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            var = self._find_variable(variables, name)
            if var is None:
                return {"error": f"Variable '{name}' not found"}

            old_formula = com_get(var, "Formula", "")
            var.Formula = formula
            result: dict[str, Any] = {
                "status": "updated",
                "name": name,
                "old_formula": old_formula,
                "new_formula": formula,
            }
            with contextlib.suppress(Exception):
                result["value"] = var.Value
            return result
        except Exception as e:
            return error_result(e)

    def query_variables(self, pattern: str = "*", case_insensitive: bool = True) -> dict[str, Any]:
        """Search variables by name pattern.

        ``Variables.Query(pFindCriterium, NamedBy, VarType, CaseInsensitive)``.
        VarType takes a ``seVariableTypeConstants`` value, and this used to pass
        0, which is not a member of that enum: Solid Edge answered every query
        with an empty collection, so variable search never found anything at
        all. Verified on Solid Edge 2026, where ``Query("*", 0, 0, True)``
        returns nothing and ``Query("*", 0, seVariableType_UserDefined, True)``
        returns the user variables.

        The enum has no "all" member, so each type is queried and the results
        merged. CaseInsensitive only takes effect when VarType is supplied --
        a bare ``Query("w*")`` misses ``Width`` -- which is the other reason
        to pass it explicitly.

        Args:
            pattern: Search pattern with wildcards (e.g., "*Length*", "V?").
            case_insensitive: Whether to ignore case (default True).

        Returns:
            Dict with matching variables.
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            matches: list[dict[str, Any]] = []
            seen: set[str] = set()
            failures: list[str] = []
            for var_type in (
                seVariableTypeConstants.seVariableType_UserDefined,
                seVariableTypeConstants.seVariableType_Dimension,
                seVariableTypeConstants.seVariableType_Simulation,
                seVariableTypeConstants.seVariableType_Text,
            ):
                try:
                    results = variables.Query(
                        pattern,
                        VariableNameBy.seVariableNameByUser,
                        var_type,
                        case_insensitive,
                    )
                except Exception as exc:  # noqa: BLE001
                    failures.append(describe_exception(exc))
                    continue
                if results is None:
                    continue
                for i in range(1, com_get(results, "Count", 0) + 1):
                    with contextlib.suppress(Exception):
                        var = results.Item(i)
                        name = com_get(var, "DisplayName", "") or com_get(var, "Name", f"var_{i}")
                        if name in seen:
                            continue
                        seen.add(name)
                        entry: dict[str, Any] = {"name": name}
                        with contextlib.suppress(Exception):
                            entry["value"] = var.Value
                        with contextlib.suppress(Exception):
                            entry["formula"] = var.Formula
                        matches.append(entry)

            # Every type raised: that is a broken query, not an empty result.
            if failures and len(failures) == 4:
                return {
                    "error": f"Variables.Query failed for every variable type: {failures[0]}",
                    "pattern": pattern,
                }

            return {"pattern": pattern, "matches": matches, "count": len(matches)}
        except Exception as e:
            return error_result(e)

    def get_variable_formula(self, name: str) -> dict[str, Any]:
        """
        Get the formula of a variable by name.

        Args:
            name: Variable display name

        Returns:
            Dict with variable formula string
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            for i in range(1, variables.Count + 1):
                try:
                    var = variables.Item(i)
                    display_name = com_get(var, "DisplayName", "")
                    if display_name == name:
                        result: dict[str, Any] = {"name": name}
                        try:
                            result["formula"] = var.Formula
                        except Exception:
                            result["formula"] = None
                        with contextlib.suppress(Exception):
                            result["value"] = var.Value
                        return result
                except Exception:
                    continue

            return {"error": f"Variable '{name}' not found"}
        except Exception as e:
            return error_result(e)

    def rename_variable(self, old_name: str, new_name: str) -> dict[str, Any]:
        """
        Rename a variable by changing its DisplayName.

        Finds the variable by its current display name and sets a new one.

        Args:
            old_name: Current variable display name
            new_name: New display name

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            var = self._find_variable(variables, old_name)
            if var is None:
                return {"error": f"Variable '{old_name}' not found"}

            # variable.DisplayName is read-only: Solid Edge 2026 answers
            # "Property 'Add.DisplayName' can not be set." Variables.PutName is
            # the rename, and it is verified against the live application.
            variables.PutName(var, new_name)
            return {
                "status": "renamed",
                "old_name": old_name,
                "new_name": new_name,
                "reads_back": variables.GetDisplayName(var),
            }
        except Exception as e:
            return error_result(e)

    def get_variable_names(self, name: str) -> dict[str, Any]:
        """
        Get both the DisplayName and SystemName of a variable.

        Useful for understanding how Solid Edge internally identifies
        variables vs. how they appear in the UI.

        Args:
            name: Variable display name to look up

        Returns:
            Dict with display_name and system_name
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            for i in range(1, variables.Count + 1):
                try:
                    var = variables.Item(i)
                    display_name = com_get(var, "DisplayName", "")
                    if display_name == name:
                        result = {"display_name": display_name}
                        try:
                            result["system_name"] = var.Name
                        except Exception:
                            result["system_name"] = None
                        with contextlib.suppress(Exception):
                            result["value"] = var.Value
                        return result
                except Exception:
                    continue

            return {"error": f"Variable '{name}' not found"}
        except Exception as e:
            return error_result(e)

    def translate_variable(self, name: str) -> dict[str, Any]:
        """
        Translate (look up) a variable by name using Variables.Translate().

        Returns the variable dispatch object's Name, Value, and Formula.

        Args:
            name: Variable name to translate

        Returns:
            Dict with variable info (name, value, formula)
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            var = variables.Translate(name)

            result = {"status": "success", "input_name": name}
            with contextlib.suppress(Exception):
                result["name"] = var.Name
            with contextlib.suppress(Exception):
                result["display_name"] = var.DisplayName
            with contextlib.suppress(Exception):
                result["value"] = var.Value
            with contextlib.suppress(Exception):
                result["formula"] = var.Formula
            return result
        except Exception as e:
            return error_result(e)

    def copy_variable_to_clipboard(self, name: str) -> dict[str, Any]:
        """
        Copy a variable definition to the clipboard.

        Uses Variables.CopyToClipboard(name). The variable can then be pasted
        into another document using add_variable_from_clipboard.

        Args:
            name: Variable display name to copy

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables
            variables.CopyToClipboard(name)
            return {"status": "copied", "name": name}
        except Exception as e:
            return error_result(e)

    def add_variable_from_clipboard(
        self, name: str, units_type: str | None = None
    ) -> dict[str, Any]:
        """
        Add a variable from the clipboard.

        Uses Variables.AddFromClipboard(name) or AddFromClipboard(name, units).
        The variable must have been previously copied with copy_variable_to_clipboard.

        Args:
            name: Name for the new variable
            units_type: Optional units type string

        Returns:
            Dict with status and variable info
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            if units_type:
                new_var = variables.AddFromClipboard(name, units_type)
            else:
                new_var = variables.AddFromClipboard(name)

            result = {"status": "added", "name": name}
            with contextlib.suppress(Exception):
                result["value"] = new_var.Value
            return result
        except Exception as e:
            return error_result(e)

    # =================================================================
    # CUSTOM PROPERTIES
    # =================================================================

    def get_custom_properties(self) -> dict[str, Any]:
        """
        Get all custom properties from the active document.

        Accesses the PropertySets collection to retrieve custom properties.

        Returns:
            Dict with list of custom properties (name/value pairs)
        """
        try:
            doc = self.doc_manager.get_active_document()
            prop_sets = doc.Properties

            properties = {}

            # Iterate through property sets to find Custom
            for ps_idx in range(1, prop_sets.Count + 1):
                try:
                    ps = prop_sets.Item(ps_idx)
                    ps_name = com_get(ps, "Name", f"Set_{ps_idx}")

                    props = {}
                    for p_idx in range(1, ps.Count + 1):
                        try:
                            prop = ps.Item(p_idx)
                            prop_name = com_get(prop, "Name", f"Prop_{p_idx}")
                            try:
                                props[prop_name] = prop.Value
                            except Exception:
                                props[prop_name] = None
                        except Exception:
                            continue

                    properties[ps_name] = props
                except Exception:
                    continue

            return {"property_sets": properties, "count": len(properties)}
        except Exception as e:
            return error_result(e)

    def set_custom_property(self, name: str, value: str) -> dict[str, Any]:
        """
        Set or create a custom property.

        Creates the property if it doesn't exist, updates it if it does.

        Args:
            name: Property name
            value: Property value (string)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            prop_sets = doc.Properties

            # Find "Custom" property set (typically the last one, index varies)
            custom_ps = None
            for ps_idx in range(1, prop_sets.Count + 1):
                try:
                    ps = prop_sets.Item(ps_idx)
                    if com_get(ps, "Name") == "Custom":
                        custom_ps = ps
                        break
                except Exception:
                    continue

            if custom_ps is None:
                return {"error": "Custom property set not found"}

            # Check if property exists
            for p_idx in range(1, custom_ps.Count + 1):
                try:
                    prop = custom_ps.Item(p_idx)
                    if com_get(prop, "Name") == name:
                        old_value = prop.Value
                        prop.Value = value
                        return {
                            "status": "updated",
                            "name": name,
                            "old_value": old_value,
                            "new_value": value,
                        }
                except Exception:
                    continue

            # Property doesn't exist, add it
            custom_ps.Add(name, value)
            return {"status": "created", "name": name, "value": value}
        except Exception as e:
            return error_result(e)

    def delete_custom_property(self, name: str) -> dict[str, Any]:
        """
        Delete a custom property by name.

        Args:
            name: Property name to delete

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            prop_sets = doc.Properties

            # Find "Custom" property set
            for ps_idx in range(1, prop_sets.Count + 1):
                try:
                    ps = prop_sets.Item(ps_idx)
                    if com_get(ps, "Name") == "Custom":
                        for p_idx in range(1, ps.Count + 1):
                            try:
                                prop = ps.Item(p_idx)
                                if com_get(prop, "Name") == name:
                                    prop.Delete()
                                    return {"status": "deleted", "name": name}
                            except Exception:
                                continue
                        return {"error": f"Property '{name}' not found in Custom set"}
                except Exception:
                    continue

            return {"error": "Custom property set not found"}
        except Exception as e:
            return error_result(e)
