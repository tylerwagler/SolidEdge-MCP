"""Variable management and custom properties."""

import contextlib
import traceback
from typing import Any, Optional

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
                        "name": var.DisplayName if hasattr(var, "DisplayName") else f"Var_{i}",
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            # Search for the variable by display name
            for i in range(1, variables.Count + 1):
                try:
                    var = variables.Item(i)
                    display_name = var.DisplayName if hasattr(var, "DisplayName") else ""
                    if display_name == name:
                        result = {"name": name, "index": i - 1}
                        with contextlib.suppress(Exception):
                            result["value"] = var.Value
                        with contextlib.suppress(Exception):
                            result["formula"] = var.Formula
                        with contextlib.suppress(Exception):
                            result["units"] = var.Units
                        return result
                except Exception:
                    continue

            return {"error": f"Variable '{name}' not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            for i in range(1, variables.Count + 1):
                try:
                    var = variables.Item(i)
                    display_name = var.DisplayName if hasattr(var, "DisplayName") else ""
                    if display_name == name:
                        old_value = var.Value
                        var.Value = value
                        return {
                            "status": "updated",
                            "name": name,
                            "old_value": old_value,
                            "new_value": value,
                        }
                except Exception:
                    continue

            return {"error": f"Variable '{name}' not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_variable(
        self, name: str, formula: str, units_type: Optional[str] = None
    ) -> dict[str, Any]:
        """
        Create a new user variable in the active document.

        Uses Variables.Add(pName, pFormula, [UnitsType]).
        Type library: Add(pName: VT_BSTR, pFormula: VT_BSTR, [UnitsType: VT_VARIANT]) -> variable*.

        Args:
            name: Variable name (e.g., 'MyWidth', 'BoltDiameter')
            formula: Variable formula/value as string (e.g., '0.025', 'V1 * 2')
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
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_variable_formula(self, name: str, formula: str) -> dict[str, Any]:
        """
        Set the formula of an existing variable by display name.

        Args:
            name: Variable display name
            formula: New formula string (e.g., '0.025', 'V1 * 2')

        Returns:
            Dict with old/new formula and current value
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            for i in range(1, variables.Count + 1):
                try:
                    var = variables.Item(i)
                    display_name = var.DisplayName if hasattr(var, "DisplayName") else ""
                    if display_name == name:
                        old_formula = ""
                        with contextlib.suppress(Exception):
                            old_formula = var.Formula
                        var.Formula = formula
                        result = {
                            "status": "updated",
                            "name": name,
                            "old_formula": old_formula,
                            "new_formula": formula,
                        }
                        with contextlib.suppress(Exception):
                            result["value"] = var.Value
                        return result
                except Exception:
                    continue

            return {"error": f"Variable '{name}' not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def query_variables(self, pattern: str = "*", case_insensitive: bool = True) -> dict[str, Any]:
        """
        Search variables by name pattern.

        Uses Variables.Query(pFindCriterium, NamedBy, VarType, CaseInsensitive)
        to find matching variables. Supports wildcards (* and ?).

        Args:
            pattern: Search pattern with wildcards (e.g., "*Length*", "V?")
            case_insensitive: Whether to ignore case (default True)

        Returns:
            Dict with matching variables
        """
        try:
            doc = self.doc_manager.get_active_document()
            variables = doc.Variables

            # Variables.Query returns a collection of matching variables
            # NamedBy: 0 = DisplayName, 1 = SystemName
            # VarType: 0 = All, 1 = Dimensions, 2 = UserVariables
            try:
                results = variables.Query(pattern, 0, 0, case_insensitive)
            except Exception:
                # Fallback: manual filtering if Query method not available
                matches = []
                import fnmatch

                for i in range(1, variables.Count + 1):
                    try:
                        var = variables.Item(i)
                        name = var.Name if hasattr(var, "Name") else str(i)
                        if case_insensitive:
                            match = fnmatch.fnmatch(name.lower(), pattern.lower())
                        else:
                            match = fnmatch.fnmatch(name, pattern)
                        if match:
                            entry = {"name": name}
                            with contextlib.suppress(Exception):
                                entry["value"] = var.Value
                            with contextlib.suppress(Exception):
                                entry["formula"] = var.Formula
                            matches.append(entry)
                    except Exception:
                        continue

                return {
                    "pattern": pattern,
                    "matches": matches,
                    "count": len(matches),
                    "method": "fallback_fnmatch",
                }

            # Process Query results
            matches = []
            if results is not None:
                try:
                    for i in range(1, results.Count + 1):
                        var = results.Item(i)
                        entry = {}
                        try:
                            entry["name"] = var.Name
                        except Exception:
                            entry["name"] = f"var_{i}"
                        with contextlib.suppress(Exception):
                            entry["value"] = var.Value
                        with contextlib.suppress(Exception):
                            entry["formula"] = var.Formula
                        matches.append(entry)
                except Exception:
                    pass

            return {"pattern": pattern, "matches": matches, "count": len(matches)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
                    display_name = var.DisplayName if hasattr(var, "DisplayName") else ""
                    if display_name == name:
                        result = {"name": name}
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            for i in range(1, variables.Count + 1):
                try:
                    var = variables.Item(i)
                    display_name = var.DisplayName if hasattr(var, "DisplayName") else ""
                    if display_name == old_name:
                        var.DisplayName = new_name
                        return {"status": "renamed", "old_name": old_name, "new_name": new_name}
                except Exception:
                    continue

            return {"error": f"Variable '{old_name}' not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
                    display_name = var.DisplayName if hasattr(var, "DisplayName") else ""
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_variable_from_clipboard(
        self, name: str, units_type: Optional[str] = None
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
                    ps_name = ps.Name if hasattr(ps, "Name") else f"Set_{ps_idx}"

                    props = {}
                    for p_idx in range(1, ps.Count + 1):
                        try:
                            prop = ps.Item(p_idx)
                            prop_name = prop.Name if hasattr(prop, "Name") else f"Prop_{p_idx}"
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
                    if hasattr(ps, "Name") and ps.Name == "Custom":
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
                    if hasattr(prop, "Name") and prop.Name == name:
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
                    if hasattr(ps, "Name") and ps.Name == "Custom":
                        for p_idx in range(1, ps.Count + 1):
                            try:
                                prop = ps.Item(p_idx)
                                if hasattr(prop, "Name") and prop.Name == name:
                                    prop.Delete()
                                    return {"status": "deleted", "name": name}
                            except Exception:
                                continue
                        return {"error": f"Property '{name}' not found in Custom set"}
                except Exception:
                    continue

            return {"error": "Custom property set not found"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
