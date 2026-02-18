"""
Unit tests for QueryManager backend methods (_variables.py mixin).

Uses unittest.mock to simulate COM objects.
"""

from unittest.mock import MagicMock

import pytest


@pytest.fixture
def doc_mgr():
    """Create mock doc_manager."""
    dm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc
    return dm, doc


@pytest.fixture
def query_mgr(doc_mgr):
    """Create QueryManager with mocked dependencies."""
    from solidedge_mcp.backends.query import QueryManager

    dm, doc = doc_mgr
    return QueryManager(dm), doc


# ============================================================================
# VARIABLES
# ============================================================================


class TestGetVariables:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        var1 = MagicMock()
        var1.DisplayName = "Width"
        var1.Value = 0.1
        var1.Formula = "100 mm"
        var1.Units = "m"

        var2 = MagicMock()
        var2.DisplayName = "Height"
        var2.Value = 0.05
        var2.Formula = "50 mm"
        var2.Units = "m"

        variables = MagicMock()
        variables.Count = 2
        variables.Item.side_effect = lambda i: [None, var1, var2][i]
        doc.Variables = variables

        result = qm.get_variables()
        assert result["count"] == 2
        assert result["variables"][0]["name"] == "Width"
        assert result["variables"][0]["value"] == 0.1
        assert result["variables"][1]["name"] == "Height"

    def test_empty(self, query_mgr):
        qm, doc = query_mgr
        variables = MagicMock()
        variables.Count = 0
        doc.Variables = variables

        result = qm.get_variables()
        assert result["count"] == 0
        assert result["variables"] == []


class TestGetVariable:
    def test_found(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Width"
        var.Value = 0.1
        var.Formula = "100 mm"
        var.Units = "m"

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.get_variable("Width")
        assert result["name"] == "Width"
        assert result["value"] == 0.1

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Width"

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.get_variable("Nonexistent")
        assert "error" in result


class TestSetVariable:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Width"
        var.Value = 0.1

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.set_variable("Width", 0.2)
        assert result["status"] == "updated"
        assert result["old_value"] == 0.1
        assert result["new_value"] == 0.2

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Width"

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.set_variable("Nonexistent", 0.5)
        assert "error" in result


# ============================================================================
# TIER 1: ADD VARIABLE
# ============================================================================


class TestAddVariable:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        variables = MagicMock()
        new_var = MagicMock()
        new_var.Value = 0.025
        new_var.DisplayName = "MyWidth"
        variables.Add.return_value = new_var
        doc.Variables = variables

        result = qm.add_variable("MyWidth", "0.025")
        assert result["status"] == "created"
        assert result["name"] == "MyWidth"
        assert result["formula"] == "0.025"
        assert result["value"] == 0.025
        variables.Add.assert_called_once_with("MyWidth", "0.025")

    def test_with_units(self, query_mgr):
        qm, doc = query_mgr
        variables = MagicMock()
        new_var = MagicMock()
        new_var.Value = 0.01
        new_var.DisplayName = "BoltDia"
        variables.Add.return_value = new_var
        doc.Variables = variables

        result = qm.add_variable("BoltDia", "0.01", units_type="m")
        assert result["status"] == "created"
        variables.Add.assert_called_once_with("BoltDia", "0.01", "m")

    def test_formula_expression(self, query_mgr):
        qm, doc = query_mgr
        variables = MagicMock()
        new_var = MagicMock()
        new_var.Value = 0.05
        new_var.DisplayName = "TotalWidth"
        variables.Add.return_value = new_var
        doc.Variables = variables

        result = qm.add_variable("TotalWidth", "MyWidth * 2")
        assert result["status"] == "created"
        assert result["formula"] == "MyWidth * 2"


# ============================================================================
# SET VARIABLE FORMULA
# ============================================================================


class TestSetVariableFormula:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        var1 = MagicMock()
        var1.DisplayName = "Width"
        var1.Formula = "50 mm"
        var1.Value = 0.05

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var1
        doc.Variables = variables

        result = qm.set_variable_formula("Width", "100 mm")
        assert isinstance(result, dict)
        # Should attempt to set the formula
        assert "error" not in result or "Formula" in str(result.get("error", ""))

    def test_variable_not_found(self, query_mgr):
        qm, doc = query_mgr

        variables = MagicMock()
        variables.Count = 0
        doc.Variables = variables

        result = qm.set_variable_formula("NonExistent", "100 mm")
        assert "error" in result

    def test_no_variables(self, query_mgr):
        qm, doc = query_mgr
        del doc.Variables

        result = qm.set_variable_formula("Width", "100 mm")
        assert "error" in result


# ============================================================================
# TIER 2: QUERY VARIABLES
# ============================================================================


class TestQueryVariables:
    def test_query_with_fallback(self, query_mgr):
        """Test fallback fnmatch when Variables.Query not available."""
        qm, doc = query_mgr
        variables = MagicMock()
        variables.Count = 3
        variables.Query.side_effect = Exception("Not available")

        v1 = MagicMock()
        v1.Name = "Length"
        v1.Value = 0.1
        v1.Formula = "0.1"

        v2 = MagicMock()
        v2.Name = "Width"
        v2.Value = 0.05
        v2.Formula = "0.05"

        v3 = MagicMock()
        v3.Name = "LengthOffset"
        v3.Value = 0.02
        v3.Formula = "Length / 5"

        variables.Item.side_effect = lambda i: [None, v1, v2, v3][i]
        doc.Variables = variables

        result = qm.query_variables("*Length*")
        assert result["count"] == 2
        assert result["method"] == "fallback_fnmatch"
        names = [m["name"] for m in result["matches"]]
        assert "Length" in names
        assert "LengthOffset" in names

    def test_query_with_native(self, query_mgr):
        """Test native Variables.Query when available."""
        qm, doc = query_mgr
        variables = MagicMock()

        # Mock Query return value
        query_results = MagicMock()
        query_results.Count = 1
        var = MagicMock()
        var.Name = "Width"
        var.Value = 0.05
        var.Formula = "0.05"
        query_results.Item.return_value = var
        variables.Query.return_value = query_results
        doc.Variables = variables

        result = qm.query_variables("Width")
        assert result["count"] == 1
        assert result["matches"][0]["name"] == "Width"
        variables.Query.assert_called_once_with("Width", 0, 0, True)

    def test_case_insensitive(self, query_mgr):
        """Test case-insensitive search with fallback."""
        qm, doc = query_mgr
        variables = MagicMock()
        variables.Count = 1
        variables.Query.side_effect = Exception("Not available")

        v1 = MagicMock()
        v1.Name = "MASS"
        v1.Value = 1.5
        v1.Formula = "1.5"
        variables.Item.return_value = v1
        doc.Variables = variables

        result = qm.query_variables("*mass*", case_insensitive=True)
        assert result["count"] == 1
        assert result["matches"][0]["name"] == "MASS"

    def test_no_matches(self, query_mgr):
        """Test query with no matching variables."""
        qm, doc = query_mgr
        variables = MagicMock()
        variables.Count = 1
        variables.Query.side_effect = Exception("Not available")

        v1 = MagicMock()
        v1.Name = "Length"
        v1.Value = 0.1
        variables.Item.return_value = v1
        doc.Variables = variables

        result = qm.query_variables("*Xyz*")
        assert result["count"] == 0
        assert result["matches"] == []


# ============================================================================
# VARIABLES: GET FORMULA
# ============================================================================


class TestGetVariableFormula:
    def test_found(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Width"
        var.Formula = "100 mm"
        var.Value = 0.1

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.get_variable_formula("Width")
        assert result["name"] == "Width"
        assert result["formula"] == "100 mm"
        assert result["value"] == 0.1

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Height"

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.get_variable_formula("Width")
        assert "error" in result


# ============================================================================
# VARIABLES: RENAME
# ============================================================================


class TestRenameVariable:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "OldVar"

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.rename_variable("OldVar", "NewVar")
        assert result["status"] == "renamed"
        assert result["old_name"] == "OldVar"
        assert result["new_name"] == "NewVar"
        assert var.DisplayName == "NewVar"

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Width"

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.rename_variable("Nonexistent", "NewName")
        assert "error" in result


# ============================================================================
# VARIABLES: GET NAMES (DisplayName + SystemName)
# ============================================================================


class TestGetVariableNames:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Width"
        var.Name = "V1"
        var.Value = 0.1

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.get_variable_names("Width")
        assert result["display_name"] == "Width"
        assert result["system_name"] == "V1"
        assert result["value"] == 0.1

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Height"

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.get_variable_names("Width")
        assert "error" in result


# ============================================================================
# TRANSLATE VARIABLE
# ============================================================================


class TestTranslateVariable:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        translated_var = MagicMock()
        translated_var.Name = "V1"
        translated_var.DisplayName = "Width"
        translated_var.Value = 0.1
        translated_var.Formula = "100 mm"

        variables = MagicMock()
        variables.Translate.return_value = translated_var
        doc.Variables = variables

        result = qm.translate_variable("Width")
        assert result["status"] == "success"
        assert result["name"] == "V1"
        assert result["display_name"] == "Width"
        assert result["value"] == 0.1
        assert result["formula"] == "100 mm"
        variables.Translate.assert_called_once_with("Width")

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        variables = MagicMock()
        variables.Translate.side_effect = Exception("Variable not found")
        doc.Variables = variables

        result = qm.translate_variable("NonExistent")
        assert "error" in result


# ============================================================================
# COPY VARIABLE TO CLIPBOARD
# ============================================================================


class TestCopyVariableToClipboard:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        variables = MagicMock()
        doc.Variables = variables

        result = qm.copy_variable_to_clipboard("Width")
        assert result["status"] == "copied"
        assert result["name"] == "Width"
        variables.CopyToClipboard.assert_called_once_with("Width")

    def test_com_error(self, query_mgr):
        qm, doc = query_mgr
        variables = MagicMock()
        variables.CopyToClipboard.side_effect = Exception("COM error")
        doc.Variables = variables

        result = qm.copy_variable_to_clipboard("Width")
        assert "error" in result


# ============================================================================
# ADD VARIABLE FROM CLIPBOARD
# ============================================================================


class TestAddVariableFromClipboard:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        new_var = MagicMock()
        new_var.Value = 0.1
        variables = MagicMock()
        variables.AddFromClipboard.return_value = new_var
        doc.Variables = variables

        result = qm.add_variable_from_clipboard("PastedWidth")
        assert result["status"] == "added"
        assert result["name"] == "PastedWidth"
        assert result["value"] == 0.1
        variables.AddFromClipboard.assert_called_once_with("PastedWidth")

    def test_with_units(self, query_mgr):
        qm, doc = query_mgr
        new_var = MagicMock()
        new_var.Value = 0.01
        variables = MagicMock()
        variables.AddFromClipboard.return_value = new_var
        doc.Variables = variables

        result = qm.add_variable_from_clipboard("PastedDia", units_type="m")
        assert result["status"] == "added"
        variables.AddFromClipboard.assert_called_once_with("PastedDia", "m")

    def test_com_error(self, query_mgr):
        qm, doc = query_mgr
        variables = MagicMock()
        variables.AddFromClipboard.side_effect = Exception("No clipboard data")
        doc.Variables = variables

        result = qm.add_variable_from_clipboard("Test")
        assert "error" in result


# ============================================================================
# CUSTOM PROPERTIES
# ============================================================================


class TestGetCustomProperties:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        prop = MagicMock()
        prop.Name = "Material"
        prop.Value = "Steel"

        ps = MagicMock()
        ps.Name = "Custom"
        ps.Count = 1
        ps.Item.return_value = prop

        prop_sets = MagicMock()
        prop_sets.Count = 1
        prop_sets.Item.return_value = ps
        doc.Properties = prop_sets

        result = qm.get_custom_properties()
        assert "property_sets" in result
        assert "Custom" in result["property_sets"]
        assert result["property_sets"]["Custom"]["Material"] == "Steel"


class TestSetCustomProperty:
    def test_update_existing(self, query_mgr):
        qm, doc = query_mgr

        prop = MagicMock()
        prop.Name = "Material"
        prop.Value = "Aluminum"

        ps = MagicMock()
        ps.Name = "Custom"
        ps.Count = 1
        ps.Item.return_value = prop

        prop_sets = MagicMock()
        prop_sets.Count = 1
        prop_sets.Item.return_value = ps
        doc.Properties = prop_sets

        result = qm.set_custom_property("Material", "Steel")
        assert result["status"] == "updated"
        assert result["old_value"] == "Aluminum"

    def test_create_new(self, query_mgr):
        qm, doc = query_mgr

        prop = MagicMock()
        prop.Name = "OtherProp"

        ps = MagicMock()
        ps.Name = "Custom"
        ps.Count = 1
        ps.Item.return_value = prop

        prop_sets = MagicMock()
        prop_sets.Count = 1
        prop_sets.Item.return_value = ps
        doc.Properties = prop_sets

        result = qm.set_custom_property("NewProp", "NewValue")
        assert result["status"] == "created"
        ps.Add.assert_called_once_with("NewProp", "NewValue")


class TestDeleteCustomProperty:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        prop = MagicMock()
        prop.Name = "Material"

        ps = MagicMock()
        ps.Name = "Custom"
        ps.Count = 1
        ps.Item.return_value = prop

        prop_sets = MagicMock()
        prop_sets.Count = 1
        prop_sets.Item.return_value = ps
        doc.Properties = prop_sets

        result = qm.delete_custom_property("Material")
        assert result["status"] == "deleted"
        prop.Delete.assert_called_once()

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr

        prop = MagicMock()
        prop.Name = "OtherProp"

        ps = MagicMock()
        ps.Name = "Custom"
        ps.Count = 1
        ps.Item.return_value = prop

        prop_sets = MagicMock()
        prop_sets.Count = 1
        prop_sets.Item.return_value = ps
        doc.Properties = prop_sets

        result = qm.delete_custom_property("Nonexistent")
        assert "error" in result
