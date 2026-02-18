"""
Unit tests for QueryManager backend methods (_materials.py mixin).

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
# GET MATERIAL TABLE
# ============================================================================


class TestGetMaterialTable:
    def test_with_variables(self, query_mgr):
        qm, doc = query_mgr

        var1 = MagicMock()
        var1.Name = "Density"
        var1.Value = 7850.0

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var1
        doc.Variables = variables

        result = qm.get_material_table()
        assert result["property_count"] >= 1
        assert "Density" in result["material_properties"]


# ============================================================================
# TIER 3: MATERIAL OPERATIONS
# ============================================================================


class TestGetMaterialList:
    def test_success(self, query_mgr):
        """Test getting material list."""
        qm, doc = query_mgr

        mat_table = MagicMock()
        mat_table.GetMaterialList.return_value = (3, ["Steel", "Aluminum", "Copper"])

        app = MagicMock()
        app.GetMaterialTable.return_value = mat_table
        qm.doc_manager.connection = MagicMock()
        qm.doc_manager.connection.get_application.return_value = app

        result = qm.get_material_list()
        assert result["count"] == 3
        assert "Steel" in result["materials"]
        assert "Aluminum" in result["materials"]

    def test_com_error(self, query_mgr):
        """Test handling COM errors."""
        qm, doc = query_mgr

        app = MagicMock()
        app.GetMaterialTable.side_effect = Exception("COM error")
        qm.doc_manager.connection = MagicMock()
        qm.doc_manager.connection.get_application.return_value = app

        result = qm.get_material_list()
        assert "error" in result


class TestSetMaterial:
    def test_success(self, query_mgr):
        """Test applying a material."""
        qm, doc = query_mgr

        mat_table = MagicMock()
        app = MagicMock()
        app.GetMaterialTable.return_value = mat_table
        qm.doc_manager.connection = MagicMock()
        qm.doc_manager.connection.get_application.return_value = app

        result = qm.set_material("Steel")
        assert result["status"] == "applied"
        assert result["material"] == "Steel"
        mat_table.ApplyMaterial.assert_called_once_with(doc, "Steel")

    def test_invalid_material(self, query_mgr):
        """Test applying non-existent material."""
        qm, doc = query_mgr

        mat_table = MagicMock()
        mat_table.ApplyMaterial.side_effect = Exception("Material not found")
        app = MagicMock()
        app.GetMaterialTable.return_value = mat_table
        qm.doc_manager.connection = MagicMock()
        qm.doc_manager.connection.get_application.return_value = app

        result = qm.set_material("InvalidMaterial")
        assert "error" in result


class TestGetMaterialProperty:
    def test_success(self, query_mgr):
        """Test getting a material property."""
        qm, doc = query_mgr

        mat_table = MagicMock()
        mat_table.GetMatPropValue.return_value = 7850.0
        app = MagicMock()
        app.GetMaterialTable.return_value = mat_table
        qm.doc_manager.connection = MagicMock()
        qm.doc_manager.connection.get_application.return_value = app

        result = qm.get_material_property("Steel", 0)
        assert result["material"] == "Steel"
        assert result["property_index"] == 0
        assert result["value"] == 7850.0
        mat_table.GetMatPropValue.assert_called_once_with("Steel", 0)

    def test_com_error(self, query_mgr):
        """Test handling COM error for invalid property."""
        qm, doc = query_mgr

        mat_table = MagicMock()
        mat_table.GetMatPropValue.side_effect = Exception("Invalid index")
        app = MagicMock()
        app.GetMaterialTable.return_value = mat_table
        qm.doc_manager.connection = MagicMock()
        qm.doc_manager.connection.get_application.return_value = app

        result = qm.get_material_property("Steel", 99)
        assert "error" in result


# ============================================================================
# BATCH 10: MATERIAL LIBRARY
# ============================================================================


class TestGetMaterialLibrary:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        mat1 = MagicMock()
        mat1.Name = "Steel"
        mat1.Density = 7850.0
        mat1.YoungsModulus = 2.1e11
        mat1.PoissonsRatio = 0.3

        mat2 = MagicMock()
        mat2.Name = "Aluminum"
        mat2.Density = 2700.0
        mat2.YoungsModulus = 7.0e10
        mat2.PoissonsRatio = 0.33

        mat_table = MagicMock()
        mat_table.Count = 2
        mat_table.Item.side_effect = lambda i: [None, mat1, mat2][i]
        doc.GetMaterialTable.return_value = mat_table

        result = qm.get_material_library()
        assert result["count"] == 2
        assert result["materials"][0]["name"] == "Steel"
        assert result["materials"][0]["density"] == 7850.0
        assert result["materials"][1]["name"] == "Aluminum"

    def test_empty(self, query_mgr):
        qm, doc = query_mgr
        mat_table = MagicMock()
        mat_table.Count = 0
        doc.GetMaterialTable.return_value = mat_table

        result = qm.get_material_library()
        assert result["count"] == 0
        assert result["materials"] == []

    def test_error(self, query_mgr):
        qm, doc = query_mgr
        doc.GetMaterialTable.side_effect = Exception("No material table")

        result = qm.get_material_library()
        assert "error" in result


class TestSetMaterialByName:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        mat = MagicMock()
        mat.Name = "Steel"
        mat.Density = 7850.0

        mat_table = MagicMock()
        mat_table.Count = 1
        mat_table.Item.return_value = mat
        doc.GetMaterialTable.return_value = mat_table

        result = qm.set_material_by_name("Steel")
        assert result["status"] == "applied"
        assert result["material"] == "Steel"

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        mat = MagicMock()
        mat.Name = "Steel"

        mat_table = MagicMock()
        mat_table.Count = 1
        mat_table.Item.return_value = mat
        doc.GetMaterialTable.return_value = mat_table

        result = qm.set_material_by_name("Titanium")
        assert "error" in result
        assert "Titanium" in result["error"]

    def test_error(self, query_mgr):
        qm, doc = query_mgr
        doc.GetMaterialTable.side_effect = Exception("COM error")

        result = qm.set_material_by_name("Steel")
        assert "error" in result


# ============================================================================
# LAYERS: GET LAYERS
# ============================================================================


class TestGetLayers:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        layer1 = MagicMock()
        layer1.Name = "Default"
        layer1.Show = True
        layer1.Locatable = True
        layer1.IsEmpty = False

        layer2 = MagicMock()
        layer2.Name = "Construction"
        layer2.Show = False
        layer2.Locatable = False
        layer2.IsEmpty = True

        layers = MagicMock()
        layers.Count = 2
        layers.Item.side_effect = lambda i: {1: layer1, 2: layer2}[i]
        doc.Layers = layers

        result = qm.get_layers()
        assert result["count"] == 2
        assert result["layers"][0]["name"] == "Default"
        assert result["layers"][0]["show"] is True
        assert result["layers"][1]["name"] == "Construction"
        assert result["layers"][1]["is_empty"] is True

    def test_no_layers(self, query_mgr):
        qm, doc = query_mgr
        del doc.Layers

        result = qm.get_layers()
        assert "error" in result


# ============================================================================
# LAYERS: ADD LAYER
# ============================================================================


class TestAddLayer:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        layers = MagicMock()
        layers.Count = 3
        layers.Add.return_value = MagicMock()
        doc.Layers = layers

        result = qm.add_layer("MyLayer")
        assert result["status"] == "added"
        assert result["name"] == "MyLayer"
        assert result["total_layers"] == 3
        layers.Add.assert_called_once_with("MyLayer")

    def test_no_layers_support(self, query_mgr):
        qm, doc = query_mgr
        del doc.Layers

        result = qm.add_layer("Test")
        assert "error" in result


# ============================================================================
# LAYERS: ACTIVATE LAYER
# ============================================================================


class TestActivateLayer:
    def test_by_index(self, query_mgr):
        qm, doc = query_mgr
        layer = MagicMock()
        layer.Name = "Layer1"
        layers = MagicMock()
        layers.Count = 3
        layers.Item.return_value = layer
        doc.Layers = layers

        result = qm.activate_layer(0)
        assert result["status"] == "activated"
        assert result["name"] == "Layer1"
        layer.Activate.assert_called_once()

    def test_by_name(self, query_mgr):
        qm, doc = query_mgr
        layer1 = MagicMock()
        layer1.Name = "Default"
        layer2 = MagicMock()
        layer2.Name = "Custom"
        layers = MagicMock()
        layers.Count = 2
        layers.Item.side_effect = lambda i: {1: layer1, 2: layer2}[i]
        doc.Layers = layers

        result = qm.activate_layer("Custom")
        assert result["status"] == "activated"
        assert result["name"] == "Custom"
        layer2.Activate.assert_called_once()

    def test_invalid_index(self, query_mgr):
        qm, doc = query_mgr
        layers = MagicMock()
        layers.Count = 2
        doc.Layers = layers

        result = qm.activate_layer(5)
        assert "error" in result

    def test_name_not_found(self, query_mgr):
        qm, doc = query_mgr
        layer1 = MagicMock()
        layer1.Name = "Default"
        layers = MagicMock()
        layers.Count = 1
        layers.Item.return_value = layer1
        doc.Layers = layers

        result = qm.activate_layer("NonExistent")
        assert "error" in result

    def test_no_layers_support(self, query_mgr):
        qm, doc = query_mgr
        del doc.Layers

        result = qm.activate_layer(0)
        assert "error" in result


# ============================================================================
# LAYERS: SET LAYER PROPERTIES
# ============================================================================


class TestSetLayerProperties:
    def test_set_show(self, query_mgr):
        qm, doc = query_mgr
        layer = MagicMock()
        layer.Name = "Layer1"
        layers = MagicMock()
        layers.Count = 1
        layers.Item.return_value = layer
        doc.Layers = layers

        result = qm.set_layer_properties(0, show=False)
        assert result["status"] == "updated"
        assert result["properties"]["show"] is False
        assert layer.Show is False

    def test_set_selectable(self, query_mgr):
        qm, doc = query_mgr
        layer = MagicMock()
        layer.Name = "Layer1"
        layers = MagicMock()
        layers.Count = 1
        layers.Item.return_value = layer
        doc.Layers = layers

        result = qm.set_layer_properties(0, selectable=True)
        assert result["status"] == "updated"
        assert result["properties"]["selectable"] is True
        assert layer.Locatable is True

    def test_set_both(self, query_mgr):
        qm, doc = query_mgr
        layer = MagicMock()
        layer.Name = "Layer1"
        layers = MagicMock()
        layers.Count = 1
        layers.Item.return_value = layer
        doc.Layers = layers

        result = qm.set_layer_properties(0, show=True, selectable=False)
        assert result["status"] == "updated"
        assert result["properties"]["show"] is True
        assert result["properties"]["selectable"] is False

    def test_by_name(self, query_mgr):
        qm, doc = query_mgr
        layer = MagicMock()
        layer.Name = "Custom"
        layers = MagicMock()
        layers.Count = 1
        layers.Item.return_value = layer
        doc.Layers = layers

        result = qm.set_layer_properties("Custom", show=False)
        assert result["status"] == "updated"
        assert result["name"] == "Custom"

    def test_no_layers_support(self, query_mgr):
        qm, doc = query_mgr
        del doc.Layers

        result = qm.set_layer_properties(0, show=True)
        assert "error" in result


# ============================================================================
# DELETE LAYER
# ============================================================================


class TestDeleteLayer:
    def test_success_by_name(self, query_mgr):
        qm, doc = query_mgr

        layer = MagicMock()
        layer.Name = "Layer1"
        layers = MagicMock()
        layers.Count = 2
        layers.Item.return_value = layer
        doc.Layers = layers

        result = qm.delete_layer("Layer1")
        assert isinstance(result, dict)

    def test_success_by_index(self, query_mgr):
        qm, doc = query_mgr

        layer = MagicMock()
        layer.Name = "Layer1"
        layers = MagicMock()
        layers.Count = 2
        layers.Item.return_value = layer
        doc.Layers = layers

        result = qm.delete_layer(0)
        assert isinstance(result, dict)

    def test_no_layers(self, query_mgr):
        qm, doc = query_mgr
        del doc.Layers

        result = qm.delete_layer("Layer1")
        assert "error" in result
