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


def fake_material_table(libraries=None, current="", properties=None):
    """A MatTable as Solid Edge really exposes it.

    It is not a collection: no Count, no Item, and no Material object to read
    properties off. GetMaterialLibraryList returns (names, count) while
    GetMaterialListFromLibrary returns (count, names), which is the opposite
    order.
    """
    libraries = libraries if libraries is not None else {"Materials": ["Steel", "Aluminum"]}
    properties = properties or {}

    mat_table = MagicMock()
    del mat_table.Count
    del mat_table.Item
    mat_table.GetMaterialLibraryList.return_value = (tuple(libraries), len(libraries))
    mat_table.GetMaterialListFromLibrary.side_effect = lambda lib: (
        len(libraries[lib]),
        tuple(libraries[lib]),
    )
    state = {"current": current}
    mat_table.GetCurrentMaterialName.side_effect = lambda _doc: state["current"]

    def apply(_doc, name, _lib):
        state["current"] = name

    mat_table.ApplyMaterialToDoc.side_effect = apply
    mat_table.GetMaterialPropValueFromDoc.side_effect = lambda _doc, index: properties[index]
    return mat_table


def wire_material_table(qm, mat_table):
    app = MagicMock()
    app.GetMaterialTable.return_value = mat_table
    qm.doc_manager.connection = MagicMock()
    qm.doc_manager.connection.get_application.return_value = app
    return app


class TestGetMaterialList:
    """MatTable.GetMaterialList raises E_FAIL; the libraries are the way in."""

    def test_success(self, query_mgr):
        qm, _doc = query_mgr
        wire_material_table(
            qm,
            fake_material_table({"Materials": ["Steel", "Aluminum"], "Materials-DIN": ["S235"]}),
        )

        result = qm.get_material_list()

        assert result["count"] == 3
        assert "Steel" in result["materials"]
        assert result["libraries"]["Materials-DIN"] == ["S235"]

    def test_no_materials_is_an_error_not_an_empty_success(self, query_mgr):
        qm, _doc = query_mgr
        wire_material_table(qm, fake_material_table({}))

        result = qm.get_material_list()

        assert "error" in result
        assert result["libraries"] == []

    def test_one_unreadable_library_does_not_lose_the_others(self, query_mgr):
        qm, _doc = query_mgr
        mat_table = fake_material_table({"Materials": ["Steel"], "Broken": ["x"]})

        def read(lib):
            if lib == "Broken":
                raise Exception("library missing")
            return (1, ("Steel",))

        mat_table.GetMaterialListFromLibrary.side_effect = read
        wire_material_table(qm, mat_table)

        result = qm.get_material_list()

        assert result["materials"] == ["Steel"]

    def test_com_error(self, query_mgr):
        qm, _doc = query_mgr
        app = MagicMock()
        app.GetMaterialTable.side_effect = Exception("COM error")
        qm.doc_manager.connection = MagicMock()
        qm.doc_manager.connection.get_application.return_value = app

        assert "error" in qm.get_material_list()


class TestSetMaterial:
    """ApplyMaterialToDoc takes the library too, and the result is verified."""

    def test_success(self, query_mgr):
        qm, doc = query_mgr
        mat_table = fake_material_table()
        wire_material_table(qm, mat_table)

        result = qm.set_material("Steel")

        assert result["status"] == "applied"
        assert result["material"] == "Steel"
        assert result["library"] == "Materials"
        mat_table.ApplyMaterialToDoc.assert_called_once_with(doc, "Steel", "Materials")

    def test_matches_a_name_case_insensitively(self, query_mgr):
        qm, _doc = query_mgr
        wire_material_table(qm, fake_material_table())

        assert qm.set_material("steel")["status"] == "applied"

    def test_invalid_material_lists_what_is_available(self, query_mgr):
        qm, _doc = query_mgr
        wire_material_table(qm, fake_material_table())

        result = qm.set_material("Unobtainium")

        assert "error" in result
        assert "Steel" in result["available"]

    def test_says_so_when_solid_edge_ignores_the_change(self, query_mgr):
        qm, _doc = query_mgr
        mat_table = fake_material_table(current="Aluminum")
        mat_table.ApplyMaterialToDoc.side_effect = None  # accepts, changes nothing
        wire_material_table(qm, mat_table)

        result = qm.set_material("Steel")

        assert "error" in result
        assert "still reports" in result["error"]


class TestGetMaterialProperty:
    """Only GetMaterialPropValueFromDoc works, and the indices start at 3."""

    def test_success(self, query_mgr):
        qm, doc = query_mgr
        mat_table = fake_material_table(properties={23: 7750.0})
        wire_material_table(qm, mat_table)

        result = qm.get_material_property("Steel", 23)

        assert result["value"] == 7750.0
        assert result["property"] == "density"
        mat_table.GetMaterialPropValueFromDoc.assert_called_once_with(doc, 23)

    def test_applies_the_named_material_first(self, query_mgr):
        qm, _doc = query_mgr
        mat_table = fake_material_table(properties={28: 0.29})
        wire_material_table(qm, mat_table)

        qm.get_material_property("Aluminum", 28)

        mat_table.ApplyMaterialToDoc.assert_called_once()

    def test_an_empty_name_reads_what_is_already_applied(self, query_mgr):
        qm, _doc = query_mgr
        mat_table = fake_material_table(current="Steel", properties={23: 7750.0})
        wire_material_table(qm, mat_table)

        result = qm.get_material_property("", 23)

        assert result["material"] == "Steel"
        mat_table.ApplyMaterialToDoc.assert_not_called()

    def test_com_error(self, query_mgr):
        qm, _doc = query_mgr
        mat_table = fake_material_table(current="Steel")
        mat_table.GetMaterialPropValueFromDoc.side_effect = Exception("COM error")
        wire_material_table(qm, mat_table)

        assert "error" in qm.get_material_property("", 23)


# ============================================================================
# BATCH 10: MATERIAL LIBRARY
# ============================================================================


class TestGetMaterialLibrary:
    """MatTable has no Count/Item and there is no Material object at all."""

    def test_success(self, query_mgr):
        qm, _doc = query_mgr
        wire_material_table(
            qm,
            fake_material_table(
                current="Steel",
                properties={
                    23: 7750.0,
                    27: 2.0e11,
                    28: 0.29,
                    24: 1.2e-05,
                    25: 50.0,
                    26: 500.0,
                    29: 3.1e08,
                    30: 6.4e08,
                    31: 0.0,
                },
            ),
        )

        result = qm.get_material_library()

        assert result["material"] == "Steel"
        assert result["properties"]["density"] == 7750.0
        assert result["properties"]["poisson_ratio"] == 0.29

    def test_no_material_applied(self, query_mgr):
        qm, _doc = query_mgr
        mat_table = fake_material_table(current="")
        mat_table.GetMaterialPropValueFromDoc.side_effect = Exception("nothing applied")
        wire_material_table(qm, mat_table)

        result = qm.get_material_library()

        assert "error" in result
        assert "No material is applied" in result["error"]

    def test_com_error(self, query_mgr):
        qm, _doc = query_mgr
        app = MagicMock()
        app.GetMaterialTable.side_effect = Exception("COM error")
        qm.doc_manager.connection = MagicMock()
        qm.doc_manager.connection.get_application.return_value = app

        assert "error" in qm.get_material_library()


class TestSetMaterialByName:
    """Document.ApplyStyle and Material.Apply are not Solid Edge methods."""

    def test_success(self, query_mgr):
        qm, doc = query_mgr
        mat_table = fake_material_table(properties={23: 7750.0})
        wire_material_table(qm, mat_table)

        result = qm.set_material_by_name("Steel")

        assert result["status"] == "applied"
        assert result["density"] == 7750.0
        mat_table.ApplyMaterialToDoc.assert_called_once_with(doc, "Steel", "Materials")
        doc.ApplyStyle.assert_not_called()

    def test_not_found(self, query_mgr):
        qm, _doc = query_mgr
        wire_material_table(qm, fake_material_table())

        result = qm.set_material_by_name("Unobtainium")

        assert "error" in result
        assert "no installed library" in result["error"]


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
        del doc.ActiveSheet

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
        del doc.ActiveSheet

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
        del doc.ActiveSheet

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
        del doc.ActiveSheet

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
        del doc.ActiveSheet

        result = qm.delete_layer("Layer1")
        assert "error" in result


# ============================================================================
# LAYERS: WHERE THEY LIVE
# ============================================================================


class TestLayersOnADraft:
    """A DraftDocument has no Layers of its own; a Sheet does.

    The gate here used to be hasattr(doc, "Layers"), the probe CLAUDE.md
    forbids, so every layer call on a draft answered "Active document does not
    support layers" -- for the document type where layers matter most.
    """

    def test_a_draft_uses_the_active_sheet(self, query_mgr):
        qm, doc = query_mgr
        del doc.Layers

        layer = MagicMock()
        layer.Name = "SheetLayer"
        layer.Show = True
        layer.Locatable = True
        layer.IsEmpty = False
        sheet_layers = MagicMock()
        sheet_layers.Count = 1
        sheet_layers.Item.return_value = layer
        doc.ActiveSheet.Layers = sheet_layers

        result = qm.get_layers()

        assert result["count"] == 1
        assert result["layers"][0]["name"] == "SheetLayer"

    def test_adding_to_a_draft_goes_to_the_sheet(self, query_mgr):
        qm, doc = query_mgr
        del doc.Layers
        sheet_layers = MagicMock()
        sheet_layers.Count = 2
        doc.ActiveSheet.Layers = sheet_layers

        result = qm.add_layer("ProbeLayer")

        assert result["status"] == "added"
        sheet_layers.Add.assert_called_once_with("ProbeLayer")

    def test_a_document_own_layers_win(self, query_mgr):
        """A part has its own; the sheet fallback must not shadow them."""
        qm, doc = query_mgr
        own = MagicMock()
        own.Count = 0
        doc.Layers = own

        qm.get_layers()

        doc.ActiveSheet.Layers.Item.assert_not_called()

    def test_neither_is_an_honest_error(self, query_mgr):
        qm, doc = query_mgr
        del doc.Layers
        del doc.ActiveSheet

        result = qm.get_layers()

        assert "error" in result
        assert "active sheet" in result["error"]
