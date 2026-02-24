"""Material table and layer management operations."""

import contextlib
import traceback
from typing import Any, Optional

from ..logging import get_logger
from ._base import QueryManagerBase

_logger = get_logger(__name__)


class MaterialsMixin(QueryManagerBase):
    """Mixin providing material and layer management methods."""

    doc_manager: Any

    def get_material_table(self) -> dict[str, Any]:
        """
        Get the available material properties from the document variables.

        Returns material-related variables from the active document.

        Returns:
            Dict with material properties
        """
        try:
            doc = self.doc_manager.get_active_document()

            material_vars = {}
            material_names = [
                "Density",
                "Mass",
                "Volume",
                "Surface_Area",
                "YoungsModulus",
                "PoissonsRatio",
                "ThermalExpansionCoefficient",
                "ThermalConductivity",
                "Material",
            ]

            try:
                variables = doc.Variables
                for i in range(1, variables.Count + 1):
                    try:
                        var = variables.Item(i)
                        name = var.Name
                        if name in material_names or name.startswith("Material"):
                            try:
                                material_vars[name] = var.Value
                            except Exception:
                                material_vars[name] = (
                                    str(var.Formula) if hasattr(var, "Formula") else "N/A"
                                )
                    except Exception:
                        continue
            except Exception:
                pass

            # Also try to get material name from properties
            try:
                if hasattr(doc, "Properties"):
                    props = doc.Properties
                    for i in range(1, props.Count + 1):
                        try:
                            prop_set = props.Item(i)
                            for j in range(1, prop_set.Count + 1):
                                try:
                                    prop = prop_set.Item(j)
                                    if "material" in prop.Name.lower():
                                        material_vars[prop.Name] = prop.Value
                                except Exception:
                                    continue
                        except Exception:
                            continue
            except Exception:
                pass

            return {"material_properties": material_vars, "property_count": len(material_vars)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_material_list(self) -> dict[str, Any]:
        """
        Get the list of available materials from the material table.

        Uses MatTable.GetMaterialList() via the Solid Edge application object.

        Returns:
            Dict with list of material names
        """
        try:
            self.doc_manager.get_active_document()

            # Get MatTable from the application
            app = self.doc_manager.connection.get_application()
            mat_table = app.GetMaterialTable()

            result = mat_table.GetMaterialList()
            # Returns (count, list_of_materials) as out-params
            if isinstance(result, tuple) and len(result) >= 2:
                result[0]
                material_list = result[1]
            else:
                material_list = result

            materials = []
            if material_list is not None:
                if isinstance(material_list, (list, tuple)):
                    materials = list(material_list)
                elif hasattr(material_list, "__iter__"):
                    materials = [str(m) for m in material_list]

            return {"materials": materials, "count": len(materials)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_material(self, material_name: str) -> dict[str, Any]:
        """
        Apply a named material to the active document.

        Uses MatTable.ApplyMaterial(pDocument, bstrMatName).

        Args:
            material_name: Name of the material (from get_material_list)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            app = self.doc_manager.connection.get_application()
            mat_table = app.GetMaterialTable()

            mat_table.ApplyMaterial(doc, material_name)

            return {"status": "applied", "material": material_name}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_material_property(self, material_name: str, property_index: int) -> dict[str, Any]:
        """
        Get a specific property value for a material.

        Uses MatTable.GetMatPropValue(bstrMatName, lPropIndex, varPropValue).

        Common property indices:
            0 = Density, 1 = Thermal Conductivity, 2 = Thermal Expansion,
            3 = Specific Heat, 4 = Young's Modulus, 5 = Poisson's Ratio,
            6 = Yield Stress, 7 = Ultimate Stress, 8 = Elongation

        Args:
            material_name: Name of the material
            property_index: Property index (see above)

        Returns:
            Dict with property value
        """
        try:
            app = self.doc_manager.connection.get_application()
            mat_table = app.GetMaterialTable()

            value = mat_table.GetMatPropValue(material_name, property_index)

            return {"material": material_name, "property_index": property_index, "value": value}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_material_library(self) -> dict[str, Any]:
        """
        Get the full material library from the active document.

        Iterates the material table to list all available materials
        with their names and density values.

        Returns:
            Dict with count and list of material info dicts
        """
        try:
            doc = self.doc_manager.get_active_document()

            mat_table = doc.GetMaterialTable()
            materials = []
            for i in range(1, mat_table.Count + 1):
                mat = mat_table.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    info["name"] = mat.Name
                with contextlib.suppress(Exception):
                    info["density"] = mat.Density
                with contextlib.suppress(Exception):
                    info["youngs_modulus"] = mat.YoungsModulus
                with contextlib.suppress(Exception):
                    info["poissons_ratio"] = mat.PoissonsRatio
                materials.append(info)
            return {"count": len(materials), "materials": materials}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_material_by_name(self, material_name: str) -> dict[str, Any]:
        """
        Look up a material by name in the material table and apply it.

        Searches the material table for the given name and applies it
        to the active document.

        Args:
            material_name: Name of the material to apply

        Returns:
            Dict with status and material info
        """
        try:
            doc = self.doc_manager.get_active_document()

            mat_table = doc.GetMaterialTable()

            # Search for the material by name
            found_mat = None
            for i in range(1, mat_table.Count + 1):
                mat = mat_table.Item(i)
                try:
                    if mat.Name == material_name:
                        found_mat = mat
                        break
                except Exception:
                    continue

            if found_mat is None:
                # Collect available names for error message
                available = []
                for i in range(1, mat_table.Count + 1):
                    with contextlib.suppress(Exception):
                        available.append(mat_table.Item(i).Name)
                return {
                    "error": f"Material '{material_name}' not found",
                    "available": available[:20],
                }

            # Apply the material
            try:
                doc.ApplyStyle(found_mat)
            except Exception:
                # Alternative: try setting via material name property
                try:
                    doc.Material = material_name
                except Exception:
                    # Another fallback
                    found_mat.Apply()

            result: dict[str, Any] = {"status": "applied", "material": material_name}
            with contextlib.suppress(Exception):
                result["density"] = found_mat.Density
            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # LAYER MANAGEMENT
    # =================================================================

    def get_layers(self) -> dict[str, Any]:
        """
        Get all layers in the active document.

        Iterates doc.Layers and reports Name, Show, Locatable, IsEmpty.

        Returns:
            Dict with list of layers
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Layers"):
                return {"error": "Active document does not support layers"}

            layers_col = doc.Layers
            layers = []

            for i in range(1, layers_col.Count + 1):
                try:
                    layer = layers_col.Item(i)
                    info = {"index": i - 1}
                    try:
                        info["name"] = layer.Name
                    except Exception:
                        info["name"] = f"Layer_{i}"
                    with contextlib.suppress(Exception):
                        info["show"] = bool(layer.Show)
                    with contextlib.suppress(Exception):
                        info["locatable"] = bool(layer.Locatable)
                    with contextlib.suppress(Exception):
                        info["is_empty"] = bool(layer.IsEmpty)
                    layers.append(info)
                except Exception:
                    continue

            return {"layers": layers, "count": len(layers)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_layer(self, name: str) -> dict[str, Any]:
        """
        Add a new layer to the active document.

        Uses doc.Layers.Add(name).

        Args:
            name: Name for the new layer

        Returns:
            Dict with status and layer info
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Layers"):
                return {"error": "Active document does not support layers"}

            layers = doc.Layers
            layers.Add(name)

            return {"status": "added", "name": name, "total_layers": layers.Count}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def activate_layer(self, name_or_index) -> dict[str, Any]:
        """
        Activate a layer by name or 0-based index.

        Finds the layer and calls layer.Activate().

        Args:
            name_or_index: Layer name (str) or 0-based index (int)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Layers"):
                return {"error": "Active document does not support layers"}

            layers = doc.Layers

            if isinstance(name_or_index, int):
                if name_or_index < 0 or name_or_index >= layers.Count:
                    return {"error": f"Invalid layer index: {name_or_index}. Count: {layers.Count}"}
                layer = layers.Item(name_or_index + 1)
            else:
                # Search by name
                layer = None
                for i in range(1, layers.Count + 1):
                    try:
                        lyr = layers.Item(i)
                        if lyr.Name == name_or_index:
                            layer = lyr
                            break
                    except Exception:
                        continue
                if layer is None:
                    return {"error": f"Layer '{name_or_index}' not found"}

            layer.Activate()
            layer_name = layer.Name if hasattr(layer, "Name") else str(name_or_index)

            return {"status": "activated", "name": layer_name}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_layer_properties(
        self, name_or_index, show: bool = None, selectable: bool = None
    ) -> dict[str, Any]:
        """
        Set layer visibility and selectability properties.

        Args:
            name_or_index: Layer name (str) or 0-based index (int)
            show: If provided, set layer visibility
            selectable: If provided, set layer selectability (Locatable)

        Returns:
            Dict with status and updated properties
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Layers"):
                return {"error": "Active document does not support layers"}

            layers = doc.Layers

            if isinstance(name_or_index, int):
                if name_or_index < 0 or name_or_index >= layers.Count:
                    return {"error": f"Invalid layer index: {name_or_index}. Count: {layers.Count}"}
                layer = layers.Item(name_or_index + 1)
            else:
                layer = None
                for i in range(1, layers.Count + 1):
                    try:
                        lyr = layers.Item(i)
                        if lyr.Name == name_or_index:
                            layer = lyr
                            break
                    except Exception:
                        continue
                if layer is None:
                    return {"error": f"Layer '{name_or_index}' not found"}

            updated = {}
            if show is not None:
                layer.Show = show
                updated["show"] = show
            if selectable is not None:
                layer.Locatable = selectable
                updated["selectable"] = selectable

            layer_name = layer.Name if hasattr(layer, "Name") else str(name_or_index)

            return {"status": "updated", "name": layer_name, "properties": updated}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def delete_layer(self, name_or_index) -> dict[str, Any]:
        """
        Delete a layer from the active document.

        Cannot delete the active layer or a layer containing objects.

        Args:
            name_or_index: Layer name (str) or 0-based index (int)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Layers"):
                return {"error": "Active document does not support layers"}

            layers = doc.Layers

            if isinstance(name_or_index, int):
                if name_or_index < 0 or name_or_index >= layers.Count:
                    return {"error": f"Invalid layer index: {name_or_index}. Count: {layers.Count}"}
                layer = layers.Item(name_or_index + 1)
            else:
                layer = None
                for i in range(1, layers.Count + 1):
                    try:
                        lyr = layers.Item(i)
                        if lyr.Name == name_or_index:
                            layer = lyr
                            break
                    except Exception:
                        continue
                if layer is None:
                    return {"error": f"Layer '{name_or_index}' not found"}

            layer_name = layer.Name if hasattr(layer, "Name") else str(name_or_index)
            layer.Delete()

            return {"status": "deleted", "name": layer_name}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
