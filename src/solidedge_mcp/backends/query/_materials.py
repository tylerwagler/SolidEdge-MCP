"""Material table and layer management operations."""

import contextlib
from typing import Any

from solidedge_mcp.backends.errors import error_result

from ..comutil import com_get
from ..constants import MatTablePropIndexConstants
from ..logging import get_logger
from ._base import QueryManagerBase

_logger = get_logger(__name__)


#: MatTable property indices worth naming in a result.
_PROPERTY_NAMES: dict[int, str] = {
    MatTablePropIndexConstants.seDensity: "density",
    MatTablePropIndexConstants.seCoefOfThermalExpansion: "thermal_expansion",
    MatTablePropIndexConstants.seThermalConductivity: "thermal_conductivity",
    MatTablePropIndexConstants.seSpecificHeat: "specific_heat",
    MatTablePropIndexConstants.seModulusElasticity: "modulus_of_elasticity",
    MatTablePropIndexConstants.sePoissonRatio: "poisson_ratio",
    MatTablePropIndexConstants.seYieldStress: "yield_stress",
    MatTablePropIndexConstants.seUltimateStress: "ultimate_stress",
    MatTablePropIndexConstants.seElongation: "elongation",
}


def _out_pair(result: Any) -> Any:
    """Pick the list out of a two-out-param return.

    GetMaterialLibraryList returns (names, count) and
    GetMaterialListFromLibrary returns (count, names), so take whichever
    element is not the integer.
    """
    if not isinstance(result, tuple) or len(result) != 2:
        return result
    first, second = result
    if isinstance(first, int) and not isinstance(second, int):
        return second
    return first


def _string_tuple(value: Any) -> list[str]:
    """A COM string array as a plain list of str."""
    if value is None:
        return []
    if isinstance(value, str):
        return [value]
    try:
        return [str(item) for item in value]
    except TypeError:
        return [str(value)]


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
                                material_vars[name] = str(com_get(var, "Formula", "N/A"))
                    except Exception:
                        continue
            except Exception:
                pass

            # Also try to get material name from properties
            try:
                props = com_get(doc, "Properties")
                if props is not None:
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
            return error_result(e)

    # =================================================================
    # MATERIAL HELPERS
    #
    # MatTable is not a collection: it has no Count and no Item, and there is
    # no Material object to read properties off. Everything goes through named
    # calls on the table itself.
    # =================================================================

    def _material_table(self) -> Any:
        """The application material table.

        GetMaterialTable is on Application, not on the document.
        """
        app = self.doc_manager.connection.get_application()
        return app.GetMaterialTable()

    def _current_material(self, mat_table: Any, doc: Any) -> str:
        """The name of the material applied to a document, or an empty string."""
        try:
            return str(mat_table.GetCurrentMaterialName(doc) or "")
        except Exception:
            return ""

    def _find_material_library(self, mat_table: Any, material_name: str) -> str | None:
        """Which library holds a material, searched case-insensitively."""
        try:
            libraries = _string_tuple(_out_pair(mat_table.GetMaterialLibraryList()))
        except Exception:
            return None
        wanted = material_name.casefold()
        for library in libraries:
            try:
                names = _string_tuple(_out_pair(mat_table.GetMaterialListFromLibrary(library)))
            except Exception:
                continue
            for name in names:
                if name.casefold() == wanted:
                    return library
        return None

    def _apply_material(
        self, material_name: str, report_properties: bool = False
    ) -> dict[str, Any]:
        """Apply a material by name and confirm the document took it."""
        try:
            doc = self.doc_manager.get_active_document()
            mat_table = self._material_table()

            library = self._find_material_library(mat_table, material_name)
            if library is None:
                available = self.get_material_list()
                return {
                    "error": f"Material '{material_name}' is in no installed library.",
                    "available": available.get("materials", [])[:20],
                }

            mat_table.ApplyMaterialToDoc(doc, material_name, library)

            applied = self._current_material(mat_table, doc)
            if applied and applied.casefold() != material_name.casefold():
                return {
                    "error": (
                        f"Solid Edge did not apply '{material_name}'; the document "
                        f"still reports '{applied}'."
                    )
                }

            result: dict[str, Any] = {
                "status": "applied",
                "material": applied or material_name,
                "library": library,
            }
            if report_properties:
                with contextlib.suppress(Exception):
                    result["density"] = mat_table.GetMaterialPropValueFromDoc(
                        doc, MatTablePropIndexConstants.seDensity
                    )
            return result
        except Exception as e:
            return error_result(e)

    def get_material_list(self) -> dict[str, Any]:
        """List every material Solid Edge can apply, by library.

        ``MatTable.GetMaterialList()`` raises E_FAIL on Solid Edge 2026. The
        pair that works is ``GetMaterialLibraryList()``, which returns
        (names, count), and ``GetMaterialListFromLibrary(library)``, which
        returns (count, names) -- the two out-params come back in opposite
        orders.

        Returns:
            Dict with the flat list of names, the per-library breakdown, and
            a count.
        """
        try:
            mat_table = self._material_table()
            libraries = _string_tuple(_out_pair(mat_table.GetMaterialLibraryList()))

            by_library: dict[str, list[str]] = {}
            materials: list[str] = []
            for library in libraries:
                try:
                    names = _string_tuple(_out_pair(mat_table.GetMaterialListFromLibrary(library)))
                except Exception as exc:
                    _logger.warning("Material library %r could not be read: %s", library, exc)
                    continue
                by_library[library] = names
                materials.extend(names)

            if not materials:
                return {
                    "error": (
                        "Solid Edge reported no materials. Check that a material "
                        "library is installed and reachable."
                    ),
                    "libraries": libraries,
                }

            return {
                "materials": materials,
                "count": len(materials),
                "libraries": by_library,
            }
        except Exception as e:
            return error_result(e)

    def set_material(self, material_name: str) -> dict[str, Any]:
        """Apply a named material to the active document.

        Args:
            material_name: A name from get_material_list.

        Returns:
            Dict with status and the name Solid Edge reports afterwards.
        """
        return self._apply_material(material_name)

    def get_material_property(self, material_name: str, property_index: int) -> dict[str, Any]:
        """Read one property of the material applied to the active document.

        ``MatTable.GetMatPropValue(name, index)`` raises E_FAIL on Solid Edge
        2026, and so does ``GetMaterialPropValueFromLibrary``. Only
        ``GetMaterialPropValueFromDoc(document, index)`` works, so the material
        has to be applied first; naming a different material applies it.

        The indices this used to document (0 = Density, 1 = Thermal
        Conductivity, ...) were invented. The real ones come from
        ``MatTablePropIndexConstants`` and start at 3.

        Args:
            material_name: The material to read. Empty means whatever is
                already applied.
            property_index: A MatTablePropIndexConstants value, for example 23
                for density or 28 for Poisson's ratio.

        Returns:
            Dict with the value and the property name, when it is a known one.
        """
        try:
            doc = self.doc_manager.get_active_document()
            mat_table = self._material_table()

            if material_name:
                applied = self._apply_material(material_name)
                if "error" in applied:
                    return applied

            value = mat_table.GetMaterialPropValueFromDoc(doc, property_index)
            result: dict[str, Any] = {
                "material": material_name or self._current_material(mat_table, doc),
                "property_index": property_index,
                "value": value,
            }
            name = _PROPERTY_NAMES.get(property_index)
            if name is not None:
                result["property"] = name
            return result
        except Exception as e:
            return error_result(e)

    def get_material_library(self) -> dict[str, Any]:
        """Report the material applied to the active document, in full.

        This used to walk ``mat_table.Count`` and ``mat_table.Item(i)`` reading
        ``Name``, ``Density``, ``YoungsModulus`` and ``PoissonsRatio``. MatTable
        has none of those, and there is no Material interface at all, so the
        call raised before it read anything. Solid Edge only reports material
        properties for the document, through
        ``GetMaterialPropValueFromDoc``, so this reports the applied material.
        Use get_material_list for what is available.

        Returns:
            Dict with the material name and every property Solid Edge gives.
        """
        try:
            doc = self.doc_manager.get_active_document()
            mat_table = self._material_table()

            name = self._current_material(mat_table, doc)
            properties: dict[str, Any] = {}
            for index, label in _PROPERTY_NAMES.items():
                try:
                    properties[label] = mat_table.GetMaterialPropValueFromDoc(doc, index)
                except Exception:
                    continue

            if not name and not properties:
                return {
                    "error": (
                        "No material is applied to the active document, so it has "
                        "no properties to report. Apply one with "
                        "manage_material(action='set', material_name=...)."
                    )
                }

            return {"material": name, "properties": properties}
        except Exception as e:
            return error_result(e)

    def set_material_by_name(self, material_name: str) -> dict[str, Any]:
        """Apply a material, checking the name against the libraries first.

        ``Document.ApplyStyle`` is not a Solid Edge method, and neither is
        ``Material.Apply``, so the old three-way fallback could only ever land
        on ``doc.Material = name``, which does not apply a material either.
        ``MatTable.ApplyMaterialToDoc(document, name, library)`` is the real
        route.

        Args:
            material_name: A name from get_material_list.

        Returns:
            Dict with status, the library it came from, and its density.
        """
        return self._apply_material(material_name, report_properties=True)

    # =================================================================
    # LAYER MANAGEMENT
    # =================================================================

    @staticmethod
    def _layers_of(doc: Any) -> tuple[Any, dict[str, Any] | None]:
        """The Layers collection for this document, wherever it lives.

        A DraftDocument has no ``Layers`` of its own -- the type library gives
        it to PartDocument, SheetMetalDocument, WeldmentDocument and
        AssemblyDocument, and to ``Sheet``. Draft layers are per sheet, so they
        come from the active sheet.

        This used to gate on ``hasattr(doc, "Layers")``, which is the probe
        CLAUDE.md forbids: it reads False both for a member that is genuinely
        absent and for one whose getter merely raised, and either way every
        layer call on a draft answered "Active document does not support
        layers" -- for a document type whose layers are a core drafting
        feature.

        Returns ``(layers, error_dict)``.
        """
        layers = com_get(doc, "Layers")
        if layers is not None:
            return layers, None

        sheet = com_get(doc, "ActiveSheet")
        layers = com_get(sheet, "Layers")
        if layers is not None:
            return layers, None

        return None, {
            "error": (
                "This document has no layers. Parts, sheet metal, weldments and "
                "assemblies carry them on the document; a draft carries them on "
                "its active sheet."
            )
        }

    def get_layers(self) -> dict[str, Any]:
        """
        Get all layers in the active document.

        Iterates doc.Layers and reports Name, Show, Locatable, IsEmpty.

        Returns:
            Dict with list of layers
        """
        try:
            doc = self.doc_manager.get_active_document()

            layers_col, err = self._layers_of(doc)
            if err:
                return err
            layers = []

            for i in range(1, layers_col.Count + 1):
                try:
                    layer = layers_col.Item(i)
                    info: dict[str, Any] = {"index": i - 1}
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
            return error_result(e)

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

            layers, err = self._layers_of(doc)
            if err:
                return err
            layers.Add(name)

            return {"status": "added", "name": name, "total_layers": layers.Count}
        except Exception as e:
            return error_result(e)

    def activate_layer(self, name_or_index: str | int) -> dict[str, Any]:
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

            layers, err = self._layers_of(doc)
            if err:
                return err

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
            layer_name = com_get(layer, "Name", str(name_or_index))

            return {"status": "activated", "name": layer_name}
        except Exception as e:
            return error_result(e)

    def set_layer_properties(
        self, name_or_index: str | int, show: bool | None = None, selectable: bool | None = None
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

            layers, err = self._layers_of(doc)
            if err:
                return err

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

            layer_name = com_get(layer, "Name", str(name_or_index))

            return {"status": "updated", "name": layer_name, "properties": updated}
        except Exception as e:
            return error_result(e)

    def delete_layer(self, name_or_index: str | int) -> dict[str, Any]:
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

            layers, err = self._layers_of(doc)
            if err:
                return err

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

            layer_name = com_get(layer, "Name", str(name_or_index))
            layer.Delete()

            return {"status": "deleted", "name": layer_name}
        except Exception as e:
            return error_result(e)
