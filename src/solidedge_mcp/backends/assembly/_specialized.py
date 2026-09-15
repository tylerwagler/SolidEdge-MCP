"""Specialized assembly subsystem operations (virtual components, tubes, frames, wiring)."""

import contextlib
import os
from typing import Any

from solidedge_mcp.backends.errors import error_result

from ..comutil import com_get
from ..logging import get_logger

_logger = get_logger(__name__)


#: StructuralFrames.Add and AddByOrientation both take a Path array of the
#: curves the frame runs along -- 3D sketch segments drawn in the assembly.
#: Both methods were being handed Occurrences instead, which Solid Edge 2026
#: rejects with E_POINTER, "Property name is invalid". There is no way to
#: build such a path from here: an AssemblyDocument exposes no Sketches
#: collection through this binding, and the Segments commands that would
#: create one raise a modal dialog, which blocks the whole server.
_NO_FRAME_PATH: dict[str, Any] = {
    "error": (
        "A structural frame runs along 3D sketch segments drawn in the assembly, "
        "and StructuralFrames.Add takes those curves as its Path. This server "
        "cannot draw or select them -- passing occurrences instead is rejected "
        "with E_POINTER. Create the frame in the Solid Edge UI."
    ),
    "unsupported": True,
}


class SpecializedMixin:
    """Mixin providing specialized assembly subsystem methods."""

    # -- Virtual Components --------------------------------------------------

    def _active_harness(self, doc: Any) -> tuple[Any, dict[str, Any] | None]:
        """The assembly's harness, creating one when there is none.

        Wires, cables, splices and bundles hang off a Harness, which hangs off
        AssemblyDocument.Harnesses. Reading them straight from the document
        raised every time.
        """
        harnesses = com_get(doc, "Harnesses")
        if harnesses is None:
            return None, {
                "error": (
                    "This assembly has no Harnesses collection, so wiring cannot "
                    "be added. Harness work needs Solid Edge Wire Harness Design."
                )
            }
        try:
            if int(com_get(harnesses, "Count", 0) or 0):
                return harnesses.Item(1), None
            return harnesses.Add(), None
        except Exception as exc:
            return None, error_result(exc, context="Could not open a wire harness")

    def add_virtual_component(
        self,
        name: str,
        component_type: str = "Part",
    ) -> dict[str, Any]:
        """
        Add a virtual component to the assembly.

        Virtual components are placeholders for parts that don't yet have
        geometry files. Uses VirtualComponentOccurrences.Add.

        Args:
            name: Name for the virtual component
            component_type: 'Part' (3), 'Assembly' (2), 'Sheetmetal' (4),
                or 'Unknown' (1)

        Returns:
            Dict with status and virtual component info
        """
        try:
            _logger.info(f"Adding virtual component: name={name}, type={component_type}")
            doc = self.doc_manager.get_active_document()

            err = self._require_assembly(doc)
            if err:
                return err

            type_map = {
                "Unknown": 1,
                "Assembly": 2,
                "Part": 3,
                "Sheetmetal": 4,
            }
            vc_type = type_map.get(component_type, 3)

            vc_occs = doc.VirtualComponentOccurrences
            vc_occ = vc_occs.Add(name, vc_type)

            result: dict[str, Any] = {
                "status": "created",
                "name": name,
                "component_type": component_type,
            }

            with contextlib.suppress(Exception):
                result["vc_name"] = vc_occ.Name

            return result
        except Exception as e:
            _logger.error(f"Failed to add virtual component: {e}")
            return error_result(e)

    def add_virtual_component_predefined(
        self,
        filename: str,
    ) -> dict[str, Any]:
        """
        Add a pre-defined virtual component from a file.

        Uses VirtualComponentOccurrences.AddAsPreDefined.

        Args:
            filename: Path to the virtual component file

        Returns:
            Dict with status and virtual component info
        """
        try:
            _logger.info(f"Adding predefined virtual component: {filename}")
            if not os.path.exists(filename):
                return {"error": f"File not found: {filename}"}

            doc = self.doc_manager.get_active_document()

            err = self._require_assembly(doc)
            if err:
                return err

            vc_occs = doc.VirtualComponentOccurrences
            vc_occ = vc_occs.AddAsPreDefined(filename)

            result: dict[str, Any] = {
                "status": "created",
                "filename": filename,
            }

            with contextlib.suppress(Exception):
                result["vc_name"] = vc_occ.Name

            return result
        except Exception as e:
            _logger.error(f"Failed to add predefined virtual component: {e}")
            return error_result(e)

    def add_virtual_component_bidm(
        self,
        doc_number: str,
        revision_id: str,
        component_type: str = "Part",
    ) -> dict[str, Any]:
        """
        Add a virtual component using BIDM (document number + revision).

        Uses VirtualComponentOccurrences.AddBIDM.

        Args:
            doc_number: Document number string
            revision_id: Revision ID string
            component_type: 'Part' (3), 'Assembly' (2), 'Sheetmetal' (4),
                or 'Unknown' (1)

        Returns:
            Dict with status and virtual component info
        """
        try:
            _logger.info(f"Adding virtual component via BIDM: doc={doc_number}, rev={revision_id}")
            doc = self.doc_manager.get_active_document()

            err = self._require_assembly(doc)
            if err:
                return err

            type_map = {
                "Unknown": 1,
                "Assembly": 2,
                "Part": 3,
                "Sheetmetal": 4,
            }
            vc_type = type_map.get(component_type, 3)

            vc_occs = doc.VirtualComponentOccurrences
            vc_occ = vc_occs.AddBIDM(doc_number, revision_id, vc_type)

            result: dict[str, Any] = {
                "status": "created",
                "doc_number": doc_number,
                "revision_id": revision_id,
                "component_type": component_type,
            }

            with contextlib.suppress(Exception):
                result["vc_name"] = vc_occ.Name

            return result
        except Exception as e:
            _logger.error(f"Failed to add virtual component via BIDM: {e}")
            return error_result(e)

    # -- Tube Operations -----------------------------------------------------

    def get_tube(self, component_index: int) -> dict[str, Any]:
        """
        Get tube information from a tube occurrence.

        Uses Occurrence.GetTube() to retrieve tube properties.

        Args:
            component_index: 0-based index of the tube component

        Returns:
            Dict with tube properties (outer diameter, wall thickness, etc.)
        """
        try:
            _logger.info(f"Getting tube info: component_index={component_index}")
            doc = self.doc_manager.get_active_document()
            occurrences, occurrence, err = self._validate_occurrence_index(doc, component_index)
            if err:
                return err

            tube = occurrence.GetTube()
            if tube is None:
                return {
                    "error": f"Component at index {component_index} "
                    "is not a tube or has no tube data"
                }

            result: dict[str, Any] = {
                "component_index": component_index,
                "is_tube": True,
            }

            with contextlib.suppress(Exception):
                result["outer_diameter"] = tube.OuterDiameter
            with contextlib.suppress(Exception):
                result["wall_thickness"] = tube.WallThickness
            with contextlib.suppress(Exception):
                result["bend_radius"] = tube.BendRadius
            with contextlib.suppress(Exception):
                result["material"] = tube.Material
            with contextlib.suppress(Exception):
                result["is_solid"] = tube.IsSolid
            with contextlib.suppress(Exception):
                result["length"] = tube.Length

            return result
        except Exception as e:
            _logger.error(f"Failed to get tube info: {e}")
            return error_result(e)

    def add_tube(
        self,
        segment_indices: list[int],
        part_filename: str,
        outer_diameter: float = 0.0,
        wall_thickness: float = 0.0,
        bend_radius: float = 0.0,
        is_solid: bool = False,
    ) -> dict[str, Any]:
        """
        Add a tube occurrence to the assembly.

        Uses Occurrences.AddTube with VARIANT-wrapped segment array.

        Args:
            segment_indices: 0-based indices of occurrences that define
                tube path segments
            part_filename: Path to the tube part file template
            outer_diameter: Outer diameter in meters (0 = use template)
            wall_thickness: Wall thickness in meters (0 = use template)
            bend_radius: Bend radius in meters (0 = use template)
            is_solid: True for solid tube, False for hollow

        Returns:
            Dict with status and tube occurrence info
        """
        try:
            _logger.info(f"Adding tube: part={part_filename}, segments={len(segment_indices)}")
            if not os.path.exists(part_filename):
                return {"error": f"File not found: {part_filename}"}

            doc = self.doc_manager.get_active_document()

            err = self._require_assembly(doc)
            if err:
                return err

            occurrences = doc.Occurrences

            # Resolve segment indices to occurrence objects
            segments = []
            for idx in segment_indices:
                if idx < 0 or idx >= occurrences.Count:
                    return {"error": f"Invalid segment index: {idx}. Count: {occurrences.Count}"}
                segments.append(occurrences.Item(idx + 1))

            v_segments = segments

            occ = occurrences.AddTube(
                v_segments,
                part_filename,
                None,  # TemplateFileName (optional)
                is_solid,  # IsSolid
                None,  # Material
                bend_radius if bend_radius > 0 else None,
                outer_diameter if outer_diameter > 0 else None,
                None,  # MinimumFlatLength
                wall_thickness if wall_thickness > 0 else None,
            )

            result: dict[str, Any] = {
                "status": "created",
                "type": "tube",
                "part_filename": part_filename,
                "num_segments": len(segments),
            }

            with contextlib.suppress(Exception):
                result["name"] = occ.Name

            return result
        except Exception as e:
            _logger.error(f"Failed to add tube: {e}")
            return error_result(e)

    # -- Structural Frames ---------------------------------------------------

    def add_structural_frame(
        self,
        part_filename: str,
        path_indices: list[int],
    ) -> dict[str, Any]:
        """
        Structural frames cannot be created through COM automation here.

        See ``_NO_FRAME_PATH``: the Path array wants the 3D sketch segments the
        frame runs along, and this server can neither draw nor select them.

        Args:
            part_filename: Path to the frame cross-section part file, unused.
            path_indices: Unused; occurrences are not a valid path.

        Returns:
            Dict with an ``unsupported`` error.
        """
        del part_filename, path_indices
        return _NO_FRAME_PATH

    def add_structural_frame_by_orientation(
        self,
        part_filename: str,
        coord_system_name: str,
        path_indices: list[int],
    ) -> dict[str, Any]:
        """
        Structural frames cannot be created through COM automation here.

        Same reason as ``add_structural_frame``: see ``_NO_FRAME_PATH``.

        Args:
            part_filename: Unused.
            coord_system_name: Unused.
            path_indices: Unused.

        Returns:
            Dict with an ``unsupported`` error.
        """
        del part_filename, coord_system_name, path_indices
        return _NO_FRAME_PATH

    def add_splice(
        self,
        x: float,
        y: float,
        z: float,
        conductor_indices: list[int],
        description: str = "",
    ) -> dict[str, Any]:
        """
        Add a splice to the assembly wiring harness.

        Uses Splices.Add with VARIANT-wrapped conductor array.

        Args:
            x: Splice X position in meters
            y: Splice Y position in meters
            z: Splice Z position in meters
            conductor_indices: 0-based indices of conductor occurrences
            description: Description text for the splice

        Returns:
            Dict with status and splice info
        """
        try:
            _logger.info(f"Adding splice at ({x},{y},{z}) with {len(conductor_indices)} conductors")
            doc = self.doc_manager.get_active_document()

            err = self._require_assembly(doc)
            if err:
                return err

            occurrences = doc.Occurrences

            conductors = []
            for idx in conductor_indices:
                if idx < 0 or idx >= occurrences.Count:
                    return {"error": f"Invalid conductor index: {idx}. Count: {occurrences.Count}"}
                conductors.append(occurrences.Item(idx + 1))

            v_conductors = conductors

            # Splices belongs to a Harness, not to the document; reading it
            # off the document always raised. A harness is created on
            # demand so the first wire in an assembly has somewhere to go.
            harness, err = self._active_harness(doc)
            if err:
                return err
            splices = harness.Splices
            splice = splices.Add(
                x,
                y,
                z,
                len(conductors),
                v_conductors,
                description,
            )

            result: dict[str, Any] = {
                "status": "created",
                "type": "splice",
                "position": [x, y, z],
                "num_conductors": len(conductors),
                "description": description,
            }

            with contextlib.suppress(Exception):
                result["name"] = splice.Name

            return result
        except Exception as e:
            _logger.error(f"Failed to add splice: {e}")
            return error_result(e)

    # -- Wires ---------------------------------------------------------------

    def add_wire(
        self,
        path_indices: list[int],
        path_directions: list[bool],
        description: str = "",
    ) -> dict[str, Any]:
        """
        Add a wire to the assembly wiring harness.

        Uses Wires.Add with VARIANT-wrapped path and direction arrays.

        Args:
            path_indices: 0-based indices of occurrences defining the wire path
            path_directions: Direction booleans for each path segment
                (True=forward, False=reverse)
            description: Description text for the wire

        Returns:
            Dict with status and wire info
        """
        try:
            _logger.info(f"Adding wire with {len(path_indices)} path segments")
            doc = self.doc_manager.get_active_document()

            err = self._require_assembly(doc)
            if err:
                return err

            if len(path_indices) != len(path_directions):
                return {"error": "path_indices and path_directions must have the same length"}

            occurrences = doc.Occurrences

            paths = []
            for idx in path_indices:
                if idx < 0 or idx >= occurrences.Count:
                    return {"error": f"Invalid path index: {idx}. Count: {occurrences.Count}"}
                paths.append(occurrences.Item(idx + 1))

            v_paths = paths
            v_dirs = path_directions

            # Wires belongs to a Harness, not to the document; reading it
            # off the document always raised. A harness is created on
            # demand so the first wire in an assembly has somewhere to go.
            harness, err = self._active_harness(doc)
            if err:
                return err
            wires = harness.Wires
            wire = wires.Add(len(paths), v_paths, v_dirs, description)

            result: dict[str, Any] = {
                "status": "created",
                "type": "wire",
                "num_paths": len(paths),
                "description": description,
            }

            with contextlib.suppress(Exception):
                result["name"] = wire.Name

            return result
        except Exception as e:
            _logger.error(f"Failed to add wire: {e}")
            return error_result(e)

    # -- Cables --------------------------------------------------------------

    def add_cable(
        self,
        path_indices: list[int],
        path_directions: list[bool],
        wire_indices: list[int],
        split_path_indices: list[int] | None = None,
        split_path_directions: list[bool] | None = None,
        description: str = "",
    ) -> dict[str, Any]:
        """
        Add a cable to the assembly wiring harness.

        Uses Cables.Add with multiple VARIANT-wrapped arrays.

        Args:
            path_indices: 0-based indices of occurrences defining cable path
            path_directions: Direction booleans for each path segment
            wire_indices: 0-based indices of wire occurrences in the cable
            split_path_indices: Optional split path occurrence indices
            split_path_directions: Optional split path direction booleans
            description: Description text for the cable

        Returns:
            Dict with status and cable info
        """
        try:
            _logger.info(f"Adding cable with {len(path_indices)} paths, {len(wire_indices)} wires")
            doc = self.doc_manager.get_active_document()

            err = self._require_assembly(doc)
            if err:
                return err

            if len(path_indices) != len(path_directions):
                return {"error": "path_indices and path_directions must have the same length"}

            occurrences = doc.Occurrences

            paths = []
            for idx in path_indices:
                if idx < 0 or idx >= occurrences.Count:
                    return {"error": f"Invalid path index: {idx}. Count: {occurrences.Count}"}
                paths.append(occurrences.Item(idx + 1))

            wires_list = []
            for idx in wire_indices:
                if idx < 0 or idx >= occurrences.Count:
                    return {"error": f"Invalid wire index: {idx}. Count: {occurrences.Count}"}
                wires_list.append(occurrences.Item(idx + 1))

            split_paths = []
            split_dirs = split_path_directions or []
            for idx in split_path_indices or []:
                if idx < 0 or idx >= occurrences.Count:
                    return {"error": f"Invalid split path index: {idx}. Count: {occurrences.Count}"}
                split_paths.append(occurrences.Item(idx + 1))

            v_paths = paths
            v_dirs = path_directions
            v_wires = wires_list
            v_split_paths = split_paths
            v_split_dirs = split_dirs

            # Cables belongs to a Harness, not to the document; reading it
            # off the document always raised. A harness is created on
            # demand so the first wire in an assembly has somewhere to go.
            harness, err = self._active_harness(doc)
            if err:
                return err
            cables = harness.Cables
            cable = cables.Add(
                len(paths),
                v_paths,
                v_dirs,
                len(wires_list),
                v_wires,
                v_split_paths,
                v_split_dirs,
                description,
            )

            result: dict[str, Any] = {
                "status": "created",
                "type": "cable",
                "num_paths": len(paths),
                "num_wires": len(wires_list),
                "description": description,
            }

            with contextlib.suppress(Exception):
                result["name"] = cable.Name

            return result
        except Exception as e:
            _logger.error(f"Failed to add cable: {e}")
            return error_result(e)

    # -- Bundles -------------------------------------------------------------

    def add_bundle(
        self,
        path_indices: list[int],
        path_directions: list[bool],
        conductor_indices: list[int],
        split_path_indices: list[int] | None = None,
        split_path_directions: list[bool] | None = None,
        description: str = "",
    ) -> dict[str, Any]:
        """
        Add a bundle to the assembly wiring harness.

        Uses Bundles.Add with multiple VARIANT-wrapped arrays.

        Args:
            path_indices: 0-based indices of occurrences defining bundle path
            path_directions: Direction booleans for each path segment
            conductor_indices: 0-based indices of conductor occurrences
            split_path_indices: Optional split path occurrence indices
            split_path_directions: Optional split path direction booleans
            description: Description text for the bundle

        Returns:
            Dict with status and bundle info
        """
        try:
            _logger.info(
                "Adding bundle with %d paths, %d conductors",
                len(path_indices),
                len(conductor_indices),
            )
            doc = self.doc_manager.get_active_document()

            err = self._require_assembly(doc)
            if err:
                return err

            if len(path_indices) != len(path_directions):
                return {"error": "path_indices and path_directions must have the same length"}

            occurrences = doc.Occurrences

            paths = []
            for idx in path_indices:
                if idx < 0 or idx >= occurrences.Count:
                    return {"error": f"Invalid path index: {idx}. Count: {occurrences.Count}"}
                paths.append(occurrences.Item(idx + 1))

            conductors = []
            for idx in conductor_indices:
                if idx < 0 or idx >= occurrences.Count:
                    return {"error": f"Invalid conductor index: {idx}. Count: {occurrences.Count}"}
                conductors.append(occurrences.Item(idx + 1))

            split_paths = []
            split_dirs = split_path_directions or []
            for idx in split_path_indices or []:
                if idx < 0 or idx >= occurrences.Count:
                    return {"error": f"Invalid split path index: {idx}. Count: {occurrences.Count}"}
                split_paths.append(occurrences.Item(idx + 1))

            v_paths = paths
            v_dirs = path_directions
            v_conductors = conductors
            v_split_paths = split_paths
            v_split_dirs = split_dirs

            # Bundles belongs to a Harness, not to the document; reading it
            # off the document always raised. A harness is created on
            # demand so the first wire in an assembly has somewhere to go.
            harness, err = self._active_harness(doc)
            if err:
                return err
            bundles = harness.Bundles
            bundle = bundles.Add(
                len(paths),
                v_paths,
                v_dirs,
                len(conductors),
                v_conductors,
                v_split_paths,
                v_split_dirs,
                description,
            )

            result: dict[str, Any] = {
                "status": "created",
                "type": "bundle",
                "num_paths": len(paths),
                "num_conductors": len(conductors),
                "description": description,
            }

            with contextlib.suppress(Exception):
                result["name"] = bundle.Name

            return result
        except Exception as e:
            _logger.error(f"Failed to add bundle: {e}")
            return error_result(e)
