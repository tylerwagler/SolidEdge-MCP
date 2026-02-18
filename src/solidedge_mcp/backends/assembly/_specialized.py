"""Specialized assembly subsystem operations (virtual components, tubes, frames, wiring)."""

import contextlib
import os
import traceback
from typing import Any


class SpecializedMixin:
    """Mixin providing specialized assembly subsystem methods."""

    # -- Virtual Components --------------------------------------------------

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
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            if not os.path.exists(filename):
                return {"error": f"File not found: {filename}"}

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            if not os.path.exists(part_filename):
                return {"error": f"File not found: {part_filename}"}

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            # Resolve segment indices to occurrence objects
            segments = []
            for idx in segment_indices:
                if idx < 0 or idx >= occurrences.Count:
                    return {"error": f"Invalid segment index: {idx}. Count: {occurrences.Count}"}
                segments.append(occurrences.Item(idx + 1))

            import pythoncom
            from win32com.client import VARIANT

            v_segments = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, segments)

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
            return {"error": str(e), "traceback": traceback.format_exc()}

    # -- Structural Frames ---------------------------------------------------

    def add_structural_frame(
        self,
        part_filename: str,
        path_indices: list[int],
    ) -> dict[str, Any]:
        """
        Add a structural frame to the assembly.

        Uses StructuralFrames.Add with VARIANT-wrapped path array.

        Args:
            part_filename: Path to the frame cross-section part file
            path_indices: 0-based indices of occurrences defining the path

        Returns:
            Dict with status and frame info
        """
        try:
            if not os.path.exists(part_filename):
                return {"error": f"File not found: {part_filename}"}

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            paths = []
            for idx in path_indices:
                if idx < 0 or idx >= occurrences.Count:
                    return {"error": f"Invalid path index: {idx}. Count: {occurrences.Count}"}
                paths.append(occurrences.Item(idx + 1))

            import pythoncom
            from win32com.client import VARIANT

            v_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, paths)

            frames = doc.StructuralFrames
            frame = frames.Add(part_filename, len(paths), v_paths)

            result: dict[str, Any] = {
                "status": "created",
                "type": "structural_frame",
                "part_filename": part_filename,
                "num_paths": len(paths),
            }

            with contextlib.suppress(Exception):
                result["name"] = frame.Name

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_structural_frame_by_orientation(
        self,
        part_filename: str,
        coord_system_name: str,
        path_indices: list[int],
    ) -> dict[str, Any]:
        """
        Add a structural frame with a specific coordinate system orientation.

        Uses StructuralFrames.AddByOrientation.

        Args:
            part_filename: Path to the frame cross-section part file
            coord_system_name: Name of the coordinate system to orient by
            path_indices: 0-based indices of occurrences defining the path

        Returns:
            Dict with status and frame info
        """
        try:
            if not os.path.exists(part_filename):
                return {"error": f"File not found: {part_filename}"}

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            paths = []
            for idx in path_indices:
                if idx < 0 or idx >= occurrences.Count:
                    return {"error": f"Invalid path index: {idx}. Count: {occurrences.Count}"}
                paths.append(occurrences.Item(idx + 1))

            import pythoncom
            from win32com.client import VARIANT

            v_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, paths)

            frames = doc.StructuralFrames
            frame = frames.AddByOrientation(coord_system_name, len(paths), v_paths)

            result: dict[str, Any] = {
                "status": "created",
                "type": "structural_frame_oriented",
                "part_filename": part_filename,
                "coord_system": coord_system_name,
                "num_paths": len(paths),
            }

            with contextlib.suppress(Exception):
                result["name"] = frame.Name

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # -- Splices -------------------------------------------------------------

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
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            occurrences = doc.Occurrences

            conductors = []
            for idx in conductor_indices:
                if idx < 0 or idx >= occurrences.Count:
                    return {"error": f"Invalid conductor index: {idx}. Count: {occurrences.Count}"}
                conductors.append(occurrences.Item(idx + 1))

            import pythoncom
            from win32com.client import VARIANT

            v_conductors = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, conductors)

            splices = doc.Splices
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

            if len(path_indices) != len(path_directions):
                return {"error": "path_indices and path_directions must have the same length"}

            occurrences = doc.Occurrences

            paths = []
            for idx in path_indices:
                if idx < 0 or idx >= occurrences.Count:
                    return {"error": f"Invalid path index: {idx}. Count: {occurrences.Count}"}
                paths.append(occurrences.Item(idx + 1))

            import pythoncom
            from win32com.client import VARIANT

            v_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, paths)
            v_dirs = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_BOOL, path_directions)

            wires = doc.Wires
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

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

            import pythoncom
            from win32com.client import VARIANT

            v_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, paths)
            v_dirs = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_BOOL, path_directions)
            v_wires = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, wires_list)
            v_split_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, split_paths)
            v_split_dirs = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_BOOL, split_dirs)

            cables = doc.Cables
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Occurrences"):
                return {"error": "Active document is not an assembly"}

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

            import pythoncom
            from win32com.client import VARIANT

            v_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, paths)
            v_dirs = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_BOOL, path_directions)
            v_conductors = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, conductors)
            v_split_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, split_paths)
            v_split_dirs = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_BOOL, split_dirs)

            bundles = doc.Bundles
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
            return {"error": str(e), "traceback": traceback.format_exc()}
