"""Assembly-level feature operations."""

import contextlib
import math
import traceback
from typing import Any

from ..constants import (
    AssemblyFeaturePropertyConstants,
    ExtentTypeConstants,
)


class AssemblyFeaturesMixin:
    """Mixin providing assembly-level feature methods."""

    def pattern_component(
        self, component_index: int, count: int, spacing: float, direction: str = "X"
    ) -> dict[str, Any]:
        """
        Create a linear pattern of a component by placing copies with offset.

        Args:
            component_index: 0-based index of the source component
            count: Number of total instances (including original)
            spacing: Distance between instances (meters)
            direction: Pattern direction - 'X', 'Y', or 'Z'

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            occurrences = doc.Occurrences

            if component_index < 0 or component_index >= occurrences.Count:
                return {"error": f"Invalid component index: {component_index}"}

            source = occurrences.Item(component_index + 1)

            # Get the source file path
            try:
                file_path = source.OccurrenceFileName
            except Exception:
                return {"error": "Cannot determine source component file path"}

            # Get source position
            try:
                base_matrix = list(source.GetMatrix())
            except Exception:
                base_matrix = [1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1]

            dir_map = {"X": 12, "Y": 13, "Z": 14}
            dir_idx = dir_map.get(direction, 12)

            placed = []
            for i in range(1, count):
                matrix = list(base_matrix)
                matrix[dir_idx] = base_matrix[dir_idx] + (spacing * i)
                occ = occurrences.AddWithMatrix(file_path, matrix)
                placed.append(occ.Name if hasattr(occ, "Name") else f"copy_{i}")

            return {
                "status": "pattern_created",
                "source_component": component_index,
                "count": count,
                "spacing": spacing,
                "direction": direction,
                "placed_names": placed,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def _get_assembly_features(self):
        """Get the AssemblyFeatures object from the active assembly document."""
        doc = self.doc_manager.get_active_document()
        af = doc.AssemblyFeatures
        return doc, af

    def recompute_assembly_features(self, options: int = 0) -> dict[str, Any]:
        """
        Recompute all assembly features.

        Args:
            options: Recompute options (0 = default)
        """
        try:
            _doc, af = self._get_assembly_features()
            af.Recompute(options)
            return {"status": "recomputed", "options": options}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def _map_extent_type(self, extent_type: str) -> int:
        """Map extent type string to constant."""
        return {
            "Finite": ExtentTypeConstants.igFinite,
            "ThroughAll": ExtentTypeConstants.igThroughAll,
        }.get(extent_type, ExtentTypeConstants.igFinite)

    def _map_extent_side(self, extent_side: str) -> int:
        """Map extent side string to constant."""
        return {
            "OneSide": AssemblyFeaturePropertyConstants.igAssemblyFeatureOneSide,
            "BothSides": AssemblyFeaturePropertyConstants.igAssemblyFeatureBothSides,
        }.get(extent_side, AssemblyFeaturePropertyConstants.igAssemblyFeatureOneSide)

    def _map_profile_side(self, profile_side: str) -> int:
        """Map profile side string to constant."""
        return {
            "Left": AssemblyFeaturePropertyConstants.igAssemblyFeatureProfileLeft,
            "Right": AssemblyFeaturePropertyConstants.igAssemblyFeatureProfileRight,
            "Symmetric": AssemblyFeaturePropertyConstants.igAssemblyFeatureProfileSymmetric,
        }.get(profile_side, AssemblyFeaturePropertyConstants.igAssemblyFeatureProfileLeft)

    def _get_scope_parts_array(self, doc, scope_parts: list[int]) -> list:
        """Resolve occurrence indices to occurrence objects."""
        occurrences = doc.Occurrences
        return [occurrences.Item(idx + 1) for idx in scope_parts]

    def create_assembly_extruded_cutout(
        self,
        scope_parts: list[int],
        extent_type: str = "Finite",
        extent_side: str = "OneSide",
        profile_side: str = "Left",
        distance: float = 0.01,
    ) -> dict[str, Any]:
        """
        Create an assembly-level extruded cutout across multiple components.

        Requires an active sketch profile (close_sketch first).

        Args:
            scope_parts: List of occurrence indices the cutout spans
            extent_type: 'Finite' or 'ThroughAll'
            extent_side: 'OneSide' or 'BothSides'
            profile_side: 'Left', 'Right', or 'Symmetric'
            distance: Cutout depth in meters (for Finite)
        """
        try:
            doc, af = self._get_assembly_features()
            profiles = self.sketch_manager.get_accumulated_profiles()
            if not profiles:
                return {"error": "No profiles available. Create and close a sketch first."}

            scope = self._get_scope_parts_array(doc, scope_parts)
            cutouts = af.AssemblyFeaturesExtrudedCutouts
            cutouts.Add(
                len(scope),
                scope,
                len(profiles),
                profiles,
                self._map_extent_type(extent_type),
                self._map_extent_side(extent_side),
                self._map_profile_side(profile_side),
                distance,
                None,
                0,
                None,
                None,
            )
            return {
                "status": "created",
                "type": "assembly_extruded_cutout",
                "extent_type": extent_type,
                "distance": distance,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_assembly_revolved_cutout(
        self,
        scope_parts: list[int],
        extent_type: str = "Finite",
        extent_side: str = "OneSide",
        profile_side: str = "Left",
        angle: float = 360.0,
    ) -> dict[str, Any]:
        """
        Create an assembly-level revolved cutout across multiple components.

        Requires an active sketch profile with axis of revolution.

        Args:
            scope_parts: List of occurrence indices the cutout spans
            extent_type: 'Finite' or 'ThroughAll'
            extent_side: 'OneSide' or 'BothSides'
            profile_side: 'Left', 'Right', or 'Symmetric'
            angle: Revolution angle in degrees (for Finite)
        """
        try:
            doc, af = self._get_assembly_features()
            profiles = self.sketch_manager.get_accumulated_profiles()
            if not profiles:
                return {"error": "No profiles available. Create and close a sketch first."}

            scope = self._get_scope_parts_array(doc, scope_parts)
            cutouts = af.AssemblyFeaturesRevolvedCutouts
            cutouts.Add(
                len(scope),
                scope,
                len(profiles),
                profiles,
                None,  # pRefAxis (uses axis set on profile)
                self._map_extent_type(extent_type),
                self._map_extent_side(extent_side),
                self._map_profile_side(profile_side),
                math.radians(angle),
                None,
                0,
                None,
                None,
            )
            return {
                "status": "created",
                "type": "assembly_revolved_cutout",
                "extent_type": extent_type,
                "angle": angle,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_assembly_hole(
        self,
        scope_parts: list[int],
        extent_type: str = "Finite",
        extent_side: str = "OneSide",
        depth: float = 0.01,
    ) -> dict[str, Any]:
        """
        Create an assembly-level hole feature across multiple components.

        Requires an active sketch profile (circular).

        Args:
            scope_parts: List of occurrence indices the hole spans
            extent_type: 'Finite' or 'ThroughAll'
            extent_side: 'OneSide' or 'BothSides'
            depth: Hole depth in meters (for Finite)
        """
        try:
            doc, af = self._get_assembly_features()
            profiles = self.sketch_manager.get_accumulated_profiles()
            if not profiles:
                return {"error": "No profiles available. Create and close a sketch first."}

            scope = self._get_scope_parts_array(doc, scope_parts)
            holes = af.AssemblyFeaturesHoles
            holes.Add(
                len(scope),
                scope,
                len(profiles),
                profiles,
                self._map_extent_side(extent_side),
                None,  # pHoledata
                self._map_extent_type(extent_type),
                depth,
                None,
                None,  # pFromSurfOrPlane, pToSurfOrPlane
                None,
                0,  # pKeyPoint, pKeyPointFlags
            )
            return {
                "status": "created",
                "type": "assembly_hole",
                "extent_type": extent_type,
                "depth": depth,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_assembly_extruded_protrusion(
        self,
        extent_type: str = "Finite",
        extent_side: str = "OneSide",
        profile_side: str = "Left",
        distance: float = 0.01,
    ) -> dict[str, Any]:
        """
        Create an assembly-level extruded protrusion.

        Requires an active sketch profile (close_sketch first).

        Args:
            extent_type: 'Finite' or 'ThroughAll'
            extent_side: 'OneSide' or 'BothSides'
            profile_side: 'Left', 'Right', or 'Symmetric'
            distance: Extrusion depth in meters (for Finite)
        """
        try:
            doc, af = self._get_assembly_features()
            profiles = self.sketch_manager.get_accumulated_profiles()
            if not profiles:
                return {"error": "No profiles available. Create and close a sketch first."}

            protrusions = af.ExtrudedProtrusions
            protrusions.Add(
                len(profiles),
                profiles,
                self._map_extent_type(extent_type),
                self._map_extent_side(extent_side),
                self._map_profile_side(profile_side),
                distance,
                None,
                0,
                None,
                None,
            )
            return {
                "status": "created",
                "type": "assembly_extruded_protrusion",
                "extent_type": extent_type,
                "distance": distance,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_assembly_revolved_protrusion(
        self,
        extent_type: str = "Finite",
        extent_side: str = "OneSide",
        profile_side: str = "Left",
        angle: float = 360.0,
    ) -> dict[str, Any]:
        """
        Create an assembly-level revolved protrusion.

        Requires an active sketch profile with axis of revolution.

        Args:
            extent_type: 'Finite' or 'ThroughAll'
            extent_side: 'OneSide' or 'BothSides'
            profile_side: 'Left', 'Right', or 'Symmetric'
            angle: Revolution angle in degrees (for Finite)
        """
        try:
            doc, af = self._get_assembly_features()
            profiles = self.sketch_manager.get_accumulated_profiles()
            if not profiles:
                return {"error": "No profiles available. Create and close a sketch first."}

            protrusions = af.RevolvedProtrusions
            protrusions.Add(
                len(profiles),
                profiles,
                None,  # pRefAxis
                self._map_extent_type(extent_type),
                self._map_extent_side(extent_side),
                self._map_profile_side(profile_side),
                math.radians(angle),
                None,
                0,
                None,
                None,
            )
            return {
                "status": "created",
                "type": "assembly_revolved_protrusion",
                "extent_type": extent_type,
                "angle": angle,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_assembly_mirror(
        self,
        feature_indices: list[int],
        plane_index: int = 1,
        mirror_type: int = 1,
    ) -> dict[str, Any]:
        """
        Mirror assembly features across a reference plane.

        Args:
            feature_indices: List of assembly feature indices to mirror (0-based)
            plane_index: Reference plane index (1-based: 1=Top, 2=Front, 3=Right)
            mirror_type: Mirror option from FeaturePropertyConstants
        """
        try:
            doc, af = self._get_assembly_features()

            # Get the mirror plane
            plane = doc.RefPlanes.Item(plane_index)

            # Get features to mirror from the assembly features collection
            features_to_mirror = []
            # Iterate available assembly feature collections to find features
            for collection_name in [
                "AssemblyFeaturesExtrudedCutouts",
                "AssemblyFeaturesRevolvedCutouts",
                "AssemblyFeaturesHoles",
            ]:
                with contextlib.suppress(Exception):
                    coll = getattr(af, collection_name)
                    for fi in feature_indices:
                        with contextlib.suppress(Exception):
                            features_to_mirror.append(coll.Item(fi + 1))

            if not features_to_mirror:
                return {"error": "No features found at the specified indices"}

            mirrors = af.AssemblyFeaturesMirrors
            mirrors.Add(
                len(features_to_mirror),
                features_to_mirror,
                plane,
                mirror_type,
            )
            return {
                "status": "created",
                "type": "assembly_mirror",
                "num_features": len(features_to_mirror),
                "plane_index": plane_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_assembly_pattern(
        self,
        feature_indices: list[int],
        pattern_type: str = "Rectangular",
    ) -> dict[str, Any]:
        """
        Pattern assembly features.

        Args:
            feature_indices: List of assembly feature indices to pattern (0-based)
            pattern_type: 'Rectangular' or 'Circular'
        """
        try:
            doc, af = self._get_assembly_features()
            profiles = self.sketch_manager.get_accumulated_profiles()

            pattern_map = {
                "Rectangular": 1,  # igRectangularPattern
                "Circular": 2,  # igCircularPattern
            }
            pattern_const = pattern_map.get(pattern_type, 1)

            # Get features to pattern
            features_to_pattern = []
            for collection_name in [
                "AssemblyFeaturesExtrudedCutouts",
                "AssemblyFeaturesRevolvedCutouts",
                "AssemblyFeaturesHoles",
            ]:
                with contextlib.suppress(Exception):
                    coll = getattr(af, collection_name)
                    for fi in feature_indices:
                        with contextlib.suppress(Exception):
                            features_to_pattern.append(coll.Item(fi + 1))

            if not features_to_pattern:
                return {"error": "No features found at the specified indices"}

            profile = profiles[0] if profiles else None
            patterns = af.AssemblyFeaturesPatterns
            patterns.Add(
                len(features_to_pattern),
                features_to_pattern,
                profile,
                pattern_const,
            )
            return {
                "status": "created",
                "type": "assembly_pattern",
                "pattern_type": pattern_type,
                "num_features": len(features_to_pattern),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_assembly_swept_protrusion(
        self,
        num_trace_curves: int = 1,
        num_cross_sections: int = 1,
    ) -> dict[str, Any]:
        """
        Create an assembly-level swept protrusion.

        Requires accumulated profiles: first profile(s) as trace curve(s),
        remaining as cross-section(s).

        Args:
            num_trace_curves: Number of trace curve profiles
            num_cross_sections: Number of cross-section profiles
        """
        try:
            doc, af = self._get_assembly_features()
            profiles = self.sketch_manager.get_accumulated_profiles()
            if len(profiles) < num_trace_curves + num_cross_sections:
                return {
                    "error": f"Need {num_trace_curves + num_cross_sections} profiles, "
                    f"got {len(profiles)}"
                }

            trace_curves = profiles[:num_trace_curves]
            cross_sections = profiles[num_trace_curves : num_trace_curves + num_cross_sections]

            import pythoncom
            from win32com.client import VARIANT

            v_origins = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
                [VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0]) for _ in cross_sections],
            )

            swept = af.AssemblyFeaturesSweptProtrusions
            swept.Add(
                num_trace_curves,
                trace_curves,
                num_cross_sections,
                cross_sections,
                v_origins,
            )
            return {
                "status": "created",
                "type": "assembly_swept_protrusion",
                "num_trace_curves": num_trace_curves,
                "num_cross_sections": num_cross_sections,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
