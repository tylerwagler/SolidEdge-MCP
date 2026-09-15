"""Assembly-level feature operations."""

import math
from typing import Any

import pythoncom
from win32com.client import VARIANT

from solidedge_mcp.backends.errors import error_result

from ..comutil import profile_origin
from ..constants import (
    AssemblyFeaturePropertyConstants,
    ExtentTypeConstants,
)
from ..logging import get_logger
from ._base import com_get, verifies_assembly_geometry

_logger = get_logger(__name__)


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
            _logger.info(
                "Creating component pattern: index=%d, count=%d, spacing=%s",
                component_index,
                count,
                spacing,
            )
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
                base_matrix = self._get_occurrence_matrix(source)
            except Exception:
                base_matrix = [1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1]

            dir_map = {"X": 12, "Y": 13, "Z": 14}
            dir_idx = dir_map.get(direction, 12)

            placed = []
            for i in range(1, count):
                matrix = list(base_matrix)
                matrix[dir_idx] = base_matrix[dir_idx] + (spacing * i)
                occ = occurrences.AddWithMatrix(file_path, matrix)
                placed.append(com_get(occ, "Name", f"copy_{i}"))

            return {
                "status": "pattern_created",
                "source_component": component_index,
                "count": count,
                "spacing": spacing,
                "direction": direction,
                "placed_names": placed,
            }
        except Exception as e:
            _logger.error(f"Failed to create component pattern: {e}")
            return error_result(e)

    def _get_assembly_features(self) -> tuple[Any, Any]:
        """Get the AssemblyFeatures object from the active assembly document."""
        doc = self.doc_manager.get_active_document()
        af = doc.AssemblyFeatures
        return doc, af

    def recompute_assembly_features(self, options: int = 0) -> dict[str, Any]:
        """Update the assembly.

        ``AssemblyFeatures.Recompute(options)`` matches the type library but
        answers E_FAIL on Solid Edge 2026 whatever it is given -- verified on
        an empty assembly and on one holding a component, with options 0. The
        object is not a normal collection either: it has no Count.

        ``AssemblyDocument.UpdateAll()`` is the update that works, and it is
        what this does.

        Args:
            options: Accepted for compatibility; UpdateAll takes no options.
        """
        try:
            _logger.info("Updating assembly (options=%s ignored by UpdateAll)", options)
            doc = self.doc_manager.get_active_document()
            err = self._require_assembly(doc)
            if err:
                return err
            doc.UpdateAll()
            return {"status": "recomputed", "method": "AssemblyDocument.UpdateAll"}
        except Exception as e:
            _logger.error(f"Failed to update assembly: {e}")
            return error_result(e)

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

    def _get_scope_parts_array(self, doc: Any, scope_parts: list[int]) -> list[Any]:
        """Resolve occurrence indices to occurrence objects."""
        occurrences = doc.Occurrences
        return [occurrences.Item(idx + 1) for idx in scope_parts]

    @verifies_assembly_geometry
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
            _logger.info(
                "Creating assembly extruded cutout: dist=%s, extent=%s",
                distance,
                extent_type,
            )
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
            _logger.error(f"Failed to create assembly extruded cutout: {e}")
            return error_result(e)

    @verifies_assembly_geometry
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
            _logger.info(f"Creating assembly revolved cutout: angle={angle}, extent={extent_type}")
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
            _logger.error(f"Failed to create assembly revolved cutout: {e}")
            return error_result(e)

    @verifies_assembly_geometry
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
            _logger.info(f"Creating assembly hole: depth={depth}, extent={extent_type}")
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
            _logger.error(f"Failed to create assembly hole: {e}")
            return error_result(e)

    @verifies_assembly_geometry
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
            _logger.info(
                "Creating assembly extruded protrusion: dist=%s, extent=%s",
                distance,
                extent_type,
            )
            doc, af = self._get_assembly_features()
            profiles = self.sketch_manager.get_accumulated_profiles()
            if not profiles:
                return {"error": "No profiles available. Create and close a sketch first."}

            # AssemblyFeatures exposes AssemblyFeaturesExtrudedProtrusions; the
            # bare ExtrudedProtrusions name belongs to Part.tlb, whose Add takes
            # 35 arguments. Add here is (nNumProfiles, pProfiles, ExtentType,
            # pExtentSide, profileSide, pdDistance, pKeyPoint, pKeyPointFlags,
            # pFromSurfOrPlane, pToSurfOrPlane).
            # The PROPERTY is ExtrudedProtrusions; its TYPE is the
            # interface AssemblyFeaturesExtrudedProtrusions. Reading the
            # interface name off AssemblyFeatures raises, and the member
            # check cannot see it because that name is real elsewhere.
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
            _logger.error(f"Failed to create assembly extruded protrusion: {e}")
            return error_result(e)

    @verifies_assembly_geometry
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
            _logger.info(
                "Creating assembly revolved protrusion: angle=%s, extent=%s",
                angle,
                extent_type,
            )
            doc, af = self._get_assembly_features()
            profiles = self.sketch_manager.get_accumulated_profiles()
            if not profiles:
                return {"error": "No profiles available. Create and close a sketch first."}

            # AssemblyFeaturesRevolvedProtrusions.Add(nNumProfiles, pProfiles,
            #     pRefAxis, ExtentType, ExtentSide, profileSide, pdAngle,
            #     KeyPointOrTangentFace, KeyPointFlags, pFromSurface, pToSurface)
            # As above: the property is RevolvedProtrusions, the interface
            # it returns is AssemblyFeaturesRevolvedProtrusions.
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
            _logger.error(f"Failed to create assembly revolved protrusion: {e}")
            return error_result(e)

    def create_assembly_mirror(
        self,
        feature_indices: list[int],
        plane_index: int = 1,
        mirror_type: int = 1,
    ) -> dict[str, Any]:
        """
        Mirror assembly features across a reference plane.

        NOT AVAILABLE via COM automation. ``AssemblyFeaturesMirrors.Add``
        returns E_ACCESSDENIED (0x80070005) on Solid Edge 2025/2026 in every
        argument combination tested, so this method never touches COM and
        always returns an ``unsupported`` error dict. The signature is kept so
        tool dispatch keeps working.

        Args:
            feature_indices: List of assembly feature indices to mirror (0-based)
            plane_index: Reference plane index (1-based: 1=Top/XY, 2=Right/YZ, 3=Front/XZ)
            mirror_type: Mirror option from FeaturePropertyConstants
        """
        _logger.warning(
            "Assembly mirror requested (features=%s, plane=%d) but "
            "AssemblyFeaturesMirrors.Add is not usable through COM; refusing.",
            feature_indices,
            plane_index,
        )
        return {
            "error": (
                "Assembly-level mirror is not available through Solid Edge COM "
                "automation (AssemblyFeaturesMirrors.Add returns E_ACCESSDENIED "
                "0x80070005 on SE 2025/2026). Mirror the feature in the part "
                "document instead, or create individual assembly features at "
                "each position."
            ),
            "unsupported": True,
            "type": "assembly_mirror",
            "feature_indices": list(feature_indices),
            "plane_index": plane_index,
            "mirror_type": mirror_type,
        }

    def create_assembly_pattern(
        self,
        feature_indices: list[int],
        pattern_type: str = "Rectangular",
    ) -> dict[str, Any]:
        """
        Pattern assembly features.

        NOT AVAILABLE via COM automation. ``AssemblyFeaturesPatterns.Add``
        returns E_ACCESSDENIED (0x80070005) on Solid Edge 2025/2026 in every
        argument combination tested, so this method never touches COM and
        always returns an ``unsupported`` error dict. The signature is kept so
        tool dispatch keeps working.

        Args:
            feature_indices: List of assembly feature indices to pattern (0-based)
            pattern_type: 'Rectangular' or 'Circular'
        """
        _logger.warning(
            "Assembly pattern requested (features=%s, type=%s) but "
            "AssemblyFeaturesPatterns.Add is not usable through COM; refusing.",
            feature_indices,
            pattern_type,
        )
        return {
            "error": (
                "Assembly-level feature pattern is not available through Solid Edge "
                "COM automation (AssemblyFeaturesPatterns.Add returns E_ACCESSDENIED "
                "0x80070005 on SE 2025/2026). Pattern the feature in the part "
                "document instead, use pattern_component() to pattern whole "
                "components, or create individual assembly features at each position."
            ),
            "unsupported": True,
            "type": "assembly_pattern",
            "feature_indices": list(feature_indices),
            "pattern_type": pattern_type,
        }

    @verifies_assembly_geometry
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
            _logger.info(
                "Creating assembly swept protrusion: traces=%d, sections=%d",
                num_trace_curves,
                num_cross_sections,
            )
            doc, af = self._get_assembly_features()
            profiles = self.sketch_manager.get_accumulated_profiles()
            if len(profiles) < num_trace_curves + num_cross_sections:
                return {
                    "error": f"Need {num_trace_curves + num_cross_sections} profiles, "
                    f"got {len(profiles)}"
                }

            trace_curves = profiles[:num_trace_curves]
            cross_sections = profiles[num_trace_curves : num_trace_curves + num_cross_sections]

            # A SAFEARRAY of SAFEARRAY(VT_R8): the inner VARIANTs are required, only

            # the outer wrapper is not. Dropping them broke the lofted cutout.

            v_origins = [
                VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list(profile_origin(p)))
                for p in cross_sections
            ]

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
            _logger.error(f"Failed to create assembly swept protrusion: {e}")
            return error_result(e)
