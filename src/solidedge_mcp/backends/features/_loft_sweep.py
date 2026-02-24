"""Loft, sweep, and helix protrusion operations."""

import traceback
from typing import Any

import pythoncom
from win32com.client import VARIANT

from ..constants import (
    DirectionConstants,
    ExtentTypeConstants,
    LoftSweepConstants,
)
from ..logging import get_logger

_logger = get_logger(__name__)


class LoftSweepMixin:
    """Mixin providing loft, sweep, and helix protrusion methods."""

    def create_loft(self, profile_indices: list[int] | None = None) -> dict[str, Any]:
        """
        Create a loft feature between multiple profiles.

        Uses accumulated profiles from close_sketch() calls. Create 2+ sketches
        on different parallel planes, close each one, then call create_loft().

        Args:
            profile_indices: Optional list of profile indices to select from
                accumulated profiles. If None, uses all accumulated profiles.

        Returns:
            Dict with status and loft info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            # Get accumulated profiles from sketch manager
            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if profile_indices is not None:
                profiles = [all_profiles[i] for i in profile_indices]
            else:
                profiles = all_profiles

            if len(profiles) < 2:
                return {
                    "error": f"Loft requires at least 2 profiles"
                    f", got {len(profiles)}. Create sketches"
                    " on different planes and close each"
                    " one before calling create_loft()."
                }

            v_profiles, v_types, v_origins = self._make_loft_variant_arrays(profiles)

            # Try LoftedProtrusions.AddSimple first (works when a base feature exists)
            try:
                model = models.Item(1)
                lp = model.LoftedProtrusions
                lp.AddSimple(
                    len(profiles),
                    v_profiles,
                    v_types,
                    v_origins,
                    DirectionConstants.igRight,
                    ExtentTypeConstants.igNone,
                    ExtentTypeConstants.igNone,
                )
                self.sketch_manager.clear_accumulated_profiles()
                return {
                    "status": "created",
                    "type": "loft",
                    "num_profiles": len(profiles),
                    "method": "LoftedProtrusions.AddSimple",
                }
            except Exception:
                pass

            # Fall back to models.AddLoftedProtrusion (works as initial feature)
            v_seg = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [])
            model = models.AddLoftedProtrusion(
                len(profiles),
                v_profiles,
                v_types,
                v_origins,
                v_seg,  # SegmentMaps (empty)
                DirectionConstants.igRight,  # MaterialSide
                ExtentTypeConstants.igNone,
                0.0,
                None,  # Start extent
                ExtentTypeConstants.igNone,
                0.0,
                None,  # End extent
                ExtentTypeConstants.igNone,
                0.0,  # Start tangent
                ExtentTypeConstants.igNone,
                0.0,  # End tangent
            )

            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "loft",
                "num_profiles": len(profiles),
                "method": "models.AddLoftedProtrusion",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_sweep(self, path_profile_index: int | None = None) -> dict[str, Any]:
        """
        Create a sweep feature along a path.

        Requires at least 2 accumulated profiles: the first is the path (open profile),
        and the second is the cross-section (closed profile). Create the path sketch
        first (open, e.g. a line or arc), then create the cross-section sketch
        (closed, e.g. a circle) on a plane perpendicular to the path start.

        Args:
            path_profile_index: Index of the path profile in accumulated profiles
                (default: 0, the first accumulated profile)

        Returns:
            Dict with status and sweep info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if len(all_profiles) < 2:
                return {
                    "error": "Sweep requires at least 2 "
                    "profiles (path + cross-section), "
                    f"got {len(all_profiles)}. Create a "
                    "path sketch and a cross-section "
                    "sketch first."
                }

            path_idx = path_profile_index if path_profile_index is not None else 0
            path_profile = all_profiles[path_idx]
            cross_sections = [p for i, p in enumerate(all_profiles) if i != path_idx]

            _CS = LoftSweepConstants.igProfileBasedCrossSection

            # Path arrays
            v_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [path_profile])
            v_path_types = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, [_CS])

            # Cross-section arrays
            v_sections = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, cross_sections)
            v_section_types = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_I4, [_CS] * len(cross_sections)
            )
            v_origins = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
                [VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0]) for _ in cross_sections],
            )
            v_seg = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [])

            # AddSweptProtrusion: 15 required params
            models.AddSweptProtrusion(
                1,
                v_paths,
                v_path_types,  # Path (1 curve)
                len(cross_sections),
                v_sections,
                v_section_types,
                v_origins,
                v_seg,  # Sections
                DirectionConstants.igRight,  # MaterialSide
                ExtentTypeConstants.igNone,
                0.0,
                None,  # Start extent
                ExtentTypeConstants.igNone,
                0.0,
                None,  # End extent
            )

            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "sweep",
                "num_cross_sections": len(cross_sections),
                "method": "models.AddSweptProtrusion",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_helix(
        self, pitch: float, height: float,
        revolutions: float | None = None, direction: str = "Right",
    ) -> dict[str, Any]:
        """
        Create a helical feature.

        Args:
            pitch: Distance between coils (meters)
            height: Total height of helix (meters)
            revolutions: Number of turns (optional, calculated from pitch/height if not given)
            direction: 'Right' or 'Left' hand helix

        Returns:
            Dict with status and helix info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            # Calculate revolutions if not provided
            if revolutions is None:
                revolutions = height / pitch

            # AddFiniteBaseHelix
            models.AddFiniteBaseHelix(
                NumberOfProfiles=1,
                ProfileArray=(profile,),
                Pitch=pitch,
                Height=height,
                Revolutions=revolutions,
            )

            return {
                "status": "created",
                "type": "helix",
                "pitch": pitch,
                "height": height,
                "revolutions": revolutions,
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_loft_thin_wall(
        self, wall_thickness: float, profile_indices: list[int] | None = None
    ) -> dict[str, Any]:
        """
        Create a thin-walled loft feature between multiple profiles.

        Uses accumulated profiles from close_sketch() calls.

        Args:
            wall_thickness: Wall thickness in meters
            profile_indices: Optional list of profile indices to select

        Returns:
            Dict with status and loft info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if profile_indices is not None:
                profiles = [all_profiles[i] for i in profile_indices]
            else:
                profiles = all_profiles

            if len(profiles) < 2:
                return {"error": f"Loft requires at least 2 profiles, got {len(profiles)}."}

            v_profiles, v_types, v_origins = self._make_loft_variant_arrays(profiles)
            v_seg = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [])

            models.AddLoftedProtrusionWithThinWall(
                len(profiles),
                v_profiles,
                v_types,
                v_origins,
                v_seg,  # SegmentMaps
                DirectionConstants.igRight,  # MaterialSide
                ExtentTypeConstants.igNone,
                0.0,
                None,  # Start extent
                ExtentTypeConstants.igNone,
                0.0,
                None,  # End extent
                ExtentTypeConstants.igNone,
                0.0,  # Start tangent
                ExtentTypeConstants.igNone,
                0.0,  # End tangent
                wall_thickness,  # WallThickness
            )

            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "loft_thin_wall",
                "wall_thickness": wall_thickness,
                "num_profiles": len(profiles),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_sweep_thin_wall(
        self, wall_thickness: float, path_profile_index: int | None = None
    ) -> dict[str, Any]:
        """
        Create a thin-walled sweep feature along a path.

        Uses accumulated profiles: first is path (open), rest are cross-sections (closed).

        Args:
            wall_thickness: Wall thickness in meters
            path_profile_index: Index of the path profile (default: 0)

        Returns:
            Dict with status and sweep info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if len(all_profiles) < 2:
                return {
                    "error": f"Sweep requires at least 2 profiles (path + cross-section), "
                    f"got {len(all_profiles)}."
                }

            path_idx = path_profile_index if path_profile_index is not None else 0
            path_profile = all_profiles[path_idx]
            cross_sections = [p for i, p in enumerate(all_profiles) if i != path_idx]

            _CS = LoftSweepConstants.igProfileBasedCrossSection

            v_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [path_profile])
            v_path_types = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, [_CS])
            v_sections = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, cross_sections)
            v_section_types = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_I4, [_CS] * len(cross_sections)
            )
            v_origins = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
                [VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0]) for _ in cross_sections],
            )
            v_seg = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [])

            models.AddSweptProtrusionWithThinWall(
                1,
                v_paths,
                v_path_types,
                len(cross_sections),
                v_sections,
                v_section_types,
                v_origins,
                v_seg,
                DirectionConstants.igRight,
                ExtentTypeConstants.igNone,
                0.0,
                None,
                ExtentTypeConstants.igNone,
                0.0,
                None,
                wall_thickness,
            )

            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "sweep_thin_wall",
                "wall_thickness": wall_thickness,
                "num_cross_sections": len(cross_sections),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_helix_sync(
        self, pitch: float, height: float, revolutions: float | None = None
    ) -> dict[str, Any]:
        """Create synchronous helix feature"""
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            if revolutions is None:
                revolutions = height / pitch

            models.AddFiniteBaseHelixSync(
                NumberOfProfiles=1,
                ProfileArray=(profile,),
                Pitch=pitch,
                Height=height,
                Revolutions=revolutions,
            )

            return {
                "status": "created",
                "type": "helix_sync",
                "pitch": pitch,
                "height": height,
                "revolutions": revolutions,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_helix_thin_wall(
        self, pitch: float, height: float, wall_thickness: float, revolutions: float | None = None
    ) -> dict[str, Any]:
        """Create thin-walled helix feature"""
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            if revolutions is None:
                revolutions = height / pitch

            models.AddFiniteBaseHelixWithThinWall(
                NumberOfProfiles=1,
                ProfileArray=(profile,),
                Pitch=pitch,
                Height=height,
                Revolutions=revolutions,
                WallThickness=wall_thickness,
            )

            return {
                "status": "created",
                "type": "helix_thin_wall",
                "pitch": pitch,
                "height": height,
                "wall_thickness": wall_thickness,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_helix_sync_thin_wall(
        self, pitch: float, height: float, wall_thickness: float, revolutions: float | None = None
    ) -> dict[str, Any]:
        """Create synchronous thin-walled helix feature"""
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            if revolutions is None:
                revolutions = height / pitch

            models.AddFiniteBaseHelixSyncWithThinWall(
                NumberOfProfiles=1,
                ProfileArray=(profile,),
                Pitch=pitch,
                Height=height,
                Revolutions=revolutions,
                WallThickness=wall_thickness,
            )

            return {
                "status": "created",
                "type": "helix_sync_thin_wall",
                "pitch": pitch,
                "height": height,
                "wall_thickness": wall_thickness,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_helix_from_to(
        self, from_plane_index: int, to_plane_index: int, pitch: float
    ) -> dict[str, Any]:
        """
        Create a helix protrusion between two reference planes.

        Uses HelixProtrusions.AddFromTo(HelixAxis, AxisStart, NumCrossSections,
        CrossSectionArray, ProfileSide, FromFace, ToFace, Pitch, HelixDir).

        Args:
            from_plane_index: 1-based index of the starting reference plane
            to_plane_index: 1-based index of the ending reference plane
            pitch: Distance between coils in meters

        Returns:
            Dict with status and helix info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}
            if not refaxis:
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            ref_planes = doc.RefPlanes

            if from_plane_index < 1 or from_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid from_plane_index: "
                    f"{from_plane_index}. "
                    f"Count: {ref_planes.Count}"
                }
            if to_plane_index < 1 or to_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid to_plane_index: {to_plane_index}. Count: {ref_planes.Count}"
                }

            from_plane = ref_planes.Item(from_plane_index)
            to_plane = ref_planes.Item(to_plane_index)

            v_profiles = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            helix = model.HelixProtrusions
            helix.AddFromTo(
                refaxis,  # HelixAxis
                DirectionConstants.igRight,  # AxisStart
                1,  # NumCrossSections
                v_profiles,  # CrossSectionArray
                DirectionConstants.igRight,  # ProfileSide
                from_plane,  # FromFace
                to_plane,  # ToFace
                pitch,  # Pitch
                DirectionConstants.igRight,  # HelixDir
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "helix_from_to",
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
                "pitch": pitch,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_helix_from_to_thin_wall(
        self,
        from_plane_index: int,
        to_plane_index: int,
        pitch: float,
        wall_thickness: float,
    ) -> dict[str, Any]:
        """
        Create a thin-walled helix protrusion between two reference planes.

        Uses HelixProtrusions.AddFromToWithThinWall(HelixAxis, AxisStart,
        NumCrossSections, CrossSectionArray, ProfileSide, FromFace, ToFace,
        Pitch, HelixDir, WallThickness).

        Args:
            from_plane_index: 1-based index of the starting reference plane
            to_plane_index: 1-based index of the ending reference plane
            pitch: Distance between coils in meters
            wall_thickness: Wall thickness in meters

        Returns:
            Dict with status and helix info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}
            if not refaxis:
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            ref_planes = doc.RefPlanes

            if from_plane_index < 1 or from_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid from_plane_index: "
                    f"{from_plane_index}. "
                    f"Count: {ref_planes.Count}"
                }
            if to_plane_index < 1 or to_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid to_plane_index: {to_plane_index}. Count: {ref_planes.Count}"
                }

            from_plane = ref_planes.Item(from_plane_index)
            to_plane = ref_planes.Item(to_plane_index)

            v_profiles = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            helix = model.HelixProtrusions
            helix.AddFromToWithThinWall(
                refaxis,  # HelixAxis
                DirectionConstants.igRight,  # AxisStart
                1,  # NumCrossSections
                v_profiles,  # CrossSectionArray
                DirectionConstants.igRight,  # ProfileSide
                from_plane,  # FromFace
                to_plane,  # ToFace
                pitch,  # Pitch
                DirectionConstants.igRight,  # HelixDir
                wall_thickness,  # WallThickness
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "helix_from_to_thin_wall",
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
                "pitch": pitch,
                "wall_thickness": wall_thickness,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_helix_from_to_sync(
        self, from_plane_index: int, to_plane_index: int, pitch: float
    ) -> dict[str, Any]:
        """
        Create a synchronous helix protrusion between two reference planes.

        Uses HelixProtrusions.AddFromToSync(HelixAxis, AxisStart,
        NumCrossSections, CrossSectionArray, ProfileSide, Height, Pitch,
        NumberOfTurns, HelixDir, FromPlane, ToPlane, TaperAngle).

        Args:
            from_plane_index: 1-based index of the starting reference plane
            to_plane_index: 1-based index of the ending reference plane
            pitch: Distance between coils in meters

        Returns:
            Dict with status and helix info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}
            if not refaxis:
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            ref_planes = doc.RefPlanes

            if from_plane_index < 1 or from_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid from_plane_index: {from_plane_index}. "
                    f"Count: {ref_planes.Count}"
                }
            if to_plane_index < 1 or to_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid to_plane_index: {to_plane_index}. Count: {ref_planes.Count}"
                }

            from_plane = ref_planes.Item(from_plane_index)
            to_plane = ref_planes.Item(to_plane_index)

            v_profiles = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            helix = model.HelixProtrusions
            helix.AddFromToSync(
                refaxis,  # HelixAxis
                DirectionConstants.igRight,  # AxisStart
                1,  # NumCrossSections
                v_profiles,  # CrossSectionArray
                DirectionConstants.igRight,  # ProfileSide
                0.0,  # Height (ignored for FromTo)
                pitch,  # Pitch
                0.0,  # NumberOfTurns (calculated from FromTo)
                DirectionConstants.igRight,  # HelixDir
                from_plane,  # FromPlane
                to_plane,  # ToPlane
                0.0,  # TaperAngle
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "helix_from_to_sync",
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
                "pitch": pitch,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_helix_from_to_sync_thin_wall(
        self,
        from_plane_index: int,
        to_plane_index: int,
        pitch: float,
        wall_thickness: float,
    ) -> dict[str, Any]:
        """
        Create a synchronous thin-walled helix protrusion between two planes.

        Uses HelixProtrusions.AddFromToSyncWithThinWall.

        Args:
            from_plane_index: 1-based index of the starting reference plane
            to_plane_index: 1-based index of the ending reference plane
            pitch: Distance between coils in meters
            wall_thickness: Wall thickness in meters

        Returns:
            Dict with status and helix info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}
            if not refaxis:
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            ref_planes = doc.RefPlanes

            if from_plane_index < 1 or from_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid from_plane_index: {from_plane_index}. "
                    f"Count: {ref_planes.Count}"
                }
            if to_plane_index < 1 or to_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid to_plane_index: {to_plane_index}. Count: {ref_planes.Count}"
                }

            from_plane = ref_planes.Item(from_plane_index)
            to_plane = ref_planes.Item(to_plane_index)

            v_profiles = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            helix = model.HelixProtrusions
            helix.AddFromToSyncWithThinWall(
                refaxis,  # HelixAxis
                DirectionConstants.igRight,  # AxisStart
                1,  # NumCrossSections
                v_profiles,  # CrossSectionArray
                DirectionConstants.igRight,  # ProfileSide
                0.0,  # Height
                pitch,  # Pitch
                0.0,  # NumberOfTurns
                DirectionConstants.igRight,  # HelixDir
                from_plane,  # FromPlane
                to_plane,  # ToPlane
                True,  # ThinWall
                False,  # AddEndCaps
                True,  # RemoveInsideMaterial
                wall_thickness,  # Thickness
                DirectionConstants.igRight,  # ThicknessSide
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "helix_from_to_sync_thin_wall",
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
                "pitch": pitch,
                "wall_thickness": wall_thickness,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_loft_with_guides(
        self,
        guide_profile_indices: list[int] | None = None,
        profile_indices: list[int] | None = None,
    ) -> dict[str, Any]:
        """
        Create a loft feature with guide curves between multiple cross-section profiles.

        Uses accumulated profiles from close_sketch() calls. At minimum 2 cross-section
        profiles and 1 guide curve profile are required. Guide curves constrain the
        shape of the loft between cross-sections.

        Create cross-section sketches on parallel planes, then create guide curve
        sketches (open profiles) connecting the cross-sections. Close each sketch
        before calling this method.

        Args:
            guide_profile_indices: Indices of accumulated profiles to use as guide
                curves. If None, raises an error (guides are required).
            profile_indices: Indices of accumulated profiles to use as cross-sections.
                If None, uses all non-guide accumulated profiles.

        Returns:
            Dict with status and loft info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if guide_profile_indices is None or len(guide_profile_indices) == 0:
                return {
                    "error": "guide_profile_indices is required for loft with guides. "
                    "Specify which accumulated profile indices are guide curves."
                }

            guide_set = set(guide_profile_indices)

            if profile_indices is not None:
                profiles = [all_profiles[i] for i in profile_indices]
            else:
                profiles = [p for i, p in enumerate(all_profiles) if i not in guide_set]

            if len(profiles) < 2:
                return {
                    "error": f"Loft requires at least 2 cross-section profiles, "
                    f"got {len(profiles)}."
                }

            guides = [all_profiles[i] for i in guide_profile_indices]

            v_profiles, v_types, v_origins = self._make_loft_variant_arrays(profiles)
            v_seg = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [])
            v_num_guides = VARIANT(pythoncom.VT_I4, len(guides))
            v_guides = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, guides)

            # Try LoftedProtrusions.AddSimple with guide params first
            try:
                model = models.Item(1)
                lp = model.LoftedProtrusions
                lp.AddSimple(
                    len(profiles),
                    v_profiles,
                    v_types,
                    v_origins,
                    DirectionConstants.igRight,  # MaterialSide
                    ExtentTypeConstants.igNone,  # StartTangentType
                    ExtentTypeConstants.igNone,  # EndTangentType
                    v_num_guides,  # NumGuideCurves (optional)
                    v_guides,  # GuideCurves (optional)
                )
                self.sketch_manager.clear_accumulated_profiles()
                return {
                    "status": "created",
                    "type": "loft_with_guides",
                    "num_profiles": len(profiles),
                    "num_guides": len(guides),
                    "method": "LoftedProtrusions.AddSimple",
                }
            except Exception:
                pass

            # Fall back to LoftedProtrusions.Add (full params)
            try:
                model = models.Item(1)
                lp = model.LoftedProtrusions
                lp.Add(
                    len(profiles),
                    v_profiles,
                    v_types,
                    v_origins,
                    v_seg,  # SegmentMaps
                    DirectionConstants.igRight,  # MaterialSide
                    ExtentTypeConstants.igNone,
                    0.0,
                    None,  # Start extent
                    ExtentTypeConstants.igNone,
                    0.0,
                    None,  # End extent
                    ExtentTypeConstants.igNone,
                    0.0,  # Start tangent
                    ExtentTypeConstants.igNone,
                    0.0,  # End tangent
                    v_num_guides,  # NumGuideCurves (optional)
                    v_guides,  # GuideCurves (optional)
                )
                self.sketch_manager.clear_accumulated_profiles()
                return {
                    "status": "created",
                    "type": "loft_with_guides",
                    "num_profiles": len(profiles),
                    "num_guides": len(guides),
                    "method": "LoftedProtrusions.Add",
                }
            except Exception:
                pass

            # Fall back to models.AddLoftedProtrusion with guide params
            models.AddLoftedProtrusion(
                len(profiles),
                v_profiles,
                v_types,
                v_origins,
                v_seg,  # SegmentMaps
                DirectionConstants.igRight,  # MaterialSide
                ExtentTypeConstants.igNone,
                0.0,
                None,  # Start extent
                ExtentTypeConstants.igNone,
                0.0,
                None,  # End extent
                ExtentTypeConstants.igNone,
                0.0,  # Start tangent
                ExtentTypeConstants.igNone,
                0.0,  # End tangent
                v_num_guides,  # NumGuideCurves (optional)
                v_guides,  # GuideCurves (optional)
            )

            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "loft_with_guides",
                "num_profiles": len(profiles),
                "num_guides": len(guides),
                "method": "models.AddLoftedProtrusion",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
