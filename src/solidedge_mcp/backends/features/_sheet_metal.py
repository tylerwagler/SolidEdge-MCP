"""Sheet metal and part feature operations (flanges, tabs, bends, dimples, etc.)."""

import contextlib
import math
from typing import Any

import pythoncom
from win32com.client import VARIANT

from solidedge_mcp.backends.errors import error_result

from ..comutil import com_get
from ..constants import (
    DirectionConstants,
    ExtentTypeConstants,
    FaceQueryConstants,
    FeaturePropertyConstants,
    KeyPointExtentConstants,
    ModelingModeConstants,
    OffsetSideConstants,
    SETargetConstructionBodyOption,
    SETargetDesignBodyOption,
    TreatmentTypeConstants,
)
from ..logging import get_logger
from ._base import verifies_collection_growth, verifies_geometry

_logger = get_logger(__name__)

# constant.tlb > FeaturePropertyConstants. Local because backends/constants.py
# does not carry these members yet.
_IG_EXTEND = 9  # igExtend - extend the rib profile to the body
_IG_THK_NORMAL_TO_PROFILE_PLANE = 12  # igThkNormalToProfilePlane

# constant.tlb > DrawnCutoutFeatureConstants
_SE_DRAWN_CUTOUT_DEPTH_LEFT = 1
_SE_DRAWN_CUTOUT_DEPTH_RIGHT = 2
_SE_DRAWN_CUTOUT_MATERIAL_INSIDE = 3
_SE_DRAWN_CUTOUT_PROFILE_RIGHT = 6

# constant.tlb > LouverFeatureConstants
_SE_LOUVER_DEPTH_DIRECTION_LEFT = 1
_SE_LOUVER_DEPTH_DIRECTION_RIGHT = 2
_SE_LOUVER_HEIGHT_NORMAL = 7


def _count(collection: Any) -> int:
    """A collection's Count, or zero when it is not a real number.

    The isinstance check matters: a unittest.mock stand-in answers ``int()``
    with 1, so converting blindly would make every mocked profile look as
    though it held a circle.
    """
    value = com_get(collection, "Count", 0)
    if isinstance(value, bool) or not isinstance(value, int | float):
        return 0
    return int(value)


def _lines_form_a_loop(profile: Any) -> bool:
    """True when every line endpoint is shared, which makes a closed chain."""
    lines = com_get(profile, "Lines2d")
    count = _count(lines)
    if count < 3:
        return False
    tally: dict[tuple[float, float], int] = {}
    for i in range(1, count + 1):
        element = lines.Item(i)
        for member in ("GetStartPoint", "GetEndPoint"):
            try:
                point = getattr(element, member)()
            except Exception:
                return False
            key = (round(float(point[0]), 9), round(float(point[1]), 9))
            tally[key] = tally.get(key, 0) + 1
    return all(shared == 2 for shared in tally.values())


class SheetMetalMixin:
    """Mixin providing sheet metal and miscellaneous part feature methods."""

    def _require_open_profile(self, profile: Any, feature: str) -> dict[str, Any] | None:
        """Refuse a closed profile for a feature that needs a line.

        A louver line and a bend line are open profiles. Handing either a
        closed shape fails with a bare E_FAIL from Solid Edge, which tells the
        caller nothing. Verified on Solid Edge 2026: the same call succeeds
        with a line and fails with a circle.
        """
        closed_kinds = []
        for name, label in (
            ("Circles2d", "circle"),
            ("Ellipses2d", "ellipse"),
            ("Boundaries2d", "closed boundary"),
        ):
            if _count(com_get(profile, name)):
                closed_kinds.append(label)

        if not closed_kinds and _lines_form_a_loop(profile):
            closed_kinds.append("closed chain of lines")

        if not closed_kinds:
            return None
        return {
            "error": (
                f"A {feature} is formed along an open line, and this sketch "
                f"contains a {closed_kinds[0]}. Draw a single line where the "
                f"{feature} should run, close the sketch, and try again."
            ),
            "profile": closed_kinds,
        }

    @verifies_geometry
    def create_base_flange(
        self, width: float, thickness: float, bend_radius: float | None = None
    ) -> dict[str, Any]:
        """
        Create a base contour flange (sheet metal).

        Uses Models.AddBaseContourFlange(pProfile, varThicknessSide,
        varExtentType, varProjectionSide, varProjectionDistance, varRadius, ...).
        The call has no thickness argument - the material thickness comes from
        the sheet metal document itself - so thickness is only echoed back.

        Args:
            width: Flange projection distance (meters)
            thickness: Material thickness (meters, not passed to COM)
            bend_radius: Bend radius (meters, optional)

        Returns:
            Dict with status and flange info
        """
        if width <= 0:
            return {
                "error": (
                    "Models.AddBaseContourFlange needs a projection distance; "
                    "pass width (in meters)."
                )
            }
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            if bend_radius is None:
                bend_radius = thickness * 2

            models.AddBaseContourFlange(
                profile,
                DirectionConstants.igRight,  # varThicknessSide
                ExtentTypeConstants.igFinite,  # varExtentType
                DirectionConstants.igRight,  # varProjectionSide
                width,  # varProjectionDistance
                bend_radius,  # varRadius
            )

            return {
                "status": "created",
                "type": "base_flange",
                "width": width,
                "thickness": thickness,
                "bend_radius": bend_radius,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_base_tab(self, thickness: float, width: float | None = None) -> dict[str, Any]:
        """
        Create a base tab (sheet metal).

        Uses Models.AddBaseTab(Profile, ExtentSide) - the only two arguments the
        call takes. Material thickness comes from the sheet metal document, so
        thickness is only echoed back.

        Args:
            thickness: Material thickness (meters, not passed to COM)
            width: Tab width (meters, not passed to COM)

        Returns:
            Dict with status and tab info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            models.AddBaseTab(profile, DirectionConstants.igRight)

            return {
                "status": "created",
                "type": "base_tab",
                "thickness": thickness,
                "width": width,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_lofted_flange(self, thickness: float) -> dict[str, Any]:
        """
        Create a lofted flange (sheet metal).

        NOT AVAILABLE via COM automation. Models.AddLoftedFlange takes
        (NumSections, CrossSections, CrossSectionTypes, Origins, OriginRefs,
        ThicknessSide, varRadius, varNeutralFactor, varBnParamType): it needs the
        cross-section profiles plus an origin and origin reference for each one,
        none of which this server can supply. The call is never made.
        """
        return {
            "error": (
                "Lofted flanges are not available through this server: "
                "Models.AddLoftedFlange needs cross-section profiles with an "
                "origin and origin reference for each section. Create the lofted "
                "flange in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "lofted_flange",
            "thickness": thickness,
        }

    @verifies_geometry
    def create_web_network(
        self,
        thickness: float = 0.0,
        depth: float = 0.0,
        direction: str = "Normal",
    ) -> dict[str, Any]:
        """
        Create a web network from the accumulated sketch profiles.

        Uses Models.AddWebNetwork(nNumProfiles, aProfiles, dThickness,
        WebDirection, dFiniteDepth, TreatmentType).

        Args:
            thickness: Web thickness in meters
            depth: Finite depth of the web in meters
            direction: 'Normal', 'Reverse' or 'Symmetric'

        Returns:
            Dict with status and web network info
        """
        if thickness <= 0:
            return {"error": "Models.AddWebNetwork needs a positive web thickness; pass thickness."}
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            profiles = self.sketch_manager.get_accumulated_profiles()
            if not profiles:
                active = self.sketch_manager.get_active_sketch()
                profiles = [active] if active else []
            if not profiles:
                return {"error": "No sketch profiles. Create and close at least one sketch first."}

            dir_map = {
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
                "Symmetric": DirectionConstants.igSymmetric,
            }
            web_direction = dir_map.get(direction, DirectionConstants.igRight)

            profile_arr = list(profiles)

            models.AddWebNetwork(
                len(profiles),
                profile_arr,
                thickness,
                web_direction,
                depth,
                TreatmentTypeConstants.seTreatmentNone,
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "web_network",
                "profile_count": len(profiles),
                "thickness": thickness,
                "depth": depth,
                "direction": direction,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_base_contour_flange_advanced(
        self,
        thickness: float,
        bend_radius: float,
        relief_type: str = "Default",
        width: float = 0.0,
    ) -> dict[str, Any]:
        """
        Create a base contour flange with bend deduction or bend allowance.

        Uses Models.AddBaseContourFlangeByBendDeductionOrBendAllowance(pProfile,
        varThicknessSide, varExtentType, varProjectionSide,
        varProjectionDistance, varRadius, ...). The bend calculation method is
        the 19th/20th optional argument, so it cannot be set without also
        supplying the twelve mitre arguments in between; only the six required
        arguments are passed. There is no thickness argument - the material
        thickness comes from the sheet metal document.

        Args:
            thickness: Material thickness (meters, not passed to COM)
            bend_radius: Bend radius in meters
            relief_type: Accepted for compatibility; not passed to COM
            width: Flange projection distance in meters (required)

        Returns:
            Dict with status and flange info
        """
        if width <= 0:
            return {
                "error": (
                    "Models.AddBaseContourFlangeByBendDeductionOrBendAllowance "
                    "needs a projection distance; pass width (in meters)."
                )
            }
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            models.AddBaseContourFlangeByBendDeductionOrBendAllowance(
                profile,
                DirectionConstants.igRight,  # varThicknessSide
                ExtentTypeConstants.igFinite,  # varExtentType
                DirectionConstants.igRight,  # varProjectionSide
                width,  # varProjectionDistance
                bend_radius,  # varRadius
            )

            return {
                "status": "created",
                "type": "base_contour_flange_advanced",
                "width": width,
                "thickness": thickness,
                "bend_radius": bend_radius,
                "relief_type": relief_type,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_base_tab_multi_profile(self, thickness: float) -> dict[str, Any]:
        """
        Create a base tab from the accumulated sketch profiles.

        Uses Models.AddBaseTabWithMultipleProfiles(NumberOfProfiles,
        ProfileArray, ExtentSide). There is no thickness argument - the material
        thickness comes from the sheet metal document.

        Args:
            thickness: Material thickness (meters, not passed to COM)

        Returns:
            Dict with status and tab info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            profiles = self.sketch_manager.get_accumulated_profiles()
            if not profiles:
                active = self.sketch_manager.get_active_sketch()
                profiles = [active] if active else []
            if not profiles:
                return {"error": "No active sketch profile"}

            profile_arr = list(profiles)

            models.AddBaseTabWithMultipleProfiles(
                len(profiles),
                profile_arr,
                DirectionConstants.igRight,
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "base_tab_multi_profile",
                "profile_count": len(profiles),
                "thickness": thickness,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_lofted_flange_advanced(self, thickness: float, bend_radius: float) -> dict[str, Any]:
        """
        Create a lofted flange with bend deduction or bend allowance.

        NOT AVAILABLE via COM automation.
        Models.AddLoftedFlangeByBendDeductionOrBendAllowance takes 25 required
        arguments - cross-section profiles, per-section origins and origin
        references, vertex maps, bend-divide and auto-relief settings - which
        this server cannot supply. The call is never made.
        """
        return {
            "error": (
                "Lofted flanges are not available through this server: "
                "Models.AddLoftedFlangeByBendDeductionOrBendAllowance needs 25 "
                "arguments including cross-section profiles, per-section origins "
                "and vertex maps. Create the lofted flange in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "lofted_flange_advanced",
            "thickness": thickness,
            "bend_radius": bend_radius,
        }

    @verifies_geometry
    def create_lofted_flange_ex(self, thickness: float) -> dict[str, Any]:
        """
        Create an extended lofted flange.

        NOT AVAILABLE via COM automation. Models.AddLoftedFlangeEx takes 25
        required arguments - cross-section profiles, per-section origins and
        origin references, vertex maps, bend-divide and auto-relief settings -
        which this server cannot supply. The call is never made.
        """
        return {
            "error": (
                "Lofted flanges are not available through this server: "
                "Models.AddLoftedFlangeEx needs 25 arguments including "
                "cross-section profiles, per-section origins and vertex maps. "
                "Create the lofted flange in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "lofted_flange_ex",
            "thickness": thickness,
        }

    @verifies_geometry
    def create_emboss(
        self,
        face_indices: list[int],
        clearance: float = 0.001,
        thickness: float = 0.0,
        thicken: bool = False,
        default_side: bool = True,
    ) -> dict[str, Any]:
        """
        Create an emboss feature using face geometry as tools.

        Uses selected faces as embossing tool geometry on the target body.
        Requires an existing base feature.

        Args:
            face_indices: List of 0-based face indices to use as emboss tools
            clearance: Clearance in meters (default 0.001)
            thickness: Thickness in meters (default 0.0)
            thicken: Enable thickening (default False)
            default_side: Default side (default True)

        Returns:
            Dict with status and emboss info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            body = model.Body

            if not face_indices:
                return {"error": "face_indices must contain at least one face index."}

            faces = body.Faces(FaceQueryConstants.igQueryAll)

            face_list = []
            for fi in face_indices:
                if fi < 0 or fi >= faces.Count:
                    return {"error": f"Invalid face index: {fi}. Body has {faces.Count} faces."}
                face_list.append(faces.Item(fi + 1))

            tools_arr = face_list

            emboss_features = model.EmbossFeatures
            emboss_features.Add(
                body, len(face_list), tools_arr, thicken, default_side, clearance, thickness
            )

            return {
                "status": "created",
                "type": "emboss",
                "face_count": len(face_list),
                "clearance": clearance,
                "thickness": thickness,
                "thicken": thicken,
            }
        except Exception as e:
            return error_result(e)

    @staticmethod
    def _flanges_unsupported(method: str, **echo: Any) -> dict[str, Any]:
        """The refusal every flange creator returns, with the evidence.

        Verified on Solid Edge 2026 against a base tab, on every horizontal
        edge, each on a fresh document: Flanges.Add, AddByMatchFace and
        AddByBendDeductionOrBendAllowance record a Flange feature (a name, a
        90-degree bend angle, a bend radius) that never solves -- the body's
        range, volume and face count are unchanged after Recompute,
        Flange.Status raises, and the UI draws only its outline. The optional
        parameters (ThicknessSide, InsideRadius, DimSide, BendAngle, and every
        FeaturePropertyConstants *Side* value) change the recorded feature and
        nothing else. AddSync and AddSyncByBendDeductionOrBendAllowance raise
        0x80004021 in an ordered document, and AddFlangeByFace raises
        E_POINTER. Refusing here leaves no dead feature behind.
        """
        return {
            "error": (
                f"{method}: the ordered Flanges.Add* calls record a flange that Solid "
                "Edge 2026 never solves (no geometry after Recompute), or raise. "
                "Flanges do build in a synchronous document: run "
                "manage_feature_tree(action='set_mode', mode='synchronous') before "
                "the base tab, then create_flange(method='sync')."
            ),
            "unsupported": True,
            "method": method,
            **echo,
        }

    @verifies_geometry
    def create_flange(
        self,
        face_index: int,
        edge_index: int,
        flange_length: float,
        side: str = "Right",
        inside_radius: float | None = None,
        bend_angle: float | None = None,
    ) -> dict[str, Any]:
        """
        Create a flange feature on a sheet metal edge.

        Adds a flange to the specified edge of a sheet metal body.
        Requires an existing sheet metal base feature.

        Args:
            face_index: 0-based face index containing the target edge
            edge_index: 0-based edge index within that face
            flange_length: Flange length in meters
            side: 'Left' (1), 'Right' (2), or 'Both' (6)
            inside_radius: Bend inside radius in meters (optional)
            bend_angle: Bend angle in degrees (optional)

        Returns:
            Dict with status and flange info
        """
        # The ordered Flanges.Add never solves (see _flanges_unsupported), but
        # Flanges.AddSync does in a synchronous document, so a caller who only
        # knows "make a flange" gets the working call when the document allows.
        doc = self.doc_manager.get_active_document()
        if self._require_synchronous_sheet(doc) is None:
            return self.create_flange_sync(
                face_index,
                edge_index,
                flange_length,
                inside_radius=inside_radius if inside_radius is not None else 0.001,
                bend_angle=bend_angle,
            )
        return self._flanges_unsupported(
            "create_flange",
            face_index=face_index,
            edge_index=edge_index,
            flange_length=flange_length,
            side=side,
            inside_radius=inside_radius,
            bend_angle=bend_angle,
        )

    @verifies_geometry
    def create_dimple(self, depth: float, direction: str = "Normal") -> dict[str, Any]:
        """
        Create a dimple feature (sheet metal).

        Creates a dimple from the active sketch profile on the sheet metal body.
        Requires an active sketch profile and an existing sheet metal base feature.

        Args:
            depth: Dimple depth in meters
            direction: 'Normal' or 'Reverse' for dimple direction

        Returns:
            Dict with status and dimple info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a sheet metal base feature first."}

            model = models.Item(1)

            # seDimpleDepthLeft=1, seDimpleDepthRight=2
            profile_side = 1 if direction == "Normal" else 2
            depth_side = 2 if direction == "Normal" else 1

            dimples = model.Dimples
            dimples.Add(profile, depth, profile_side, depth_side)

            return {"status": "created", "type": "dimple", "depth": depth, "direction": direction}
        except Exception as e:
            return error_result(e)

    @verifies_collection_growth("Models.*.Etches")
    def create_etch(self) -> dict[str, Any]:
        """
        Create an etch feature (sheet metal).

        Etches the active sketch profile into the sheet metal body.
        Requires an active sketch profile and an existing sheet metal base feature.

        Returns:
            Dict with status and etch info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a sheet metal base feature first."}

            model = models.Item(1)

            etches = model.Etches
            etches.Add(profile)

            return {"status": "created", "type": "etch"}
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_rib(self, thickness: float, direction: str = "Normal") -> dict[str, Any]:
        """
        Create a rib feature from the active sketch profile.

        Ribs are structural reinforcements that extend from a profile to
        existing geometry. Requires an active sketch profile.

        Uses Ribs.Add(RibProfile, ProfileExtensionType, ThicknessType,
        MaterialSide, ThicknessSide, Thickness, [FiniteDepth]).

        Args:
            thickness: Rib thickness in meters
            direction: 'Normal', 'Reverse', or 'Symmetric'

        Returns:
            Dict with status and rib info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            dir_map = {
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
                "Symmetric": DirectionConstants.igSymmetric,
            }
            side = dir_map.get(direction, DirectionConstants.igRight)

            ribs = model.Ribs
            ribs.Add(
                profile,
                _IG_EXTEND,  # ProfileExtensionType
                _IG_THK_NORMAL_TO_PROFILE_PLANE,  # ThicknessType
                side,  # MaterialSide
                DirectionConstants.igSymmetric,  # ThicknessSide
                thickness,
            )

            return {
                "status": "created",
                "type": "rib",
                "thickness": thickness,
                "direction": direction,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_lip(self, depth: float, direction: str = "Normal") -> dict[str, Any]:
        """
        Create a lip feature from the active sketch profile.

        NOT AVAILABLE via COM automation. Lips.Add takes (NumberOfEdges, Edges,
        SideFace, CapFace, Width, Height, [Type]) - it is driven by body edges
        plus a side face and a cap face, not by a sketch profile, and those
        selections cannot be made here. The call is never made.

        Args:
            depth: Lip depth/height in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "Lip features are not available through this server: Lips.Add "
                "needs the body edges to run the lip along plus a side face and "
                "a cap face, which cannot be selected here. Create the lip in "
                "the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "lip",
            "depth": depth,
            "direction": direction,
        }

    @verifies_geometry
    def create_drawn_cutout(self, depth: float, direction: str = "Normal") -> dict[str, Any]:
        """
        Create a drawn cutout feature (sheet metal).

        Creates a formed cutout from the active sketch profile. Unlike extruded
        cutouts, drawn cutouts follow the material's bend characteristics.
        Requires an active sketch profile and existing base feature.

        Uses DrawnCutouts.Add(Profile, Depth, ProfileSide, DepthSide,
        MaterialSide, [DieRadius], [TaperAngle], [ProfileCornerRadius],
        [RoundEdges], [RoundCorners]).

        Args:
            depth: Cutout depth in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            depth_side = (
                _SE_DRAWN_CUTOUT_DEPTH_RIGHT
                if direction == "Normal"
                else _SE_DRAWN_CUTOUT_DEPTH_LEFT
            )

            drawn_cutouts = model.DrawnCutouts
            drawn_cutouts.Add(
                profile,
                depth,
                _SE_DRAWN_CUTOUT_PROFILE_RIGHT,  # ProfileSide
                depth_side,  # DepthSide
                _SE_DRAWN_CUTOUT_MATERIAL_INSIDE,  # MaterialSide
            )

            return {
                "status": "created",
                "type": "drawn_cutout",
                "depth": depth,
                "direction": direction,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_bead(self, depth: float, direction: str = "Normal") -> dict[str, Any]:
        """
        Create a bead feature (sheet metal stiffener).

        NOT AVAILABLE via COM automation. Beads.Add takes thirteen arguments -
        nNumPathProfiles, ProfileArray, BeadType, BeadHeight, BeadWidth,
        BeadTaperAngle, BeadFormRadius, BeadPunchRadius, BeadDieRadius,
        BeadRoundOption, BeadSide, EndConditionType, EndPunchWidth. The bead
        cross-section is defined by eight of those values, none of which this
        tool receives, and inventing them would silently build the wrong shape.
        The call is never made.

        Args:
            depth: Bead depth in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "Bead features are not available through this server: Beads.Add "
                "needs a full bead cross-section (type, height, width, taper "
                "angle, form/punch/die radii, round option, end condition and "
                "end punch width) that this tool cannot supply. Create the bead "
                "in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "bead",
            "depth": depth,
            "direction": direction,
        }

    @verifies_geometry
    def create_louver(
        self, depth: float, direction: str = "Normal", height: float = 0.0
    ) -> dict[str, Any]:
        """
        Create a louver feature (sheet metal vent).

        Louvers are formed openings used for ventilation in sheet metal parts.
        Requires an active sketch profile and an existing sheet metal base feature.

        Uses Louvers.Add(Profile, Depth, DepthDirection, Height,
        HeightDirection, [Type], [RoundType], [DieRadius], [DimensionType]).

        Args:
            depth: Louver depth in meters
            direction: 'Normal' or 'Reverse'
            height: Louver height in meters (required by the COM call)

        Returns:
            Dict with status and louver info
        """
        # Louvers.Add records a Louver (Louver_1, in the Louvers collection)
        # that never solves: the body's range, volume and face count are
        # unchanged, with the line on the base plane or on a plane through
        # the top face, both depth directions, depths of 1 and 3 mm, and with
        # every Type / RoundType / DieRadius / DimensionType combination the
        # optional parameters take (Solid Edge 2026, nine placements). Say so
        # rather than leave a dead feature behind.
        return {
            "error": (
                "Louvers.Add records a louver that Solid Edge 2026 never solves "
                "(no geometry, whatever the placement or type), and Louvers.AddSync "
                "on a synchronous tab face answers 0x807B0086. Use the Solid Edge "
                "UI for louvers; create_dimple and create_drawn_cutout work."
            ),
            "unsupported": True,
            "depth": depth,
            "direction": direction,
            "height": height,
        }

    @verifies_geometry
    def create_gusset(self, thickness: float, direction: str = "Normal") -> dict[str, Any]:
        """
        Create a gusset feature (sheet metal reinforcement).

        Gussets are triangular reinforcement plates used in sheet metal.
        Requires an active sketch profile and an existing sheet metal base feature.

        Args:
            thickness: Gusset thickness in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and gusset info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            # igLeft=1, igRight=2
            side = 2 if direction == "Normal" else 1

            # Gussets has no Add. Part.tlb gives AddByProfile(pGussetProfile,
            # fpcMaterialSide, fpcThicknessSide, dGussetWidth, dTaperAngle,
            # gcGussetShape, dFormedRadius, gcRounding, dPunchRadius,
            # dDieRadius) and AddByBend; the profile form is the one a drawn
            # sketch fits. igUserDrawnProfile = 2, igRoundShape = 3,
            # igGussetNone = 0, from constant.tlb > GussetConstants.
            gussets = model.Gussets
            gussets.AddByProfile(
                profile,  # pGussetProfile
                side,  # fpcMaterialSide
                side,  # fpcThicknessSide
                thickness,  # dGussetWidth
                0.0,  # dTaperAngle
                2,  # gcGussetShape: igUserDrawnProfile
                0.0,  # dFormedRadius
                0,  # gcRounding: igGussetNone
                0.0,  # dPunchRadius
                0.0,  # dDieRadius
            )

            return {
                "status": "created",
                "type": "gusset",
                "thickness": thickness,
                "direction": direction,
            }
        except Exception as e:
            return error_result(e)

    def _find_cylinder_end_face(self, body: Any, cyl_face: Any) -> Any | None:
        """Find a face adjacent to a cylindrical face by shared edge topology.

        For holes, the cylindrical face shares circular edges with the planar
        faces at each end. Returns the first such adjacent face found.

        Args:
            body: The model body COM object
            cyl_face: The cylindrical face COM object

        Returns:
            The adjacent face COM object, or None if not found
        """
        cyl_edges = cyl_face.Edges
        if com_get(cyl_edges, "Count", 0) == 0:
            return None

        all_faces = body.Faces(FaceQueryConstants.igQueryAll)

        for fi in range(1, all_faces.Count + 1):
            candidate = all_faces.Item(fi)
            try:
                if candidate == cyl_face:
                    continue
            except Exception:
                pass

            try:
                cand_edges = candidate.Edges
                if com_get(cand_edges, "Count", 0) == 0:
                    continue
                for ej in range(1, cand_edges.Count + 1):
                    cand_edge = cand_edges.Item(ej)
                    for ci in range(1, cyl_edges.Count + 1):
                        cyl_edge = cyl_edges.Item(ci)
                        try:
                            if cand_edge == cyl_edge:
                                return candidate
                        except Exception:
                            pass
            except Exception:
                continue

        return None

    @verifies_collection_growth("Models.*.Threads")
    def create_thread(
        self,
        face_index: int,
        thread_diameter: float | None = None,
        thread_depth: float | None = None,
        physical: bool = False,
    ) -> dict[str, Any]:
        """
        Create a thread feature on a cylindrical face.

        Uses HoleData + Threads.Add/AddEx for proper thread creation.
        Automatically detects hole diameter from face geometry if not specified.

        Args:
            face_index: 0-based index of the cylindrical face to thread
            thread_diameter: Thread nominal diameter in meters (auto-detected if None)
            thread_depth: Thread depth in meters (full depth if None)
            physical: If True, creates modeled thread geometry via AddEx.
                      If False (default), creates a cosmetic thread via Add.

        Returns:
            Dict with status and thread info
        """
        # Threads.Add(HoleData, 1, [cylinder], [end face]) answers E_INVALIDARG
        # on Solid Edge 2026 for an extruded boss and for a cut hole alike, with
        # every HoleData this server can build: igTappedHole bare, with
        # ThreadMinorDiameter and ThreadDepth, igRegularThread with
        # ThreadExternalDiameter; a ThreadDescription ("M8") is refused by
        # HoleDataCollection.Add itself. A thread that comes with its hole
        # (create_hole(method='threaded')) is the route that works.
        return {
            "error": (
                "Threads.Add will not thread an existing cylindrical face on Solid "
                "Edge 2026 (E_INVALIDARG for a boss and a hole, every HoleData tried). "
                "Cut a threaded hole with create_hole(method='threaded') instead."
            ),
            "unsupported": True,
            "face_index": face_index,
            "thread_diameter": thread_diameter,
            "thread_depth": thread_depth,
            "physical": physical,
        }

    @verifies_geometry
    @verifies_geometry
    def create_slot(self, width: float, depth: float, direction: str = "Normal") -> dict[str, Any]:
        """
        Create a slot feature from the active sketch profile.

        NOT AVAILABLE via COM automation. Slots.Add takes 22 arguments, two of
        which are KeyPointOrTangentFace objects and two more of which are the
        From/To faces or planes of the extent. Those are user selections this
        server cannot make, so the call is never made.

        Args:
            depth: Slot depth in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with an unsupported error
        """
        if direction != "Normal":
            # Verified on Solid Edge 2026: with igLeft the slot is recorded and
            # cuts nothing (faces 6 -> 6), with igRight it cuts (6 -> 10).
            return {
                "error": (
                    "A slot cuts only with direction='Normal' on Solid Edge 2026; "
                    "'Reverse' records a slot that removes nothing. Draw the path on "
                    "the other face instead."
                ),
                "unsupported": True,
                "width": width,
                "depth": depth,
                "direction": direction,
            }
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}
            err = self._require_open_profile(profile, "slot")
            if err:
                return err
            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a sheet metal base feature first."}
            model = models.Item(1)
            # Slots.Add(Profile, SlotType, SlotEndCondition, SlotWidth, SlotOffsetWidth,
            #   SlotOffsetDepth, ExtentType, ExtentSide, FiniteDistance, KeyPointFlags,
            #   KeyPointOrTangentFace, ExtentType2, ExtentSide2, FiniteDistance2,
            #   KeyPointFlags2, KeyPointOrTangentFace2, FromFaceOrPlane, FromOffsetSide,
            #   FromOffsetDistance, ToFaceOrPlane, ToOffsetSide, ToOffsetDistance).
            # The keypoint and face slots take None for a finite extent; verified on
            # Solid Edge 2026 with a line path: 6 -> 10 faces.
            extent = (
                ExtentTypeConstants.igThroughAll if depth <= 0 else ExtentTypeConstants.igFinite
            )
            model.Slots.Add(
                profile,
                FeaturePropertyConstants.igRegularSlot,
                FeaturePropertyConstants.igNullConstant,
                width,
                0.0,
                0.0,
                extent,
                DirectionConstants.igRight,
                depth,
                0,
                None,
                ExtentTypeConstants.igNone,
                DirectionConstants.igRight,
                0.0,
                0,
                None,
                None,
                OffsetSideConstants.seOffsetNone,
                0.0,
                None,
                OffsetSideConstants.seOffsetNone,
                0.0,
            )
            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "slot",
                "width": width,
                "depth": depth,
                "direction": direction,
                "extent": "through_all" if depth <= 0 else "finite",
            }
        except Exception as e:
            return error_result(e)

    @verifies_collection_growth("Models.*.Splits")
    def create_split(self, plane_index: int = 1) -> dict[str, Any]:
        """
        Create a split feature to divide a body.

        NOT AVAILABLE via COM automation. Splits.Add takes (nNumTargets,
        TargetArray, nNumTools, ToolsArray, TargetDesignBodyOption,
        TargetConstructionBodyOption): the tools are the surfaces or planes that
        cut the target bodies, not the active sketch profile, and this server
        cannot select them. The call is never made.

        Args:
            direction: 'Normal' or 'Reverse' - which side to keep

        Returns:
            Dict with an unsupported error
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}
            model = models.Item(1)
            planes = doc.RefPlanes
            if plane_index < 1 or plane_index > planes.Count:
                return {
                    "error": (
                        f"Invalid plane index: {plane_index}. Document has {planes.Count} "
                        "reference planes (1-based)."
                    )
                }
            plane = planes.Item(plane_index)
            # Splits.Add(nNumTargets, TargetArray, nNumTools, ToolsArray,
            #   TargetDesignBodyOption, TargetConstructionBodyOption). Verified on
            # Solid Edge 2026: the model's own Body as the target and a reference
            # plane as the tool splits a box into two design bodies (Models 1 -> 2,
            # Splits 0 -> 1); the first model keeps its face count, so the
            # verification watches Splits.
            model.Splits.Add(
                1,
                [model.Body],
                1,
                [plane],
                SETargetDesignBodyOption.igCreateMultipleDesignBodiesOnNonManifoldOption,
                SETargetConstructionBodyOption.igCreateMultipleConstructionBodiesOnNonManifoldOption,
            )
            return {
                "status": "created",
                "type": "split",
                "plane_index": plane_index,
                "splits": com_get(model.Splits, "Count"),
                "models": com_get(doc.Models, "Count"),
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_flange_by_match_face(
        self,
        face_index: int,
        edge_index: int,
        flange_length: float,
        side: str = "Right",
        inside_radius: float = 0.001,
    ) -> dict[str, Any]:
        """
        Create a flange by matching an existing face edge.

        Uses Flanges.AddByMatchFace to add a flange that matches the
        geometry of a target face on the sheet metal body.

        Args:
            face_index: 0-based face index containing the target edge
            edge_index: 0-based edge index within that face
            flange_length: Flange length in meters
            side: 'Left' (1), 'Right' (2), or 'Both' (6)
            inside_radius: Bend inside radius in meters

        Returns:
            Dict with status and flange info
        """
        return self._flanges_unsupported(
            "create_flange_by_match_face",
            face_index=face_index,
            edge_index=edge_index,
            flange_length=flange_length,
            side=side,
            inside_radius=inside_radius,
        )

    @staticmethod
    def _require_synchronous_sheet(doc: Any) -> dict[str, Any] | None:
        """Flanges.AddSync builds only in a synchronous sheet-metal document.

        Verified on Solid Edge 2026: in a document set to synchronous before
        its base tab, AddSync on any horizontal tab edge builds (6 -> 14
        faces, the range grows by the flange length, InsideRadius and
        BendAngle are honoured). In an ordered document it raises 0x80004021,
        and switching the mode after an ordered tab answers E_FAIL -- the
        tab itself has to be synchronous, so this does not switch for you.
        """
        mode = com_get(doc, "ModelingMode")
        if mode is None or mode == ModelingModeConstants.seModelingModeSynchronous:
            return None
        return {
            "error": (
                "A synchronous flange needs a synchronous sheet-metal document whose "
                "base tab was made in that mode; switching an ordered tab afterwards "
                "answers E_FAIL on Solid Edge 2026. Create the document, run "
                "manage_feature_tree(action='set_mode', mode='synchronous') before "
                "create_sheet_metal_base, then create_flange(method='sync')."
            ),
            "modeling_mode": "ordered",
            "unsupported": True,
        }

    def _flange_edge(
        self, doc: Any, face_index: int, edge_index: int
    ) -> tuple[Any, Any] | dict[str, Any]:
        """The (model, edge) a flange is located on, or the error dict."""
        models = doc.Models
        if models.Count == 0:
            return {"error": "No base feature exists. Create a sheet metal base tab first."}
        model = models.Item(1)
        faces = model.Body.Faces(FaceQueryConstants.igQueryAll)
        if face_index < 0 or face_index >= faces.Count:
            return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}
        edges = faces.Item(face_index + 1).Edges
        if edge_index < 0 or edge_index >= edges.Count:
            return {"error": f"Invalid edge index: {edge_index}. Face has {edges.Count} edges."}
        return model, edges.Item(edge_index + 1)

    @verifies_geometry
    def create_flange_sync(
        self,
        face_index: int,
        edge_index: int,
        flange_length: float,
        inside_radius: float = 0.001,
        bend_angle: float | None = None,
    ) -> dict[str, Any]:
        """
        Create a synchronous flange feature.

        Uses Flanges.AddSync to add a flange in synchronous modeling mode.

        Args:
            face_index: 0-based face index containing the target edge
            edge_index: 0-based edge index within that face
            flange_length: Flange length in meters
            inside_radius: Bend inside radius in meters

        Returns:
            Dict with status and flange info
        """
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_synchronous_sheet(doc)
            if err:
                return err
            located = self._flange_edge(doc, face_index, edge_index)
            if isinstance(located, dict):
                return located
            model, edge = located
            # Flanges.AddSync(pLocatedEdge, FlangeLength, ThicknessSide, InsideRadius,
            #   DimSide, BRType, BRWidth, BRLength, CRType, NeutralFactor, BnParamType,
            #   BendAngle, FlangeType, PartialFlangeStartPoint). The VARIANT slots
            #   left empty take Solid Edge's defaults; verified with InsideRadius
            #   0.003 and BendAngle 45 degrees read back from the feature.
            empty = VARIANT(pythoncom.VT_EMPTY, None)
            angle = math.radians(bend_angle) if bend_angle is not None else empty
            model.Flanges.AddSync(
                edge,
                flange_length,
                empty,
                inside_radius,
                empty,
                empty,
                empty,
                empty,
                empty,
                empty,
                empty,
                angle,
            )
            return {
                "status": "created",
                "type": "flange_sync",
                "face_index": face_index,
                "edge_index": edge_index,
                "flange_length": flange_length,
                "inside_radius": inside_radius,
                "bend_angle": bend_angle,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_flange_by_face(
        self,
        face_index: int,
        edge_index: int,
        ref_face_index: int,
        flange_length: float,
        side: str = "Right",
        bend_radius: float = 0.001,
    ) -> dict[str, Any]:
        """
        Create a flange by face reference.

        Uses Flanges.AddFlangeByFace which references a target face
        for flange direction and orientation.

        Args:
            face_index: 0-based face index containing the target edge
            edge_index: 0-based edge index within that face
            ref_face_index: 0-based index of the reference face
            flange_length: Flange length in meters
            side: 'Left' (1), 'Right' (2), or 'Both' (6)
            bend_radius: Bend radius in meters

        Returns:
            Dict with status and flange info
        """
        return self._flanges_unsupported(
            "create_flange_by_face",
            face_index=face_index,
            edge_index=edge_index,
            ref_face_index=ref_face_index,
            flange_length=flange_length,
            side=side,
            bend_radius=bend_radius,
        )

    @verifies_geometry
    def create_flange_with_bend_calc(
        self,
        face_index: int,
        edge_index: int,
        flange_length: float,
        side: str = "Right",
        bend_deduction: float = 0.0,
    ) -> dict[str, Any]:
        """
        Create a flange with bend deduction/allowance calculation.

        Uses Flanges.AddByBendDeductionOrBendAllowance for precise
        bend calculation control.

        Args:
            face_index: 0-based face index containing the target edge
            edge_index: 0-based edge index within that face
            flange_length: Flange length in meters
            side: 'Left' (1), 'Right' (2), or 'Both' (6)
            bend_deduction: Bend deduction value in meters

        Returns:
            Dict with status and flange info
        """
        doc = self.doc_manager.get_active_document()
        if self._require_synchronous_sheet(doc) is None:
            return self.create_flange_sync_with_bend_calc(
                face_index, edge_index, flange_length, bend_deduction=bend_deduction
            )
        return self._flanges_unsupported(
            "create_flange_with_bend_calc",
            face_index=face_index,
            edge_index=edge_index,
            flange_length=flange_length,
            side=side,
            bend_deduction=bend_deduction,
        )

    @verifies_geometry
    def create_flange_sync_with_bend_calc(
        self,
        face_index: int,
        edge_index: int,
        flange_length: float,
        bend_deduction: float = 0.0,
    ) -> dict[str, Any]:
        """
        Create a synchronous flange with bend deduction/allowance.

        Uses Flanges.AddSyncByBendDeductionOrBendAllowance for
        synchronous mode with precise bend calculation.

        Args:
            face_index: 0-based face index containing the target edge
            edge_index: 0-based edge index within that face
            flange_length: Flange length in meters
            bend_deduction: Bend deduction value in meters

        Returns:
            Dict with status and flange info
        """
        if bend_deduction:
            return {
                "error": (
                    "A bend deduction through AddSyncByBendDeductionOrBendAllowance "
                    "(BendCalculationMethod, BendCalculationMethodValue) has not been "
                    "driven live; only the default calculation is. Pass bend_deduction=0 "
                    "or use create_flange(method='sync')."
                ),
                "unsupported": True,
                "bend_deduction": bend_deduction,
            }
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_synchronous_sheet(doc)
            if err:
                return err
            located = self._flange_edge(doc, face_index, edge_index)
            if isinstance(located, dict):
                return located
            model, edge = located
            # Verified on Solid Edge 2026: the two-argument form builds the same
            # flange AddSync does (6 -> 14 faces).
            model.Flanges.AddSyncByBendDeductionOrBendAllowance(edge, flange_length)
            return {
                "status": "created",
                "type": "flange_sync_with_bend_calc",
                "face_index": face_index,
                "edge_index": edge_index,
                "flange_length": flange_length,
                "bend_deduction": bend_deduction,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_contour_flange_ex(
        self,
        thickness: float,
        bend_radius: float = 0.001,
        direction: str = "Normal",
    ) -> dict[str, Any]:
        """
        Create an extended contour flange from the active profile.

        Uses ContourFlanges.AddEx to create a contour flange with
        keypoint/tangent face support.

        Args:
            thickness: Material projection distance in meters
            bend_radius: Bend radius in meters
            direction: 'Normal' or 'Reverse' for projection side

        Returns:
            Dict with status and contour flange info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a sheet metal base feature first."}

            del models

            # ContourFlanges.AddEx and Add answer E_FAIL to every open profile
            # this server can draw against a tab (Solid Edge 2026: lines from
            # the tab's corner and from the interior of an edge, on the base
            # planes and on planes perpendicular to the edge, both projection
            # sides, both APIs -- 52 combinations). The profile a contour
            # flange wants is attached to a thickness edge in the UI; nothing
            # drawn through ProfileSets satisfies it. Say so, without the call.
            return {
                "error": (
                    "ContourFlanges.AddEx rejects every open profile this server "
                    "can draw (E_FAIL on Solid Edge 2026). Use create_flange's sync "
                    "methods, create_lofted_flange, or the Solid Edge UI."
                ),
                "unsupported": True,
                "thickness": thickness,
                "bend_radius": bend_radius,
                "direction": direction,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_contour_flange_sync(
        self,
        face_index: int,
        edge_index: int,
        thickness: float,
        bend_radius: float = 0.001,
        direction: str = "Normal",
    ) -> dict[str, Any]:
        """
        Create a synchronous contour flange.

        Uses ContourFlanges.AddSync with a reference edge for
        synchronous modeling mode.

        Args:
            face_index: 0-based face index containing the reference edge
            edge_index: 0-based edge index within that face
            thickness: Material projection distance in meters
            bend_radius: Bend radius in meters
            direction: 'Normal' or 'Reverse' for projection side

        Returns:
            Dict with status and contour flange info
        """
        # ContourFlanges.AddSync(profile, edge, igFinite, side, distance, radius,
        # ...) answers E_INVALIDARG in a synchronous document with the open line
        # drawn from the tab edge on the perpendicular base plane, both
        # projection sides (Solid Edge 2026), as AddEx does in an ordered one.
        return {
            "error": (
                "ContourFlanges.AddSync rejects the open profile this server can draw "
                "(E_INVALIDARG on Solid Edge 2026, both sides). Use create_flange in a "
                "synchronous document, or the Solid Edge UI."
            ),
            "unsupported": True,
            "face_index": face_index,
            "edge_index": edge_index,
            "thickness": thickness,
            "bend_radius": bend_radius,
            "direction": direction,
        }

    @verifies_geometry
    def create_contour_flange_sync_with_bend(
        self,
        face_index: int,
        edge_index: int,
        thickness: float,
        bend_radius: float = 0.001,
        direction: str = "Normal",
        bend_deduction: float = 0.0,
    ) -> dict[str, Any]:
        """
        Create a synchronous contour flange with bend deduction/allowance.

        Uses ContourFlanges.AddSyncByBendDeductionOrBendAllowance for
        synchronous mode with precise bend calculation.

        Args:
            face_index: 0-based face index containing the reference edge
            edge_index: 0-based edge index within that face
            thickness: Material projection distance in meters
            bend_radius: Bend radius in meters
            direction: 'Normal' or 'Reverse' for projection side
            bend_deduction: Bend deduction value in meters

        Returns:
            Dict with status and contour flange info
        """
        try:
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            model, face, edge, err = self._get_edge_from_face(face_index, edge_index)
            if err:
                return err

            dir_side = (
                DirectionConstants.igRight if direction == "Normal" else DirectionConstants.igLeft
            )

            contour_flanges = model.ContourFlanges
            # AddSyncByBendDeductionOrBendAllowance(pProfile, pRefEdge,
            #   varExtentType, varProjectionSide, varProjectionDistance,
            #   varBendRadius, vtBRType, vtBRWidth, vtBRLength, vtCRType, ...)
            contour_flanges.AddSyncByBendDeductionOrBendAllowance(
                profile,
                edge,
                ExtentTypeConstants.igFinite,
                dir_side,
                thickness,
                bend_radius,
                0,  # vtBRType
                0.0,  # vtBRWidth
                0.0,  # vtBRLength
                0,  # vtCRType
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "contour_flange_sync_with_bend",
                "face_index": face_index,
                "edge_index": edge_index,
                "thickness": thickness,
                "bend_radius": bend_radius,
                "direction": direction,
                "bend_deduction": bend_deduction,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_hem(
        self,
        face_index: int,
        edge_index: int,
        hem_width: float = 0.005,
        bend_radius: float = 0.001,
        hem_type: str = "Closed",
    ) -> dict[str, Any]:
        """
        Create a hem feature on a sheet metal edge.

        Uses Hems.Add to fold an edge back on itself. Hem types include
        Closed, Open, S-Flange, Curl, etc.

        Args:
            face_index: 0-based face index containing the target edge
            edge_index: 0-based edge index within that face
            hem_width: Hem flange length in meters
            bend_radius: Bend radius in meters
            hem_type: 'Closed' (1), 'Open' (2), 'SFlange' (3), 'Curl' (4)

        Returns:
            Dict with status and hem info
        """
        try:
            model, face, edge, err = self._get_edge_from_face(face_index, edge_index)
            if err:
                return err

            # HemFeatureConstants
            hem_type_map = {
                "Closed": 1,  # seHemTypeClosed
                "Open": 2,  # seHemTypeOpen
                "SFlange": 3,  # seHemTypeSFlange
                "Curl": 4,  # seHemTypeCurl
                "OpenLoop": 5,  # seHemTypeOpenLoop
                "ClosedLoop": 6,  # seHemTypeClosedLoop
                "CenteredLoop": 7,  # seHemTypeCenteredLoop
            }
            hem_type_const = hem_type_map.get(hem_type, 1)

            hems = model.Hems
            # Add(InputEdge, HemType, BendRadius1, FlangeLength1, ...)
            hems.Add(edge, hem_type_const, bend_radius, hem_width)

            return {
                "status": "created",
                "type": "hem",
                "face_index": face_index,
                "edge_index": edge_index,
                "hem_width": hem_width,
                "bend_radius": bend_radius,
                "hem_type": hem_type,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_jog(
        self,
        jog_offset: float = 0.005,
        jog_angle: float = 90.0,
        direction: str = "Normal",
        moving_side: str = "Right",
    ) -> dict[str, Any]:
        """
        Create a jog feature on a sheet metal body.

        Uses Jogs.AddFinite with the active sketch profile to create
        a step/jog in the sheet metal.

        Args:
            jog_offset: Jog offset distance in meters
            jog_angle: Jog bend angle in degrees (converted to radians internally)
            direction: 'Normal' (16) or 'Reverse' (17) for jog direction
            moving_side: 'Right' (12) or 'Left' (11) for which side moves

        Returns:
            Dict with status and jog info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a sheet metal base feature first."}

            model = models.Item(1)

            # JogFeatureConstants
            # seJogMaterialInside=13, seJogMaterialOutside=14, seJogMaterialBendOutside=15
            material_side = 13  # seJogMaterialInside

            # seJogMoveLeft=11, seJogMoveRight=12
            move_side = 12 if moving_side == "Right" else 11

            # seJogNormal=16, seJogReverseNormal=17
            jog_dir = 16 if direction == "Normal" else 17

            jogs = model.Jogs
            # AddFinite(Profile, Extent, MaterialSide, MovingSide, JogDirection, ...)
            jogs.AddFinite(profile, jog_offset, material_side, move_side, jog_dir)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "jog",
                "jog_offset": jog_offset,
                "jog_angle": jog_angle,
                "direction": direction,
                "moving_side": moving_side,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_close_corner(
        self,
        face_index: int,
        edge_index: int,
        closure_type: str = "Close",
    ) -> dict[str, Any]:
        """
        Create a close corner feature on sheet metal.

        Uses CloseCorners.Add to close a gap between two flanges at
        a corner of the sheet metal body.

        Args:
            face_index: 0-based face index containing the corner edge
            edge_index: 0-based edge index at the corner
            closure_type: 'Close' (1) or 'Overlap' (2)

        Returns:
            Dict with status and close corner info
        """
        try:
            model, face, edge, err = self._get_edge_from_face(face_index, edge_index)
            if err:
                return err

            # CloseCornerFeatureConstants
            # seCloseCornerCloseFaces=1, seCloseCornerOverlapFaces=2
            closure_map = {
                "Close": 1,
                "Overlap": 2,
            }
            closure_const = closure_map.get(closure_type, 1)

            close_corners = model.CloseCorners
            # Add(InputEdge, ClosureType, ...)
            close_corners.Add(edge, closure_const)

            return {
                "status": "created",
                "type": "close_corner",
                "face_index": face_index,
                "edge_index": edge_index,
                "closure_type": closure_type,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_multi_edge_flange(
        self,
        face_index: int,
        edge_indices: list[int],
        flange_length: float,
        side: str = "Right",
    ) -> dict[str, Any]:
        """
        Create a multi-edge flange on multiple edges.

        Uses MultiEdgeFlanges.Add to create flanges on multiple edges
        simultaneously with consistent parameters.

        Args:
            face_index: 0-based face index containing the edges
            edge_indices: List of 0-based edge indices within that face
            flange_length: Flange length in meters
            side: 'Left' (1), 'Right' (2), or 'Both' (6)

        Returns:
            Dict with status and multi-edge flange info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a sheet metal base feature first."}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)
            face_edges = face.Edges
            if com_get(face_edges, "Count", 0) == 0:
                return {"error": f"Face {face_index} has no edges."}

            # Collect the requested edges
            edge_list = []
            for ei in edge_indices:
                if ei < 0 or ei >= face_edges.Count:
                    return {
                        "error": f"Invalid edge index: {ei}. Face has {face_edges.Count} edges."
                    }
                edge_list.append(face_edges.Item(ei + 1))

            side_map = {
                "Left": DirectionConstants.igLeft,
                "Right": DirectionConstants.igRight,
                "Both": DirectionConstants.igBoth,
            }
            side_const = side_map.get(side, DirectionConstants.igRight)

            edge_arr = edge_list

            multi_edge_flanges = model.MultiEdgeFlanges
            # Add(NumberOfEdges, Edges, FlangeSide, dFlangeLength, ...)
            multi_edge_flanges.Add(len(edge_list), edge_arr, side_const, flange_length)

            return {
                "status": "created",
                "type": "multi_edge_flange",
                "face_index": face_index,
                "edge_count": len(edge_list),
                "flange_length": flange_length,
                "side": side,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_bend_with_calc(
        self,
        bend_angle: float = 90.0,
        direction: str = "Normal",
        moving_side: str = "Right",
        bend_deduction: float = 0.0,
    ) -> dict[str, Any]:
        """
        Create a bend feature with bend deduction/allowance.

        Uses Bends.AddByBendDeductionOrBendAllowance to create a bend
        in the sheet metal using the active sketch profile as the bend line.

        Args:
            bend_angle: Bend angle in degrees (converted to radians)
            direction: 'Normal' (7) or 'Reverse' (8)
            moving_side: 'Right' (5) or 'Left' (6) for which side moves
            bend_deduction: Bend deduction value in meters

        Returns:
            Dict with status and bend info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a sheet metal base feature first."}

            err = self._require_open_profile(profile, "bend")
            if err:
                return err

            model = models.Item(1)

            bend_angle_rad = math.radians(bend_angle)

            # BendFeatureConstants
            # seBendPZLInside=11 (bend position zone line)
            bend_pzl = 11

            # seBendMoveRight=5, seBendMoveLeft=6
            move_side = 5 if moving_side == "Right" else 6

            # seBendNormal=7, seBendReverseNormal=8
            bend_dir = 7 if direction == "Normal" else 8

            bends = model.Bends
            # AddByBendDeductionOrBendAllowance(Profile, BendAngle, BendPZLSide,
            #   MovingSide, BendDirection, ...)
            bends.AddByBendDeductionOrBendAllowance(
                profile, bend_angle_rad, bend_pzl, move_side, bend_dir
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "bend_with_calc",
                "bend_angle": bend_angle,
                "direction": direction,
                "moving_side": moving_side,
                "bend_deduction": bend_deduction,
            }
        except Exception as e:
            return error_result(e)

    def convert_part_to_sheet_metal(
        self,
        thickness: float = 0.001,
    ) -> dict[str, Any]:
        """
        Convert a part document to sheet metal.

        Attempts to convert the current part document to a sheet metal
        document by saving it as a .psm file and reopening, or by using
        the Solid Edge command API.

        Args:
            thickness: Sheet metal thickness in meters

        Returns:
            Dict with status and conversion info
        """
        try:
            # Try using the StartCommand API to invoke the Convert to Sheet Metal command
            # This is a UI-level command approach
            app = self.doc_manager.connection_manager.get_application()

            with contextlib.suppress(Exception):
                # Try SE command for converting to sheet metal
                # Command ID for "Convert to Sheet Metal" may vary
                app.StartCommand(45000)  # seSheetMetalSelectCommand = 45000

                return {
                    "status": "command_invoked",
                    "type": "convert_part_to_sheet_metal",
                    "thickness": thickness,
                    "note": "Convert to Sheet Metal command invoked. "
                    "User interaction may be required to complete the conversion.",
                }

            # Alternative: Save as .psm and reopen
            return {
                "error": "Convert to sheet metal command not available. "
                "To create a sheet metal part, use create_sheet_metal_document() instead, "
                "or save the part as .psm format.",
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_dimple_ex(
        self,
        depth: float,
        direction: str = "Normal",
        punch_tool_diameter: float = 0.01,
    ) -> dict[str, Any]:
        """
        Create an extended dimple feature (sheet metal).

        Uses Dimples.AddEx with multi-profile support and additional
        parameters for punch tool diameter control.

        Args:
            depth: Dimple depth in meters
            direction: 'Normal' or 'Reverse' for dimple direction
            punch_tool_diameter: Punch tool diameter in meters

        Returns:
            Dict with status and dimple info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a sheet metal base feature first."}

            model = models.Item(1)

            # DimpleFeatureConstants
            # seDimpleProfileLeft=5 (inside), seDimpleProfileRight=6 (outside)
            profile_side = 5 if direction == "Normal" else 6
            # seDimpleDepthLeft=1, seDimpleDepthRight=2
            depth_side = 2 if direction == "Normal" else 1

            dimples = model.Dimples
            # AddEx(NumberOfProfiles, ProfileArray, Depth, ProfileSide, DepthSide,
            #   PunchRadius, ...)
            punch_radius = punch_tool_diameter / 2.0
            dimples.AddEx(1, (profile,), depth, profile_side, depth_side, punch_radius)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "dimple_ex",
                "depth": depth,
                "direction": direction,
                "punch_tool_diameter": punch_tool_diameter,
            }
        except Exception as e:
            return error_result(e)

    @verifies_collection_growth("Models.*.Threads")
    def create_thread_ex(
        self,
        face_index: int,
        thread_diameter: float | None = None,
        thread_depth: float | None = None,
    ) -> dict[str, Any]:
        """
        Create a physical (modeled) thread on a cylindrical face.

        Wrapper around create_thread with physical=True.
        Physical threads modify the actual body geometry unlike cosmetic threads.

        Args:
            face_index: 0-based index of the cylindrical face
            thread_diameter: Thread diameter in meters (auto-detected if None)
            thread_depth: Thread depth in meters (full depth if None)

        Returns:
            Dict with status and thread info
        """
        return self.create_thread(
            face_index=face_index,
            thread_diameter=thread_diameter,
            thread_depth=thread_depth,
            physical=True,
        )

    @verifies_geometry
    def create_slot_ex(
        self, width: float, depth: float, direction: str = "Normal"
    ) -> dict[str, Any]:
        """
        Create an extended slot feature with width and depth control.

        NOT AVAILABLE via COM automation. Slots.AddEx takes 18 arguments,
        including a KeyPointOrTangentFace object and the From/To faces or planes
        of the extent, which are user selections this server cannot make. The
        call is never made.

        Args:
            width: Slot width in meters
            depth: Slot depth in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "Slot features are not available through this server: "
                "Slots.AddEx requires 18 arguments including a "
                "KeyPointOrTangentFace object and From/To extent faces that "
                "cannot be selected here. Use an extruded cutout, or create the "
                "slot in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "slot_ex",
            "width": width,
            "depth": depth,
            "direction": direction,
        }

    @verifies_geometry
    def create_slot_sync(self, width: float, depth: float) -> dict[str, Any]:
        """
        Create a synchronous slot feature.

        NOT AVAILABLE via COM automation. Slots.AddSync takes the same 18
        arguments as Slots.AddEx, including a KeyPointOrTangentFace object and
        the From/To extent faces, which cannot be selected here. The call is
        never made.

        Args:
            width: Slot width in meters
            depth: Slot depth in meters

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "Slot features are not available through this server: "
                "Slots.AddSync requires 18 arguments including a "
                "KeyPointOrTangentFace object and From/To extent faces that "
                "cannot be selected here. Use an extruded cutout, or create the "
                "slot in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "slot_sync",
            "width": width,
            "depth": depth,
        }

    @verifies_geometry
    def create_drawn_cutout_ex(self, depth: float, direction: str = "Normal") -> dict[str, Any]:
        """
        Create an extended drawn cutout feature (sheet metal).

        Uses DrawnCutouts.AddEx(NumberOfProfiles, ProfileArray, Depth,
        ProfileSide, DepthSide, MaterialSide, [DieRadius], [TaperAngle],
        [ProfileCornerRadius], [RoundEdges], [RoundCorners]) with multi-profile
        support.

        Args:
            depth: Cutout depth in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()

            profiles = self.sketch_manager.get_accumulated_profiles()
            if not profiles:
                active = self.sketch_manager.get_active_sketch()
                profiles = [active] if active else []
            if not profiles:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            depth_side = (
                _SE_DRAWN_CUTOUT_DEPTH_RIGHT
                if direction == "Normal"
                else _SE_DRAWN_CUTOUT_DEPTH_LEFT
            )
            profile_arr = list(profiles)

            drawn_cutouts = model.DrawnCutouts
            drawn_cutouts.AddEx(
                len(profiles),
                profile_arr,
                depth,
                _SE_DRAWN_CUTOUT_PROFILE_RIGHT,  # ProfileSide
                depth_side,  # DepthSide
                _SE_DRAWN_CUTOUT_MATERIAL_INSIDE,  # MaterialSide
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "drawn_cutout_ex",
                "profile_count": len(profiles),
                "depth": depth,
                "direction": direction,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_louver_sync(self, depth: float) -> dict[str, Any]:
        """
        Create a synchronous louver feature (sheet metal).

        NOT AVAILABLE via COM automation. Louvers.AddSync takes (Face, Origin,
        Orientation, Length, Depth, Height, ...): it is placed on a picked face
        with explicit origin and orientation coordinate arrays rather than from
        a sketch profile, and this server cannot supply them. The call is never
        made.

        Args:
            depth: Louver depth in meters

        Returns:
            Dict with an unsupported error
        """
        return {
            "error": (
                "Synchronous louvers are not available through this server: "
                "Louvers.AddSync needs a target face plus origin and orientation "
                "coordinate arrays, which cannot be supplied here. Use "
                "create_louver(method='basic') with a sketch profile, or create "
                "the louver in the Solid Edge UI."
            ),
            "unsupported": True,
            "type": "louver_sync",
            "depth": depth,
        }

    @verifies_geometry
    def create_flange_match_face_with_bend(
        self,
        face_index: int,
        edge_index: int,
        flange_length: float,
        side: str = "Right",
        inside_radius: float = 0.001,
    ) -> dict[str, Any]:
        """
        Create a flange by match face with bend deduction/allowance.

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index on the face
            flange_length: Flange length in meters
            side: 'Right' or 'Left'
            inside_radius: Inside bend radius in meters

        Returns:
            Dict with status and flange info
        """
        return self._flanges_unsupported(
            "create_flange_match_face_with_bend",
            face_index=face_index,
            edge_index=edge_index,
            flange_length=flange_length,
            side=side,
            inside_radius=inside_radius,
        )

    @verifies_geometry
    def create_flange_by_face_with_bend(
        self,
        face_index: int,
        edge_index: int,
        ref_face_index: int,
        flange_length: float,
        side: str = "Right",
        bend_radius: float = 0.001,
    ) -> dict[str, Any]:
        """
        Create a flange by face reference with bend deduction/allowance.

        Args:
            face_index: 0-based face index containing the edge
            edge_index: 0-based edge index on the face
            ref_face_index: 0-based reference face index
            flange_length: Flange length in meters
            side: 'Right' or 'Left'
            bend_radius: Bend radius in meters

        Returns:
            Dict with status and flange info
        """
        return self._flanges_unsupported(
            "create_flange_by_face_with_bend",
            face_index=face_index,
            edge_index=edge_index,
            ref_face_index=ref_face_index,
            flange_length=flange_length,
            side=side,
            bend_radius=bend_radius,
        )

    @verifies_geometry
    def create_contour_flange_v3(
        self,
        thickness: float,
        bend_radius: float = 0.001,
        direction: str = "Normal",
    ) -> dict[str, Any]:
        """
        Create an extended v3 contour flange from the active profile.

        Uses ContourFlanges.Add3.

        Args:
            thickness: Material thickness in meters
            bend_radius: Bend radius in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and contour flange info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            dir_const = (
                DirectionConstants.igRight if direction == "Normal" else DirectionConstants.igLeft
            )

            contour_flanges = model.ContourFlanges
            contour_flanges.Add3(
                profile,
                ExtentTypeConstants.igFinite,
                dir_const,
                thickness,
                None,
                0,
                bend_radius,
                0,
                0.001,
                0.001,
                0,
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "contour_flange_v3",
                "thickness": thickness,
                "bend_radius": bend_radius,
                "direction": direction,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_contour_flange_sync_ex(
        self,
        face_index: int,
        edge_index: int,
        thickness: float,
        bend_radius: float = 0.001,
        direction: str = "Normal",
    ) -> dict[str, Any]:
        """
        Create a synchronous extended contour flange.

        Uses ContourFlanges.AddSyncEx.

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index on the face
            thickness: Material thickness in meters
            bend_radius: Bend radius in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and contour flange info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)

            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face_index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)
            edges = face.Edges
            if edge_index < 0 or edge_index >= edges.Count:
                return {"error": f"Invalid edge_index: {edge_index}. Face has {edges.Count} edges."}

            edge = edges.Item(edge_index + 1)
            dir_const = (
                DirectionConstants.igRight if direction == "Normal" else DirectionConstants.igLeft
            )

            contour_flanges = model.ContourFlanges
            contour_flanges.AddSyncEx(
                profile,
                edge,
                ExtentTypeConstants.igFinite,
                dir_const,
                thickness,
                bend_radius,
                0,
                0.001,
                0.001,
                0,
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "contour_flange_sync_ex",
                "face_index": face_index,
                "edge_index": edge_index,
                "thickness": thickness,
                "bend_radius": bend_radius,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_bend(
        self,
        bend_angle: float = 90.0,
        direction: str = "Normal",
        moving_side: str = "Right",
        bend_radius: float = 0.001,
    ) -> dict[str, Any]:
        """
        Create a basic bend feature from the active profile.

        Uses Bends.Add(Profile, BendAngle, BendPZLSide, MovingSide,
        BendDirection, BendRadius).

        Args:
            bend_angle: Bend angle in degrees
            direction: 'Normal' or 'Reverse'
            moving_side: 'Right' or 'Left'
            bend_radius: Bend radius in meters

        Returns:
            Dict with status and bend info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            err = self._require_open_profile(profile, "bend")
            if err:
                return err

            model = models.Item(1)
            angle_rad = math.radians(bend_angle)

            # BendFeatureConstants: seBendNormal=7, seBendReverseNormal=8
            # seBendMoveRight=5, seBendMoveLeft=6, seBendPZLInside=11
            dir_const = 7 if direction == "Normal" else 8
            move_const = 5 if moving_side == "Right" else 6

            bends = model.Bends
            bends.Add(profile, angle_rad, 11, move_const, dir_const, bend_radius)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "bend",
                "bend_angle": bend_angle,
                "direction": direction,
                "moving_side": moving_side,
                "bend_radius": bend_radius,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_slot_multi_body(
        self, width: float, depth: float, direction: str = "Normal"
    ) -> dict[str, Any]:
        """
        Create a slot feature that spans multiple bodies.

        Uses Slots.AddMultiBody.

        Args:
            width: Slot width in meters
            depth: Slot depth in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and slot info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            dir_const = (
                DirectionConstants.igRight if direction == "Normal" else DirectionConstants.igLeft
            )

            body = model.Body
            body_arr = [body]

            slots = model.Slots
            slots.AddMultiBody(
                1,
                (profile,),
                dir_const,
                ExtentTypeConstants.igFinite,
                width,
                0.0,
                0.0,
                ExtentTypeConstants.igFinite,
                dir_const,
                depth,
                KeyPointExtentConstants.igTangentNormal,
                None,
                None,
                OffsetSideConstants.seOffsetNone,
                0.0,
                None,
                OffsetSideConstants.seOffsetNone,
                0.0,
                1,
                body_arr,
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "slot_multi_body",
                "width": width,
                "depth": depth,
                "direction": direction,
            }
        except Exception as e:
            return error_result(e)

    @verifies_geometry
    def create_slot_sync_multi_body(
        self, width: float, depth: float, direction: str = "Normal"
    ) -> dict[str, Any]:
        """
        Create a synchronous slot feature that spans multiple bodies.

        Uses Slots.AddSyncMultiBody.

        Args:
            width: Slot width in meters
            depth: Slot depth in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and slot info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            dir_const = (
                DirectionConstants.igRight if direction == "Normal" else DirectionConstants.igLeft
            )

            body = model.Body
            body_arr = [body]

            slots = model.Slots
            slots.AddSyncMultiBody(
                1,
                (profile,),
                dir_const,
                ExtentTypeConstants.igFinite,
                width,
                0.0,
                0.0,
                ExtentTypeConstants.igFinite,
                dir_const,
                depth,
                KeyPointExtentConstants.igTangentNormal,
                None,
                None,
                OffsetSideConstants.seOffsetNone,
                0.0,
                None,
                OffsetSideConstants.seOffsetNone,
                0.0,
                1,
                body_arr,
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "slot_sync_multi_body",
                "width": width,
                "depth": depth,
                "direction": direction,
            }
        except Exception as e:
            return error_result(e)
