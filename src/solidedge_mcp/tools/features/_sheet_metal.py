"""Sheet metal tools."""

from typing import Any, Literal

from solidedge_mcp.backends.validation import validate_numerics
from solidedge_mcp.managers import feature_manager


def create_flange(
    method: Literal[
        "basic",
        "by_match_face",
        "sync",
        "by_face",
        "with_bend_calc",
        "sync_with_bend_calc",
        "match_face_with_bend",
        "by_face_with_bend",
    ] = "basic",
    face_index: int = 0,
    edge_index: int = 0,
    flange_length: float = 0.0,
    side: Literal["Left", "Right", "Both"] = "Right",
    inside_radius: float | None = None,
    bend_angle: float | None = None,
    bend_deduction: float = 0.0,
    ref_face_index: int = 0,
    bend_radius: float = 0.001,
) -> dict[str, Any]:
    """Create a flange on an edge of a sheet metal part.

    Lengths and radii in meters; bend_angle in degrees. face_index,
    edge_index (within that face) and ref_face_index are 0-based.
    inside_radius: basic / by_match_face / sync / match_face_with_bend
    (defaults to 0.001 when omitted). bend_angle: basic only.
    bend_deduction: *_bend_calc methods. ref_face_index and bend_radius:
    by_face / by_face_with_bend.
    """
    err = validate_numerics(
        flange_length=flange_length,
        bend_deduction=bend_deduction,
        bend_radius=bend_radius,
    )
    if err:
        return err
    match method:
        case "basic":
            return feature_manager.create_flange(
                face_index=face_index,
                edge_index=edge_index,
                flange_length=flange_length,
                side=side,
                inside_radius=inside_radius,
                bend_angle=bend_angle,
            )
        case "by_match_face":
            return feature_manager.create_flange_by_match_face(
                face_index=face_index,
                edge_index=edge_index,
                flange_length=flange_length,
                side=side,
                inside_radius=inside_radius or 0.001,
            )
        case "sync":
            return feature_manager.create_flange_sync(
                face_index=face_index,
                edge_index=edge_index,
                flange_length=flange_length,
                inside_radius=inside_radius or 0.001,
                bend_angle=bend_angle,
            )
        case "by_face":
            return feature_manager.create_flange_by_face(
                face_index=face_index,
                edge_index=edge_index,
                ref_face_index=ref_face_index,
                flange_length=flange_length,
                side=side,
                bend_radius=bend_radius,
            )
        case "with_bend_calc":
            return feature_manager.create_flange_with_bend_calc(
                face_index=face_index,
                edge_index=edge_index,
                flange_length=flange_length,
                side=side,
                bend_deduction=bend_deduction,
            )
        case "sync_with_bend_calc":
            return feature_manager.create_flange_sync_with_bend_calc(
                face_index=face_index,
                edge_index=edge_index,
                flange_length=flange_length,
                bend_deduction=bend_deduction,
            )
        case "match_face_with_bend":
            return feature_manager.create_flange_match_face_with_bend(
                face_index=face_index,
                edge_index=edge_index,
                flange_length=flange_length,
                side=side,
                inside_radius=inside_radius or 0.001,
            )
        case "by_face_with_bend":
            return feature_manager.create_flange_by_face_with_bend(
                face_index=face_index,
                edge_index=edge_index,
                ref_face_index=ref_face_index,
                flange_length=flange_length,
                side=side,
                bend_radius=bend_radius,
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def create_contour_flange(
    method: Literal["ex", "sync", "sync_with_bend", "v3", "sync_ex"] = "ex",
    thickness: float = 0.0,
    bend_radius: float = 0.001,
    direction: Literal["Normal", "Reverse"] = "Normal",
    face_index: int = 0,
    edge_index: int = 0,
    bend_deduction: float = 0.0,
) -> dict[str, Any]:
    """Sweep the active sketch profile into a contour flange (sheet metal).

    Thickness, radii and deduction in meters. direction is the side the
    material is projected to. face_index and edge_index (0-based) select the
    attachment edge and apply to the sync* methods only. bend_deduction:
    sync_with_bend only.
    """
    err = validate_numerics(
        thickness=thickness,
        bend_radius=bend_radius,
        bend_deduction=bend_deduction,
    )
    if err:
        return err
    match method:
        case "ex":
            return feature_manager.create_contour_flange_ex(
                thickness=thickness, bend_radius=bend_radius, direction=direction
            )
        case "sync":
            return feature_manager.create_contour_flange_sync(
                face_index=face_index,
                edge_index=edge_index,
                thickness=thickness,
                bend_radius=bend_radius,
                direction=direction,
            )
        case "sync_with_bend":
            return feature_manager.create_contour_flange_sync_with_bend(
                face_index=face_index,
                edge_index=edge_index,
                thickness=thickness,
                bend_radius=bend_radius,
                direction=direction,
                bend_deduction=bend_deduction,
            )
        case "v3":
            return feature_manager.create_contour_flange_v3(
                thickness=thickness, bend_radius=bend_radius, direction=direction
            )
        case "sync_ex":
            return feature_manager.create_contour_flange_sync_ex(
                face_index=face_index,
                edge_index=edge_index,
                thickness=thickness,
                bend_radius=bend_radius,
                direction=direction,
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def create_sheet_metal_base(
    type: Literal["flange", "tab", "contour_advanced", "tab_multi_profile"] = "flange",
    thickness: float = 0.0,
    width: float | None = None,
    bend_radius: float | None = None,
    relief_type: str = "Default",
) -> dict[str, Any]:
    """Create the base feature of a sheet metal part from the active sketch.

    All dimensions in meters. thickness applies to every type.
    width: flange (defaults to 0.0), tab, and contour_advanced, where it is
    the flange projection distance and must be > 0. bend_radius: flange and
    contour_advanced (defaults to 0.001). relief_type is accepted for
    contour_advanced but the COM call does not currently consume it.
    """
    err = validate_numerics(thickness=thickness, width=width)
    if err:
        return err
    match type:
        case "flange":
            return feature_manager.create_base_flange(
                width=width or 0.0, thickness=thickness, bend_radius=bend_radius
            )
        case "tab":
            return feature_manager.create_base_tab(thickness=thickness, width=width)
        case "contour_advanced":
            return feature_manager.create_base_contour_flange_advanced(
                thickness=thickness,
                bend_radius=bend_radius or 0.001,
                relief_type=relief_type,
                width=width or 0.0,
            )
        case "tab_multi_profile":
            return feature_manager.create_base_tab_multi_profile(thickness=thickness)
        case _:
            return {"error": f"Unknown type: {type}"}


def create_lofted_flange(
    method: Literal["basic", "advanced", "ex"] = "basic",
    thickness: float = 0.0,
    bend_radius: float = 0.0,
) -> dict[str, Any]:
    """Create a lofted flange between two sketch profiles (sheet metal).

    Dimensions in meters. bend_radius applies to 'advanced' only, which uses
    bend deduction/allowance; 'basic' and 'ex' take thickness alone.
    Every method is unsupported: Models.AddLoftedFlange* needs cross-section
    profiles with a per-section origin, origin reference and vertex map that
    cannot be built here. Create the lofted flange in the Solid Edge UI.
    """
    err = validate_numerics(thickness=thickness, bend_radius=bend_radius)
    if err:
        return err
    match method:
        case "basic":
            return feature_manager.create_lofted_flange(thickness=thickness)
        case "advanced":
            return feature_manager.create_lofted_flange_advanced(
                thickness=thickness, bend_radius=bend_radius
            )
        case "ex":
            return feature_manager.create_lofted_flange_ex(thickness=thickness)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_bend(
    method: Literal["basic", "with_calc"] = "basic",
    bend_angle: float = 90.0,
    direction: Literal["Normal", "Reverse"] = "Normal",
    moving_side: Literal["Left", "Right"] = "Right",
    bend_deduction: float = 0.0,
) -> dict[str, Any]:
    """Bend a sheet metal face along the active sketch line.

    The sketch must be an open line across the face: a closed shape is
    refused. bend_angle in degrees; bend_deduction in meters. direction is the
    side the material folds toward; moving_side selects which side of the bend
    line moves. bend_deduction applies to 'with_calc' only.
    """
    err = validate_numerics(bend_angle=bend_angle, bend_deduction=bend_deduction)
    if err:
        return err
    match method:
        case "basic":
            return feature_manager.create_bend(
                bend_angle=bend_angle, direction=direction, moving_side=moving_side
            )
        case "with_calc":
            return feature_manager.create_bend_with_calc(
                bend_angle=bend_angle,
                direction=direction,
                moving_side=moving_side,
                bend_deduction=bend_deduction,
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def create_slot(
    method: Literal["basic", "ex", "sync", "multi_body", "sync_multi_body"] = "basic",
    width: float = 0.0,
    depth: float = 0.0,
    direction: Literal["Normal", "Reverse"] = "Normal",
) -> dict[str, Any]:
    """Create a slot from the active sketch profile.

    width and depth in meters (depth <= 0 cuts through all). 'basic' cuts a slot
    along an OPEN line path with direction='Normal' (Slots.Add; 'Reverse' records
    a slot that removes nothing on Solid Edge 2026). 'ex' is unsupported;
    'sync', 'multi_body' and 'sync_multi_body' take width + depth + direction.
    """
    err = validate_numerics(width=width, depth=depth)
    if err:
        return err
    match method:
        case "basic":
            return feature_manager.create_slot(width=width, depth=depth, direction=direction)
        case "ex":
            return feature_manager.create_slot_ex(width=width, depth=depth, direction=direction)
        case "sync":
            return feature_manager.create_slot_sync(width=width, depth=depth)
        case "multi_body":
            return feature_manager.create_slot_multi_body(
                width=width, depth=depth, direction=direction
            )
        case "sync_multi_body":
            return feature_manager.create_slot_sync_multi_body(
                width=width, depth=depth, direction=direction
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def create_thread(
    method: Literal["basic", "physical"] = "basic",
    face_index: int = 0,
    thread_diameter: float = 0.0,
    thread_depth: float = 0.0,
) -> dict[str, Any]:
    """Create a thread on a cylindrical face.

    'basic' is cosmetic; 'physical' cuts real geometry. face_index is 0-based.
    Diameter and depth in meters; leave at 0 to auto-detect from the face.
    """
    err = validate_numerics(thread_diameter=thread_diameter, thread_depth=thread_depth)
    if err:
        return err
    diameter = thread_diameter if thread_diameter > 0 else None
    depth = thread_depth if thread_depth > 0 else None

    match method:
        case "basic":
            return feature_manager.create_thread(
                face_index=face_index, thread_diameter=diameter, thread_depth=depth
            )
        case "physical":
            return feature_manager.create_thread_ex(
                face_index=face_index, thread_diameter=diameter, thread_depth=depth
            )
        case _:
            return {"error": f"Unknown method: {method}. Use 'basic' or 'physical'."}


def create_drawn_cutout(
    method: Literal["basic", "ex"] = "basic",
    depth: float = 0.0,
    direction: Literal["Normal", "Reverse"] = "Normal",
) -> dict[str, Any]:
    """Create a drawn cutout from the active sketch profile (sheet metal).

    depth in meters. direction is the side the material is drawn toward.
    'ex' uses the extended COM overload; both take the same parameters.
    """
    err = validate_numerics(depth=depth)
    if err:
        return err
    match method:
        case "basic":
            return feature_manager.create_drawn_cutout(depth=depth, direction=direction)
        case "ex":
            return feature_manager.create_drawn_cutout_ex(depth=depth, direction=direction)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_dimple(
    method: Literal["basic", "ex"] = "basic",
    depth: float = 0.0,
    direction: Literal["Normal", "Reverse"] = "Normal",
    punch_tool_diameter: float = 0.01,
) -> dict[str, Any]:
    """Create a dimple from the active sketch profile (sheet metal).

    depth and punch_tool_diameter in meters. direction is the side the
    material is pushed toward. punch_tool_diameter applies to 'ex' only.
    """
    err = validate_numerics(depth=depth, punch_tool_diameter=punch_tool_diameter)
    if err:
        return err
    match method:
        case "basic":
            return feature_manager.create_dimple(depth=depth, direction=direction)
        case "ex":
            return feature_manager.create_dimple_ex(
                depth=depth, direction=direction, punch_tool_diameter=punch_tool_diameter
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def create_louver(
    method: Literal["basic", "sync"] = "basic",
    depth: float = 0.0,
    height: float = 0.0,
    direction: Literal["Normal", "Reverse"] = "Normal",
) -> dict[str, Any]:
    """Create a louver from the active sketch profile (sheet metal).

    The sketch must be an open line where the louver runs: a closed shape is
    refused. depth and height in meters; Louvers.Add requires a positive height, so
    'basic' fails without it. direction is the side the material is formed
    toward. 'sync' (unsupported: Louvers.AddSync needs a target face plus
    origin and orientation coordinate arrays that cannot be supplied here);
    use 'basic'.
    """
    err = validate_numerics(depth=depth, height=height)
    if err:
        return err
    match method:
        case "basic":
            return feature_manager.create_louver(depth=depth, direction=direction, height=height)
        case "sync":
            return feature_manager.create_louver_sync(depth=depth)
        case _:
            return {"error": f"Unknown method: {method}"}


def sheet_metal_misc(
    action: Literal["hem", "jog", "close_corner", "multi_edge_flange", "convert"] = "hem",
    face_index: int = 0,
    edge_index: int = 0,
    hem_width: float = 0.005,
    bend_radius: float = 0.001,
    hem_type: Literal[
        "Closed", "Open", "SFlange", "Curl", "OpenLoop", "ClosedLoop", "CenteredLoop"
    ] = "Closed",
    jog_offset: float = 0.005,
    jog_angle: float = 90.0,
    direction: Literal["Normal", "Reverse"] = "Normal",
    moving_side: Literal["Left", "Right"] = "Right",
    closure_type: Literal["Close", "Overlap"] = "Close",
    edge_indices: list[int] | None = None,
    flange_length: float = 0.0,
    side: Literal["Left", "Right", "Both"] = "Right",
    thickness: float = 0.001,
) -> dict[str, Any]:
    """Hems, jogs, corner closures, multi-edge flanges, and part conversion.

    Lengths in meters, jog_angle in degrees; face/edge indices are 0-based.
    hem: face_index, edge_index, hem_width, bend_radius, hem_type.
    jog: jog_offset, jog_angle, direction, moving_side (uses active sketch).
    close_corner: face_index, edge_index, closure_type.
    multi_edge_flange: face_index, edge_indices, flange_length, side.
    convert: turns the solid part into sheet metal of the given thickness.
    """
    err = validate_numerics(
        hem_width=hem_width,
        bend_radius=bend_radius,
        jog_offset=jog_offset,
        jog_angle=jog_angle,
        flange_length=flange_length,
        thickness=thickness,
    )
    if err:
        return err
    match action:
        case "hem":
            return feature_manager.create_hem(
                face_index=face_index,
                edge_index=edge_index,
                hem_width=hem_width,
                bend_radius=bend_radius,
                hem_type=hem_type,
            )
        case "jog":
            return feature_manager.create_jog(
                jog_offset=jog_offset,
                jog_angle=jog_angle,
                direction=direction,
                moving_side=moving_side,
            )
        case "close_corner":
            return feature_manager.create_close_corner(
                face_index=face_index, edge_index=edge_index, closure_type=closure_type
            )
        case "multi_edge_flange":
            return feature_manager.create_multi_edge_flange(
                face_index=face_index,
                edge_indices=edge_indices or [],
                flange_length=flange_length,
                side=side,
            )
        case "convert":
            return feature_manager.convert_part_to_sheet_metal(thickness=thickness)
        case _:
            return {"error": f"Unknown action: {action}"}


def create_stamped(
    type: Literal["bead", "gusset"] = "bead",
    depth: float = 0.0,
) -> dict[str, Any]:
    """Create a stamped feature from the active sketch: a bead or a gusset.

    depth in meters (the gusset uses it as the material thickness).
    'bead' (unsupported: Beads.Add needs a full bead cross-section - type,
    height, width, taper angle, form/punch/die radii, end condition - that
    this tool cannot supply); use 'gusset', or add the bead in the UI.
    """
    err = validate_numerics(depth=depth)
    if err:
        return err
    match type:
        case "bead":
            return feature_manager.create_bead(depth=depth)
        case "gusset":
            # The backend parameter is the plate thickness; this tool only
            # exposes `depth`, which is what Solid Edge uses for it here.
            return feature_manager.create_gusset(thickness=depth)
        case _:
            return {"error": f"Unknown type: {type}"}


def create_surface_mark(
    type: Literal["emboss", "etch"] = "emboss",
    face_indices: list[int] | None = None,
    clearance: float = 0.001,
    thickness: float = 0.0,
    thicken: bool = False,
    default_side: bool = True,
) -> dict[str, Any]:
    """Emboss or etch the active sketch profile onto the body.

    clearance and thickness in meters; face_indices are 0-based. All
    parameters other than the discriminator apply to 'emboss'; 'etch' takes
    the active profile alone.
    """
    err = validate_numerics(clearance=clearance, thickness=thickness)
    if err:
        return err
    match type:
        case "emboss":
            return feature_manager.create_emboss(
                face_indices=face_indices or [],
                clearance=clearance,
                thickness=thickness,
                thicken=thicken,
                default_side=default_side,
            )
        case "etch":
            return feature_manager.create_etch()
        case _:
            return {"error": f"Unknown type: {type}"}


def create_reinforcement(
    type: Literal["rib", "lip"] = "rib",
    thickness: float = 0.0,
    direction: Literal["Normal", "Reverse", "Symmetric"] = "Normal",
) -> dict[str, Any]:
    """Create a reinforcement from the active sketch: a rib or a lip.

    thickness in meters (the lip uses it as the lip depth). direction is the
    side material is added to; 'Symmetric' is accepted by 'rib' only.
    'lip' (unsupported: Lips.Add needs the body edges to run the lip along
    plus a side face and a cap face, which cannot be selected here); use
    'rib', or add the lip in the Solid Edge UI.
    """
    err = validate_numerics(thickness=thickness)
    if err:
        return err
    match type:
        case "rib":
            return feature_manager.create_rib(thickness=thickness, direction=direction)
        case "lip":
            return feature_manager.create_lip(depth=thickness)
        case _:
            return {"error": f"Unknown type: {type}"}


def create_web_network(
    thickness: float = 0.0,
    depth: float = 0.0,
    direction: Literal["Normal", "Reverse", "Symmetric"] = "Normal",
) -> dict[str, Any]:
    """Create a web network from the active sketch (sheet metal / plastic part).

    The web geometry comes from the open sketch profiles. thickness (the web
    thickness, required and must be > 0) and depth (the finite web depth) are
    in meters; direction is the side material is added to.
    """
    err = validate_numerics(thickness=thickness, depth=depth)
    if err:
        return err
    return feature_manager.create_web_network(thickness=thickness, depth=depth, direction=direction)


def create_split(plane_index: int = 1) -> dict[str, Any]:
    """Split the solid body with a reference plane into two design bodies.

    plane_index is 1-based (1=Top, 2=Right, 3=Front, 4+ = planes made with
    create_ref_plane); the plane has to pass through the body, so it is
    usually an offset plane. Reports the Splits and Models counts afterwards.
    """
    return feature_manager.create_split(plane_index=plane_index)
