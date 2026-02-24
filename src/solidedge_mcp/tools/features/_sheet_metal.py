"""Sheet metal tools."""

from typing import Any

from solidedge_mcp.backends.validation import validate_numerics
from solidedge_mcp.managers import feature_manager


def create_flange(
    method: str = "basic",
    face_index: int = 0,
    edge_index: int = 0,
    flange_length: float = 0.0,
    side: str = "Right",
    inside_radius: float | None = None,
    bend_angle: float | None = None,
    bend_deduction: float = 0.0,
    ref_face_index: int = 0,
    bend_radius: float = 0.001,
) -> dict[str, Any]:
    """Create a flange on an edge (sheet metal).

    method: 'basic' | 'by_match_face' | 'sync' | 'by_face'
        | 'with_bend_calc' | 'sync_with_bend_calc'
        | 'match_face_with_bend' | 'by_face_with_bend'

    Lengths/radii in meters. bend_angle in degrees.
    """
    err = validate_numerics(
        flange_length=flange_length, bend_deduction=bend_deduction,
        bend_radius=bend_radius,
    )
    if err:
        return err
    match method:
        case "basic":
            return feature_manager.create_flange(
                face_index,
                edge_index,
                flange_length,
                side,
                inside_radius,
                bend_angle,
            )
        case "by_match_face":
            return feature_manager.create_flange_by_match_face(
                face_index,
                edge_index,
                flange_length,
                side,
                inside_radius or 0.001,
            )
        case "sync":
            return feature_manager.create_flange_sync(
                face_index,
                edge_index,
                flange_length,
                inside_radius or 0.001,
            )
        case "by_face":
            return feature_manager.create_flange_by_face(
                face_index,
                edge_index,
                ref_face_index,
                flange_length,
                side,
                bend_radius,
            )
        case "with_bend_calc":
            return feature_manager.create_flange_with_bend_calc(
                face_index,
                edge_index,
                flange_length,
                side,
                bend_deduction,
            )
        case "sync_with_bend_calc":
            return feature_manager.create_flange_sync_with_bend_calc(
                face_index,
                edge_index,
                flange_length,
                bend_deduction,
            )
        case "match_face_with_bend":
            return feature_manager.create_flange_match_face_with_bend(
                face_index,
                edge_index,
                flange_length,
                side,
                inside_radius or 0.001,
            )
        case "by_face_with_bend":
            return feature_manager.create_flange_by_face_with_bend(
                face_index,
                edge_index,
                ref_face_index,
                flange_length,
                side,
                bend_radius,
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def create_contour_flange(
    method: str = "ex",
    thickness: float = 0.0,
    bend_radius: float = 0.001,
    direction: str = "Normal",
    face_index: int = 0,
    edge_index: int = 0,
    bend_deduction: float = 0.0,
) -> dict[str, Any]:
    """Create a contour flange (sheet metal).

    method: 'ex' | 'sync' | 'sync_with_bend' | 'v3'
        | 'sync_ex'

    Dimensions in meters.
    """
    err = validate_numerics(
        thickness=thickness, bend_radius=bend_radius,
        bend_deduction=bend_deduction,
    )
    if err:
        return err
    match method:
        case "ex":
            return feature_manager.create_contour_flange_ex(thickness, bend_radius, direction)
        case "sync":
            return feature_manager.create_contour_flange_sync(
                face_index,
                edge_index,
                thickness,
                bend_radius,
                direction,
            )
        case "sync_with_bend":
            return feature_manager.create_contour_flange_sync_with_bend(
                face_index,
                edge_index,
                thickness,
                bend_radius,
                direction,
                bend_deduction,
            )
        case "v3":
            return feature_manager.create_contour_flange_v3(thickness, bend_radius, direction)
        case "sync_ex":
            return feature_manager.create_contour_flange_sync_ex(
                face_index,
                edge_index,
                thickness,
                bend_radius,
                direction,
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def create_sheet_metal_base(
    type: str = "flange",
    thickness: float = 0.0,
    width: float | None = None,
    bend_radius: float | None = None,
    relief_type: str = "Default",
) -> dict[str, Any]:
    """Create a sheet metal base feature.

    type: 'flange' | 'tab' | 'contour_advanced'
        | 'tab_multi_profile'

    Dimensions in meters.
    """
    err = validate_numerics(thickness=thickness)
    if err:
        return err
    match type:
        case "flange":
            return feature_manager.create_base_flange(width or 0.0, thickness, bend_radius)
        case "tab":
            return feature_manager.create_base_tab(thickness, width)
        case "contour_advanced":
            return feature_manager.create_base_contour_flange_advanced(
                thickness,
                bend_radius or 0.001,
                relief_type,
            )
        case "tab_multi_profile":
            return feature_manager.create_base_tab_multi_profile(thickness)
        case _:
            return {"error": f"Unknown type: {type}"}


def create_lofted_flange(
    method: str = "basic",
    thickness: float = 0.0,
    bend_radius: float = 0.0,
) -> dict[str, Any]:
    """Create a lofted flange (sheet metal).

    method: 'basic' | 'advanced' | 'ex'

    Dimensions in meters.
    """
    err = validate_numerics(thickness=thickness, bend_radius=bend_radius)
    if err:
        return err
    match method:
        case "basic":
            return feature_manager.create_lofted_flange(thickness)
        case "advanced":
            return feature_manager.create_lofted_flange_advanced(thickness, bend_radius)
        case "ex":
            return feature_manager.create_lofted_flange_ex(thickness)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_bend(
    method: str = "basic",
    bend_angle: float = 90.0,
    direction: str = "Normal",
    moving_side: str = "Right",
    bend_deduction: float = 0.0,
) -> dict[str, Any]:
    """Create a bend feature on sheet metal.

    method: 'basic' | 'with_calc'

    bend_angle in degrees.
    """
    err = validate_numerics(bend_angle=bend_angle, bend_deduction=bend_deduction)
    if err:
        return err
    match method:
        case "basic":
            return feature_manager.create_bend(bend_angle, direction, moving_side)
        case "with_calc":
            return feature_manager.create_bend_with_calc(
                bend_angle,
                direction,
                moving_side,
                bend_deduction,
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def create_slot(
    method: str = "basic",
    width: float = 0.0,
    depth: float = 0.0,
    direction: str = "Normal",
) -> dict[str, Any]:
    """Create a slot feature.

    method: 'basic' | 'ex' | 'sync' | 'multi_body'
        | 'sync_multi_body'

    Dimensions in meters.
    """
    err = validate_numerics(width=width, depth=depth)
    if err:
        return err
    match method:
        case "basic":
            return feature_manager.create_slot(width, direction)
        case "ex":
            return feature_manager.create_slot_ex(width, depth, direction)
        case "sync":
            return feature_manager.create_slot_sync(width, depth)
        case "multi_body":
            return feature_manager.create_slot_multi_body(width, depth, direction)
        case "sync_multi_body":
            return feature_manager.create_slot_sync_multi_body(width, depth, direction)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_thread(
    method: str = "basic",
    face_index: int = 0,
    thread_diameter: float = 0.0,
    thread_depth: float = 0.0,
) -> dict[str, Any]:
    """Create a thread on a cylindrical face.

    method: 'basic' (cosmetic) | 'physical' (modeled geometry)

    face_index is 0-based. Diameter/depth in meters; 0 = auto-detect.
    """
    err = validate_numerics(thread_diameter=thread_diameter, thread_depth=thread_depth)
    if err:
        return err
    diameter = thread_diameter if thread_diameter > 0 else None
    depth = thread_depth if thread_depth > 0 else None

    match method:
        case "basic":
            return feature_manager.create_thread(
                face_index, thread_diameter=diameter, thread_depth=depth
            )
        case "physical":
            return feature_manager.create_thread_ex(
                face_index, thread_diameter=diameter, thread_depth=depth
            )
        case _:
            return {"error": f"Unknown method: {method}. Use 'basic' or 'physical'."}


def create_drawn_cutout(
    method: str = "basic",
    depth: float = 0.0,
    direction: str = "Normal",
) -> dict[str, Any]:
    """Create a drawn cutout (sheet metal).

    method: 'basic' | 'ex'

    depth in meters.
    """
    err = validate_numerics(depth=depth)
    if err:
        return err
    match method:
        case "basic":
            return feature_manager.create_drawn_cutout(depth, direction)
        case "ex":
            return feature_manager.create_drawn_cutout_ex(depth, direction)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_dimple(
    method: str = "basic",
    depth: float = 0.0,
    direction: str = "Normal",
    punch_tool_diameter: float = 0.01,
) -> dict[str, Any]:
    """Create a dimple feature.

    method: 'basic' | 'ex'

    Dimensions in meters.
    """
    err = validate_numerics(depth=depth, punch_tool_diameter=punch_tool_diameter)
    if err:
        return err
    match method:
        case "basic":
            return feature_manager.create_dimple(depth, direction)
        case "ex":
            return feature_manager.create_dimple_ex(depth, direction, punch_tool_diameter)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_louver(
    method: str = "basic",
    depth: float = 0.0,
) -> dict[str, Any]:
    """Create a louver feature.

    method: 'basic' | 'sync'

    depth in meters.
    """
    err = validate_numerics(depth=depth)
    if err:
        return err
    match method:
        case "basic":
            return feature_manager.create_louver(depth)
        case "sync":
            return feature_manager.create_louver_sync(depth)
        case _:
            return {"error": f"Unknown method: {method}"}


def sheet_metal_misc(
    action: str = "hem",
    face_index: int = 0,
    edge_index: int = 0,
    hem_width: float = 0.005,
    bend_radius: float = 0.001,
    hem_type: str = "Closed",
    jog_offset: float = 0.005,
    jog_angle: float = 90.0,
    direction: str = "Normal",
    moving_side: str = "Right",
    closure_type: str = "Close",
    edge_indices: list[int] | None = None,
    flange_length: float = 0.0,
    side: str = "Right",
    thickness: float = 0.001,
) -> dict[str, Any]:
    """Miscellaneous sheet metal operations.

    action: 'hem' | 'jog' | 'close_corner' | 'multi_edge_flange'
        | 'convert'

    Dimensions in meters. jog_angle in degrees.
    """
    err = validate_numerics(
        hem_width=hem_width, bend_radius=bend_radius,
        jog_offset=jog_offset, jog_angle=jog_angle,
        flange_length=flange_length, thickness=thickness,
    )
    if err:
        return err
    match action:
        case "hem":
            return feature_manager.create_hem(
                face_index, edge_index, hem_width, bend_radius, hem_type
            )
        case "jog":
            return feature_manager.create_jog(jog_offset, jog_angle, direction, moving_side)
        case "close_corner":
            return feature_manager.create_close_corner(face_index, edge_index, closure_type)
        case "multi_edge_flange":
            return feature_manager.create_multi_edge_flange(
                face_index, edge_indices or [], flange_length, side
            )
        case "convert":
            return feature_manager.convert_part_to_sheet_metal(thickness)
        case _:
            return {"error": f"Unknown action: {action}"}


def create_stamped(
    type: str = "bead",
    depth: float = 0.0,
) -> dict[str, Any]:
    """Create a stamped feature (bead or gusset).

    type: 'bead' | 'gusset'

    depth in meters.
    """
    err = validate_numerics(depth=depth)
    if err:
        return err
    match type:
        case "bead":
            return feature_manager.create_bead(depth)
        case "gusset":
            return feature_manager.create_gusset(depth)
        case _:
            return {"error": f"Unknown type: {type}"}


def create_surface_mark(
    type: str = "emboss",
    face_indices: list[int] | None = None,
    clearance: float = 0.001,
    thickness: float = 0.0,
    thicken: bool = False,
    default_side: bool = True,
) -> dict[str, Any]:
    """Create a surface mark (emboss or etch).

    type: 'emboss' | 'etch'

    Dimensions in meters.
    """
    err = validate_numerics(clearance=clearance, thickness=thickness)
    if err:
        return err
    match type:
        case "emboss":
            return feature_manager.create_emboss(
                face_indices or [], clearance, thickness, thicken, default_side
            )
        case "etch":
            return feature_manager.create_etch()
        case _:
            return {"error": f"Unknown type: {type}"}


def create_reinforcement(
    type: str = "rib",
    thickness: float = 0.0,
    direction: str = "Normal",
) -> dict[str, Any]:
    """Create a reinforcement feature (rib or lip).

    type: 'rib' | 'lip'

    thickness in meters.
    """
    err = validate_numerics(thickness=thickness)
    if err:
        return err
    match type:
        case "rib":
            return feature_manager.create_rib(thickness, direction)
        case "lip":
            return feature_manager.create_lip(thickness)
        case _:
            return {"error": f"Unknown type: {type}"}


def create_web_network() -> dict[str, Any]:
    """Create a web network (sheet metal)."""
    return feature_manager.create_web_network()


def create_split() -> dict[str, Any]:
    """Split the body using the active profile."""
    return feature_manager.create_split()
