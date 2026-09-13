"""Extrude protrusion tools."""

from typing import Any, Literal

from solidedge_mcp.backends.validation import validate_numerics
from solidedge_mcp.managers import feature_manager


def create_extrude(
    method: Literal[
        "finite",
        "infinite",
        "through_next",
        "from_to",
        "thin_wall",
        "symmetric",
        "through_next_v2",
        "from_to_v2",
        "by_keypoint",
        "from_to_single",
        "through_next_single",
    ] = "finite",
    distance: float = 0.0,
    direction: Literal["Normal", "Reverse", "Symmetric"] = "Normal",
    operation: Literal["Add", "Cut", "Intersect"] = "Add",
    wall_thickness: float = 0.0,
    from_plane_index: int = 0,
    to_plane_index: int = 0,
) -> dict[str, Any]:
    """Create an extruded protrusion from the active sketch profile.

    Lengths in meters. distance: finite/thin_wall/symmetric/by_keypoint.
    wall_thickness: thin_wall. direction 'Symmetric' only for finite/thin_wall
    (others treat it as Normal). from/to_plane_index: required for from_to*
    methods; 1-based (1=Top/XY, 2=Right/YZ, 3=Front/XZ).
    operation applies to 'finite' only: 'Cut' removes material and needs an
    existing base feature; 'Intersect' is unsupported.
    'by_keypoint' (unsupported: needs a KeyPoint or tangent face object this
    server cannot select); use 'finite' or 'from_to'.
    """
    err = validate_numerics(distance=distance, wall_thickness=wall_thickness)
    if err:
        return err
    match method:
        case "finite":
            return feature_manager.create_extrude(
                distance, operation=operation, direction=direction
            )
        case "infinite":
            return feature_manager.create_extrude_infinite(direction)
        case "through_next":
            return feature_manager.create_extrude_through_next(direction)
        case "from_to":
            return feature_manager.create_extrude_from_to(from_plane_index, to_plane_index)
        case "thin_wall":
            return feature_manager.create_extrude_thin_wall(distance, wall_thickness, direction)
        case "symmetric":
            return feature_manager.create_extrude_symmetric(distance)
        case "through_next_v2":
            return feature_manager.create_extrude_through_next_v2(direction)
        case "from_to_v2":
            return feature_manager.create_extrude_from_to_v2(from_plane_index, to_plane_index)
        case "by_keypoint":
            return feature_manager.create_extrude_by_keypoint(direction)
        case "from_to_single":
            return feature_manager.create_extrude_from_to_single(from_plane_index, to_plane_index)
        case "through_next_single":
            return feature_manager.create_extrude_through_next_single(direction)
        case _:
            return {"error": f"Unknown method: {method}"}
