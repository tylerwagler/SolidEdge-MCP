"""Hole creation tools."""

from typing import Any, Literal

from solidedge_mcp.backends.validation import validate_numerics
from solidedge_mcp.managers import feature_manager


def create_hole(
    method: Literal[
        "finite",
        "through_all",
        "from_to",
        "through_next",
        "sync",
        "finite_ex",
        "from_to_ex",
        "through_next_ex",
        "through_all_ex",
        "sync_ex",
        "multi_body",
        "sync_multi_body",
    ] = "finite",
    x: float = 0.0,
    y: float = 0.0,
    diameter: float = 0.0,
    depth: float = 0.0,
    direction: Literal["Normal", "Reverse"] = "Normal",
    plane_index: int = 1,
    from_plane_index: int = 0,
    to_plane_index: int = 0,
) -> dict[str, Any]:
    """Create a hole at (x, y) on a reference plane.

    All lengths in meters. depth: finite*, sync*, multi_body. plane_index is
    1-based (1=Top/XY, 2=Right/YZ, 3=Front/XZ; ignored by finite/from_to*).
    from/to_plane_index: required for from_to/from_to_ex; 1-based.
    """
    err = validate_numerics(x=x, y=y, diameter=diameter, depth=depth)
    if err:
        return err
    match method:
        case "finite":
            return feature_manager.create_hole(
                x, y, diameter, depth, plane_index=plane_index, direction=direction
            )
        case "through_all":
            return feature_manager.create_hole_through_all(x, y, diameter, plane_index, direction)
        case "from_to":
            return feature_manager.create_hole_from_to(
                x,
                y,
                diameter,
                from_plane_index,
                to_plane_index,
            )
        case "through_next":
            return feature_manager.create_hole_through_next(x, y, diameter, direction, plane_index)
        case "sync":
            return feature_manager.create_hole_sync(x, y, diameter, depth, plane_index)
        case "finite_ex":
            return feature_manager.create_hole_finite_ex(
                x, y, diameter, depth, direction, plane_index
            )
        case "from_to_ex":
            return feature_manager.create_hole_from_to_ex(
                x,
                y,
                diameter,
                from_plane_index,
                to_plane_index,
            )
        case "through_next_ex":
            return feature_manager.create_hole_through_next_ex(
                x, y, diameter, direction, plane_index
            )
        case "through_all_ex":
            return feature_manager.create_hole_through_all_ex(x, y, diameter, plane_index)
        case "sync_ex":
            return feature_manager.create_hole_sync_ex(x, y, diameter, depth, plane_index)
        case "multi_body":
            return feature_manager.create_hole_multi_body(
                x, y, diameter, depth, plane_index, direction
            )
        case "sync_multi_body":
            return feature_manager.create_hole_sync_multi_body(x, y, diameter, depth, plane_index)
        case _:
            return {"error": f"Unknown method: {method}"}
