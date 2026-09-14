"""Revolve protrusion tools."""

from typing import Any, Literal

from solidedge_mcp.backends.validation import validate_numerics
from solidedge_mcp.managers import feature_manager


def create_revolve(
    method: Literal[
        "full",
        "finite",
        "sync",
        "finite_sync",
        "thin_wall",
        "by_keypoint",
        "full_360",
        "by_keypoint_sync",
    ] = "full",
    angle: float = 360.0,
    wall_thickness: float = 0.0,
    treatment_type: Literal["None", "Draft", "Crown", "CrownAndDraft"] = "None",
) -> dict[str, Any]:
    """Create a revolved protrusion around the sketch's revolve axis.

    angle in degrees (used by all but by_keypoint*). wall_thickness in meters
    (thin_wall). The revolve axis always comes from the sketch: set it with
    manage_sketch(action='set_axis') before closing. 'by_keypoint' and
    'by_keypoint_sync' are unsupported (they
    need a KeyPoint or tangent face object this server cannot select); use
    'full'. 'full_360' only accepts treatment_type='None' -
    RevolvedProtrusions.Add has no draft or crown slots.
    """
    err = validate_numerics(angle=angle, wall_thickness=wall_thickness)
    if err:
        return err
    match method:
        case "full":
            return feature_manager.create_revolve(angle=angle)
        case "finite":
            return feature_manager.create_revolve_finite(angle=angle)
        case "sync":
            return feature_manager.create_revolve_sync(angle=angle)
        case "finite_sync":
            return feature_manager.create_revolve_finite_sync(angle=angle)
        case "thin_wall":
            return feature_manager.create_revolve_thin_wall(
                angle=angle, wall_thickness=wall_thickness
            )
        case "by_keypoint":
            return feature_manager.create_revolve_by_keypoint()
        case "full_360":
            return feature_manager.create_revolve_full(angle=angle, treatment_type=treatment_type)
        case "by_keypoint_sync":
            return feature_manager.create_revolve_by_keypoint_sync()
        case _:
            return {"error": f"Unknown method: {method}"}
