"""Revolve protrusion tools."""

from typing import Any

from solidedge_mcp.backends.validation import validate_numerics
from solidedge_mcp.managers import feature_manager


def create_revolve(
    method: str = "full",
    angle: float = 360.0,
    axis_type: str = "CenterLine",
    wall_thickness: float = 0.0,
    treatment_type: str = "None",
) -> dict[str, Any]:
    """Create a revolved protrusion around the set axis.

    method: 'full' | 'finite' | 'sync' | 'finite_sync'
        | 'thin_wall' | 'by_keypoint' | 'full_360'
        | 'by_keypoint_sync'

    angle in degrees. wall_thickness in meters.
    """
    err = validate_numerics(angle=angle, wall_thickness=wall_thickness)
    if err:
        return err
    match method:
        case "full":
            return feature_manager.create_revolve(angle)
        case "finite":
            return feature_manager.create_revolve_finite(angle, axis_type)
        case "sync":
            return feature_manager.create_revolve_sync(angle)
        case "finite_sync":
            return feature_manager.create_revolve_finite_sync(angle)
        case "thin_wall":
            return feature_manager.create_revolve_thin_wall(angle, wall_thickness)
        case "by_keypoint":
            return feature_manager.create_revolve_by_keypoint()
        case "full_360":
            return feature_manager.create_revolve_full(angle, treatment_type)
        case "by_keypoint_sync":
            return feature_manager.create_revolve_by_keypoint_sync()
        case _:
            return {"error": f"Unknown method: {method}"}
