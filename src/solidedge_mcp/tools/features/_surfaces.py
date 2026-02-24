"""Surface creation tools."""

from solidedge_mcp.backends.validation import validate_numerics
from solidedge_mcp.managers import feature_manager


def create_extruded_surface(
    method: str = "finite",
    distance: float = 0.0,
    direction: str = "Normal",
    end_caps: bool = True,
    from_plane_index: int = 0,
    to_plane_index: int = 0,
    keypoint_type: str = "End",
    treatment_type: str = "None",
    draft_angle: float = 0.0,
) -> dict:
    """Create an extruded surface.

    method: 'finite' | 'from_to' | 'by_keypoint'
        | 'by_curves' | 'full'

    distance in meters. draft_angle in degrees. Plane indices are 1-based.
    """
    err = validate_numerics(distance=distance, draft_angle=draft_angle)
    if err:
        return err
    match method:
        case "finite":
            return feature_manager.create_extruded_surface(distance, direction, end_caps)
        case "from_to":
            return feature_manager.create_extruded_surface_from_to(from_plane_index, to_plane_index)
        case "by_keypoint":
            return feature_manager.create_extruded_surface_by_keypoint(keypoint_type)
        case "by_curves":
            return feature_manager.create_extruded_surface_by_curves(distance, direction)
        case "full":
            return feature_manager.create_extruded_surface_full(
                distance,
                direction,
                treatment_type,
                draft_angle,
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def create_revolved_surface(
    method: str = "finite",
    angle: float = 360.0,
    want_end_caps: bool = False,
    keypoint_type: str = "End",
) -> dict:
    """Create a revolved surface.

    method: 'finite' | 'sync' | 'by_keypoint' | 'full'
        | 'full_sync'

    angle in degrees.
    """
    err = validate_numerics(angle=angle)
    if err:
        return err
    match method:
        case "finite":
            return feature_manager.create_revolved_surface(angle, want_end_caps)
        case "sync":
            return feature_manager.create_revolved_surface_sync(angle, want_end_caps)
        case "by_keypoint":
            return feature_manager.create_revolved_surface_by_keypoint(keypoint_type, want_end_caps)
        case "full":
            return feature_manager.create_revolved_surface_full(angle, want_end_caps)
        case "full_sync":
            return feature_manager.create_revolved_surface_full_sync(angle, want_end_caps)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_lofted_surface(
    method: str = "basic",
    want_end_caps: bool = False,
) -> dict:
    """Create a lofted surface.

    method: 'basic' | 'v2'
    """
    match method:
        case "basic":
            return feature_manager.create_lofted_surface(want_end_caps)
        case "v2":
            return feature_manager.create_lofted_surface_v2(want_end_caps)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_swept_surface(
    method: str = "basic",
    path_profile_index: int | None = None,
    want_end_caps: bool = False,
) -> dict:
    """Create a swept surface.

    method: 'basic' | 'ex'
    """
    match method:
        case "basic":
            return feature_manager.create_swept_surface(path_profile_index, want_end_caps)
        case "ex":
            return feature_manager.create_swept_surface_ex(path_profile_index, want_end_caps)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_bounded_surface(
    want_end_caps: bool = True,
    periodic: bool = False,
) -> dict:
    """Create a bounded (blue) surface from accumulated profiles."""
    return feature_manager.create_bounded_surface(want_end_caps, periodic)
