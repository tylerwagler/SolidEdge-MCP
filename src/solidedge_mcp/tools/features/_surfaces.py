"""Surface creation tools."""

from typing import Any, Literal

from solidedge_mcp.backends.validation import validate_numerics
from solidedge_mcp.managers import feature_manager


def create_extruded_surface(
    method: Literal["finite", "from_to", "by_keypoint", "by_curves", "full"] = "finite",
    distance: float = 0.0,
    direction: Literal["Normal", "Symmetric"] = "Normal",
    end_caps: bool = True,
    from_plane_index: int = 0,
    to_plane_index: int = 0,
    keypoint_type: Literal["Start", "End"] = "End",
    treatment_type: Literal["None", "Draft", "Crown", "CrownAndDraft"] = "None",
    draft_angle: float = 0.0,
) -> dict[str, Any]:
    """Create an extruded surface from the active sketch.

    distance in meters and direction: finite/by_curves/full. end_caps: finite.
    from/to_plane_index: required for from_to; 1-based (1=Top/XY, 2=Right/YZ,
    3=Front/XZ). keypoint_type: by_keypoint. treatment_type and draft_angle
    (degrees): full.
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
    method: Literal["finite", "sync", "by_keypoint", "full", "full_sync"] = "finite",
    angle: float = 360.0,
    want_end_caps: bool = False,
    keypoint_type: Literal["Start", "End"] = "End",
) -> dict[str, Any]:
    """Create a revolved surface around the sketch's revolve axis.

    angle in degrees (all but by_keypoint). keypoint_type: by_keypoint only.
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
    method: Literal["basic", "v2"] = "basic",
    want_end_caps: bool = False,
) -> dict[str, Any]:
    """Create a lofted surface through the accumulated sketch profiles."""
    match method:
        case "basic":
            return feature_manager.create_lofted_surface(want_end_caps)
        case "v2":
            return feature_manager.create_lofted_surface_v2(want_end_caps)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_swept_surface(
    method: Literal["basic", "ex"] = "basic",
    path_profile_index: int | None = None,
    want_end_caps: bool = False,
) -> dict[str, Any]:
    """Create a swept surface along a path profile.

    path_profile_index: 0-based index of the path profile (default: auto).
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
) -> dict[str, Any]:
    """Create a bounded (blue) surface from the accumulated sketch profiles."""
    return feature_manager.create_bounded_surface(want_end_caps, periodic)
