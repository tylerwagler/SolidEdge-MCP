"""Cutout tools (extruded, revolved, normal, lofted, swept, helix)."""

from typing import Any, Literal

from solidedge_mcp.backends.validation import validate_numerics
from solidedge_mcp.managers import feature_manager


def create_extruded_cutout(
    method: Literal[
        "finite",
        "through_all",
        "through_next",
        "from_to",
        "from_to_v2",
        "by_keypoint",
        "through_next_single",
        "multi_body",
        "from_to_multi_body",
        "through_all_multi_body",
    ] = "finite",
    distance: float = 0.0,
    direction: Literal["Normal", "Reverse"] = "Normal",
    from_plane_index: int = 0,
    to_plane_index: int = 0,
) -> dict[str, Any]:
    """Create an extruded cutout (removes material) from the active sketch.

    distance in meters (finite, multi_body). from/to_plane_index: required for
    from_to* methods; 1-based (1=Top/XY, 2=Right/YZ, 3=Front/XZ).
    'by_keypoint' (unsupported: needs a KeyPoint or tangent face object this
    server cannot select); use 'finite', 'from_to' or 'through_all'.
    """
    err = validate_numerics(distance=distance)
    if err:
        return err
    match method:
        case "finite":
            return feature_manager.create_extruded_cutout(distance, direction)
        case "through_all":
            return feature_manager.create_extruded_cutout_through_all(direction)
        case "through_next":
            return feature_manager.create_extruded_cutout_through_next(direction)
        case "from_to":
            return feature_manager.create_extruded_cutout_from_to(from_plane_index, to_plane_index)
        case "from_to_v2":
            return feature_manager.create_extruded_cutout_from_to_v2(
                from_plane_index, to_plane_index
            )
        case "by_keypoint":
            return feature_manager.create_extruded_cutout_by_keypoint(direction)
        case "through_next_single":
            return feature_manager.create_extruded_cutout_through_next_single(direction)
        case "multi_body":
            return feature_manager.create_extruded_cutout_multi_body(distance, direction)
        case "from_to_multi_body":
            return feature_manager.create_extruded_cutout_from_to_multi_body(
                from_plane_index, to_plane_index
            )
        case "through_all_multi_body":
            return feature_manager.create_extruded_cutout_through_all_multi_body(direction)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_revolved_cutout(
    method: Literal["finite", "sync", "by_keypoint", "multi_body", "full", "full_sync"] = "finite",
    angle: float = 360.0,
) -> dict[str, Any]:
    """Create a revolved cutout around the sketch's revolve axis.

    angle in degrees. 'by_keypoint' (unsupported: needs a KeyPoint or tangent
    face object this server cannot select); use 'finite' or 'full'.
    """
    err = validate_numerics(angle=angle)
    if err:
        return err
    match method:
        case "finite":
            return feature_manager.create_revolved_cutout(angle)
        case "sync":
            return feature_manager.create_revolved_cutout_sync(angle)
        case "by_keypoint":
            return feature_manager.create_revolved_cutout_by_keypoint()
        case "multi_body":
            return feature_manager.create_revolved_cutout_multi_body(angle)
        case "full":
            return feature_manager.create_revolved_cutout_full(angle)
        case "full_sync":
            return feature_manager.create_revolved_cutout_full_sync(angle)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_normal_cutout(
    method: Literal["finite", "through_all", "from_to", "through_next", "by_keypoint"] = "finite",
    distance: float = 0.0,
    direction: Literal["Normal", "Reverse"] = "Normal",
    from_plane_index: int = 0,
    to_plane_index: int = 0,
) -> dict[str, Any]:
    """Create a normal cutout (perpendicular to a face) from the active sketch.

    distance in meters (finite). from/to_plane_index: required for from_to;
    1-based (1=Top/XY, 2=Right/YZ, 3=Front/XZ).
    'by_keypoint' (unsupported: needs a KeyPoint or tangent face object this
    server cannot select); use 'finite' or 'through_all'.
    """
    err = validate_numerics(distance=distance)
    if err:
        return err
    match method:
        case "finite":
            return feature_manager.create_normal_cutout(distance, direction)
        case "through_all":
            return feature_manager.create_normal_cutout_through_all(direction)
        case "from_to":
            return feature_manager.create_normal_cutout_from_to(from_plane_index, to_plane_index)
        case "through_next":
            return feature_manager.create_normal_cutout_through_next(direction)
        case "by_keypoint":
            return feature_manager.create_normal_cutout_by_keypoint(direction)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_lofted_cutout(
    method: Literal["basic", "full"] = "basic",
    profile_indices: list[int] | None = None,
) -> dict[str, Any]:
    """Create a lofted cutout between multiple closed profiles.

    profile_indices: optional 0-based indices into accumulated profiles
    (default: all).
    """
    match method:
        case "basic":
            return feature_manager.create_lofted_cutout(profile_indices)
        case "full":
            return feature_manager.create_lofted_cutout_full(profile_indices)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_swept_cutout(
    method: Literal["basic", "multi_body"] = "basic",
    path_profile_index: int | None = None,
) -> dict[str, Any]:
    """Create a swept cutout along a path profile.

    path_profile_index: 0-based index of the path profile (multi_body only;
    default: auto).
    """
    match method:
        case "basic":
            return feature_manager.create_swept_cutout()
        case "multi_body":
            return feature_manager.create_swept_cutout_multi_body(path_profile_index)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_helix_cutout(
    method: Literal["finite", "sync", "from_to", "from_to_sync"] = "finite",
    pitch: float = 0.0,
    height: float = 0.0,
    revolutions: float | None = None,
    direction: Literal["Right", "Left"] = "Right",
    from_plane_index: int = 0,
    to_plane_index: int = 0,
) -> dict[str, Any]:
    """Create a helical cutout from the active sketch.

    pitch/height in meters. height, revolutions (default height/pitch) and
    direction (hand of helix) apply to finite/sync. from/to_plane_index:
    required for from_to*; 1-based (1=Top/XY, 2=Right/YZ, 3=Front/XZ).
    """
    err = validate_numerics(pitch=pitch, height=height)
    if err:
        return err
    match method:
        case "finite":
            return feature_manager.create_helix_cutout(pitch, height, revolutions, direction)
        case "sync":
            return feature_manager.create_helix_cutout_sync(pitch, height, revolutions, direction)
        case "from_to":
            return feature_manager.create_helix_cutout_from_to(
                from_plane_index, to_plane_index, pitch
            )
        case "from_to_sync":
            return feature_manager.create_helix_cutout_from_to_sync(
                from_plane_index, to_plane_index, pitch
            )
        case _:
            return {"error": f"Unknown method: {method}"}
