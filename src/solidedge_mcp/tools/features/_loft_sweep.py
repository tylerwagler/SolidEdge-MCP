"""Helix, loft, and sweep tools."""

from solidedge_mcp.managers import feature_manager


def create_helix(
    method: str = "finite",
    pitch: float = 0.0,
    height: float = 0.0,
    revolutions: float | None = None,
    direction: str = "Right",
    wall_thickness: float = 0.0,
    from_plane_index: int = 0,
    to_plane_index: int = 0,
) -> dict:
    """Create a helical feature (springs, threads).

    method: 'finite' | 'sync' | 'thin_wall'
        | 'sync_thin_wall' | 'from_to' | 'from_to_thin_wall'
        | 'from_to_sync' | 'from_to_sync_thin_wall'

    Parameters (used per method):
        pitch: Helix pitch distance.
        height: Helix height (finite, sync, thin_wall,
            sync_thin_wall).
        revolutions: Number of revolutions.
        direction: 'Right' or 'Left' (finite).
        wall_thickness: Wall thickness (thin_wall,
            sync_thin_wall, from_to_thin_wall,
            from_to_sync_thin_wall).
        from_plane_index: 1-based ref plane (from_to variants).
        to_plane_index: 1-based ref plane (from_to variants).
    """
    match method:
        case "finite":
            return feature_manager.create_helix(pitch, height, revolutions, direction)
        case "sync":
            return feature_manager.create_helix_sync(pitch, height, revolutions)
        case "thin_wall":
            return feature_manager.create_helix_thin_wall(
                pitch, height, wall_thickness, revolutions
            )
        case "sync_thin_wall":
            return feature_manager.create_helix_sync_thin_wall(
                pitch, height, wall_thickness, revolutions
            )
        case "from_to":
            return feature_manager.create_helix_from_to(from_plane_index, to_plane_index, pitch)
        case "from_to_thin_wall":
            return feature_manager.create_helix_from_to_thin_wall(
                from_plane_index,
                to_plane_index,
                pitch,
                wall_thickness,
            )
        case "from_to_sync":
            return feature_manager.create_helix_from_to_sync(
                from_plane_index, to_plane_index, pitch
            )
        case "from_to_sync_thin_wall":
            return feature_manager.create_helix_from_to_sync_thin_wall(
                from_plane_index,
                to_plane_index,
                pitch,
                wall_thickness,
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def create_loft(
    method: str = "solid",
    profile_indices: list | None = None,
    wall_thickness: float = 0.0,
    guide_profile_indices: list[int] | None = None,
) -> dict:
    """Create a loft feature between multiple profiles.

    method: 'solid' | 'thin_wall' | 'with_guides'

    Parameters (used per method):
        profile_indices: Indices of accumulated profiles.
        wall_thickness: Wall thickness (thin_wall).
        guide_profile_indices: Indices of guide curve profiles
            (with_guides).
    """
    match method:
        case "solid":
            return feature_manager.create_loft(profile_indices)
        case "thin_wall":
            return feature_manager.create_loft_thin_wall(wall_thickness, profile_indices)
        case "with_guides":
            return feature_manager.create_loft_with_guides(guide_profile_indices, profile_indices)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_sweep(
    method: str = "solid",
    path_profile_index: int | None = None,
    wall_thickness: float = 0.0,
) -> dict:
    """Create a sweep feature along a path.

    method: 'solid' | 'thin_wall'

    Parameters (used per method):
        path_profile_index: Index of path profile.
        wall_thickness: Wall thickness (thin_wall).
    """
    match method:
        case "solid":
            return feature_manager.create_sweep(path_profile_index)
        case "thin_wall":
            return feature_manager.create_sweep_thin_wall(wall_thickness, path_profile_index)
        case _:
            return {"error": f"Unknown method: {method}"}
