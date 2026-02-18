"""Feature modeling tools for Solid Edge MCP.

Consolidated composite tools that use a `method` (or `shape`/`type`/`action`)
discriminator parameter to dispatch to the correct backend method.
"""

from solidedge_mcp.managers import feature_manager

# =====================================================================
# Group 1: create_extrude (11 -> 1)
# =====================================================================


def create_extrude(
    method: str = "finite",
    distance: float = 0.0,
    direction: str = "Normal",
    wall_thickness: float = 0.0,
    from_plane_index: int = 0,
    to_plane_index: int = 0,
) -> dict:
    """Create an extruded protrusion from the active sketch profile.

    method: 'finite' | 'infinite' | 'through_next' | 'from_to'
        | 'thin_wall' | 'symmetric' | 'through_next_v2'
        | 'from_to_v2' | 'by_keypoint' | 'from_to_single'
        | 'through_next_single'

    Parameters (used per method):
        distance: Extrusion distance (finite, thin_wall, symmetric).
        direction: 'Normal'|'Reverse'|'Both' (finite, infinite,
            through_next, thin_wall, through_next_v2, by_keypoint,
            through_next_single).
        wall_thickness: Wall thickness (thin_wall).
        from_plane_index: 1-based ref plane (from_to, from_to_v2,
            from_to_single).
        to_plane_index: 1-based ref plane (from_to, from_to_v2,
            from_to_single).
    """
    match method:
        case "finite":
            return feature_manager.create_extrude(distance, direction)
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


# =====================================================================
# Group 2: create_revolve (8 -> 1)
# =====================================================================


def create_revolve(
    method: str = "full",
    angle: float = 360.0,
    axis_type: str = "CenterLine",
    wall_thickness: float = 0.0,
    treatment_type: str = "None",
) -> dict:
    """Create a revolved protrusion around the set axis.

    method: 'full' | 'finite' | 'sync' | 'finite_sync'
        | 'thin_wall' | 'by_keypoint' | 'full_360'
        | 'by_keypoint_sync'

    Parameters (used per method):
        angle: Revolve angle in degrees (full, finite, sync,
            finite_sync, thin_wall, full_360).
        axis_type: 'CenterLine' or other axis type (finite).
        wall_thickness: Wall thickness (thin_wall).
        treatment_type: Treatment type string (full_360).
    """
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


# =====================================================================
# Group 3: create_extruded_cutout (10 -> 1)
# =====================================================================


def create_extruded_cutout(
    method: str = "finite",
    distance: float = 0.0,
    direction: str = "Normal",
    from_plane_index: int = 0,
    to_plane_index: int = 0,
) -> dict:
    """Create an extruded cutout (removes material).

    method: 'finite' | 'through_all' | 'through_next' | 'from_to'
        | 'from_to_v2' | 'by_keypoint' | 'through_next_single'
        | 'multi_body' | 'from_to_multi_body'
        | 'through_all_multi_body'

    Parameters (used per method):
        distance: Cut depth (finite, multi_body).
        direction: 'Normal'|'Reverse'|'Both' (finite, through_all,
            through_next, by_keypoint, through_next_single,
            multi_body, through_all_multi_body).
        from_plane_index: 1-based ref plane (from_to,
            from_to_v2, from_to_multi_body).
        to_plane_index: 1-based ref plane (from_to,
            from_to_v2, from_to_multi_body).
    """
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


# =====================================================================
# Group 4: create_revolved_cutout (6 -> 1)
# =====================================================================


def create_revolved_cutout(
    method: str = "finite",
    angle: float = 360.0,
) -> dict:
    """Create a revolved cutout around the set axis.

    method: 'finite' | 'sync' | 'by_keypoint' | 'multi_body'
        | 'full' | 'full_sync'

    Parameters (used per method):
        angle: Revolve angle in degrees (finite, sync, multi_body,
            full, full_sync).
    """
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


# =====================================================================
# Group 5: create_normal_cutout (5 -> 1)
# =====================================================================


def create_normal_cutout(
    method: str = "finite",
    distance: float = 0.0,
    direction: str = "Normal",
    from_plane_index: int = 0,
    to_plane_index: int = 0,
) -> dict:
    """Create a normal cutout perpendicular to a face.

    method: 'finite' | 'through_all' | 'from_to'
        | 'through_next' | 'by_keypoint'

    Parameters (used per method):
        distance: Cut depth (finite).
        direction: 'Normal'|'Reverse'|'Both' (finite, through_all,
            through_next, by_keypoint).
        from_plane_index: 1-based ref plane (from_to).
        to_plane_index: 1-based ref plane (from_to).
    """
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


# =====================================================================
# Group 6: create_lofted_cutout (2 -> 1)
# =====================================================================


def create_lofted_cutout(
    method: str = "basic",
    profile_indices: list[int] | None = None,
) -> dict:
    """Create a lofted cutout between multiple profiles.

    method: 'basic' | 'full'

    Parameters:
        profile_indices: Indices of accumulated profiles.
    """
    match method:
        case "basic":
            return feature_manager.create_lofted_cutout(profile_indices)
        case "full":
            return feature_manager.create_lofted_cutout_full(profile_indices)
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 7: create_swept_cutout (2 -> 1)
# =====================================================================


def create_swept_cutout(
    method: str = "basic",
    path_profile_index: int | None = None,
) -> dict:
    """Create a swept cutout along a path.

    method: 'basic' | 'multi_body'

    Parameters:
        path_profile_index: Index of path profile (multi_body).
    """
    match method:
        case "basic":
            return feature_manager.create_swept_cutout()
        case "multi_body":
            return feature_manager.create_swept_cutout_multi_body(path_profile_index)
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 8: create_helix (8 -> 1)
# =====================================================================


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


# =====================================================================
# Group 9: create_helix_cutout (4 -> 1)
# =====================================================================


def create_helix_cutout(
    method: str = "finite",
    pitch: float = 0.0,
    height: float = 0.0,
    revolutions: float | None = None,
    direction: str = "Right",
    from_plane_index: int = 0,
    to_plane_index: int = 0,
) -> dict:
    """Create a helical cutout.

    method: 'finite' | 'sync' | 'from_to' | 'from_to_sync'

    Parameters (used per method):
        pitch: Helix pitch distance.
        height: Helix height (finite, sync).
        revolutions: Number of revolutions (finite, sync).
        direction: 'Right' or 'Left' (finite, sync).
        from_plane_index: 1-based ref plane (from_to,
            from_to_sync).
        to_plane_index: 1-based ref plane (from_to,
            from_to_sync).
    """
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


# =====================================================================
# Group 10: create_loft (3 -> 1)
# =====================================================================


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


# =====================================================================
# Group 11: create_sweep (2 -> 1)
# =====================================================================


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


# =====================================================================
# Group 12: create_extruded_surface (5 -> 1)
# =====================================================================


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

    Parameters (used per method):
        distance: Extrusion distance (finite, by_curves, full).
        direction: Direction (finite, by_curves, full).
        end_caps: Whether to add end caps (finite).
        from_plane_index: 1-based ref plane (from_to).
        to_plane_index: 1-based ref plane (from_to).
        keypoint_type: 'Start' or 'End' (by_keypoint).
        treatment_type: Treatment type string (full).
        draft_angle: Draft angle in degrees (full).
    """
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


# =====================================================================
# Group 13: create_revolved_surface (5 -> 1)
# =====================================================================


def create_revolved_surface(
    method: str = "finite",
    angle: float = 360.0,
    want_end_caps: bool = False,
    keypoint_type: str = "End",
) -> dict:
    """Create a revolved surface.

    method: 'finite' | 'sync' | 'by_keypoint' | 'full'
        | 'full_sync'

    Parameters (used per method):
        angle: Revolve angle in degrees.
        want_end_caps: Whether to add end caps.
        keypoint_type: 'Start' or 'End' (by_keypoint).
    """
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


# =====================================================================
# Group 14: create_lofted_surface (2 -> 1)
# =====================================================================


def create_lofted_surface(
    method: str = "basic",
    want_end_caps: bool = False,
) -> dict:
    """Create a lofted surface.

    method: 'basic' | 'v2'

    Parameters:
        want_end_caps: Whether to add end caps.
    """
    match method:
        case "basic":
            return feature_manager.create_lofted_surface(want_end_caps)
        case "v2":
            return feature_manager.create_lofted_surface_v2(want_end_caps)
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 15: create_swept_surface (2 -> 1)
# =====================================================================


def create_swept_surface(
    method: str = "basic",
    path_profile_index: int | None = None,
    want_end_caps: bool = False,
) -> dict:
    """Create a swept surface.

    method: 'basic' | 'ex'

    Parameters:
        path_profile_index: Index of path profile.
        want_end_caps: Whether to add end caps.
    """
    match method:
        case "basic":
            return feature_manager.create_swept_surface(path_profile_index, want_end_caps)
        case "ex":
            return feature_manager.create_swept_surface_ex(path_profile_index, want_end_caps)
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 16: thicken (2 -> 1)
# =====================================================================


def thicken(
    method: str = "basic",
    thickness: float = 0.0,
    direction: str = "Both",
) -> dict:
    """Thicken a surface to create a solid.

    method: 'basic' | 'sync'

    Parameters:
        thickness: Thicken distance.
        direction: 'Both'|'Normal'|'Reverse'.
    """
    match method:
        case "basic":
            return feature_manager.thicken_surface(thickness, direction)
        case "sync":
            return feature_manager.create_thicken_sync(thickness, direction)
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 17: create_primitive (5 -> 1)
# =====================================================================


def create_primitive(
    shape: str = "box_two_points",
    x1: float = 0.0,
    y1: float = 0.0,
    z1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    z2: float = 0.0,
    x3: float = 0.0,
    y3: float = 0.0,
    z3: float = 0.0,
    length: float = 0.0,
    width: float = 0.0,
    height: float = 0.0,
    radius: float = 0.0,
    depth: float = 0.0,
    plane_index: int = 1,
) -> dict:
    """Create a primitive solid shape.

    shape: 'box_two_points' | 'box_center' | 'box_three_points'
        | 'cylinder' | 'sphere'

    Parameters (used per shape):
        x1,y1,z1,x2,y2,z2: Corner coords (box_two_points),
            or center+corner (box_center uses x1,y1,z1 as
            center), or first two points (box_three_points).
        x3,y3,z3: Third point (box_three_points).
        length,width,height: Dimensions (box_center).
        radius: Radius (cylinder, sphere).
        depth: Cylinder height (cylinder).
        plane_index: 1=Top/XY, 2=Right/YZ, 3=Front/XZ.
    """
    match shape:
        case "box_two_points":
            return feature_manager.create_box_by_two_points(x1, y1, z1, x2, y2, z2, plane_index)
        case "box_center":
            return feature_manager.create_box_by_center(
                x1,
                y1,
                z1,
                length,
                width,
                height,
                plane_index,
            )
        case "box_three_points":
            return feature_manager.create_box_by_three_points(
                x1,
                y1,
                z1,
                x2,
                y2,
                z2,
                x3,
                y3,
                z3,
                plane_index,
            )
        case "cylinder":
            return feature_manager.create_cylinder(x1, y1, z1, radius, depth, plane_index)
        case "sphere":
            return feature_manager.create_sphere(x1, y1, z1, radius, plane_index)
        case _:
            return {"error": f"Unknown shape: {shape}"}


# =====================================================================
# Group 18: create_primitive_cutout (3 -> 1)
# =====================================================================


def create_primitive_cutout(
    shape: str = "box",
    x1: float = 0.0,
    y1: float = 0.0,
    z1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    z2: float = 0.0,
    radius: float = 0.0,
    height: float = 0.0,
    plane_index: int = 1,
) -> dict:
    """Create a primitive cutout shape.

    shape: 'box' | 'cylinder' | 'sphere'

    Parameters (used per shape):
        x1,y1,z1,x2,y2,z2: Corner coords (box).
        x1,y1,z1: Center coords (cylinder, sphere).
        radius: Radius (cylinder, sphere).
        height: Cylinder height (cylinder).
        plane_index: 1=Top/XY, 2=Right/YZ, 3=Front/XZ.
    """
    match shape:
        case "box":
            return feature_manager.create_box_cutout_by_two_points(
                x1, y1, z1, x2, y2, z2, plane_index
            )
        case "cylinder":
            return feature_manager.create_cylinder_cutout(x1, y1, z1, radius, height, plane_index)
        case "sphere":
            return feature_manager.create_sphere_cutout(x1, y1, z1, radius, plane_index)
        case _:
            return {"error": f"Unknown shape: {shape}"}


# =====================================================================
# Group 19: create_hole (12 -> 1)
# =====================================================================


def create_hole(
    method: str = "finite",
    x: float = 0.0,
    y: float = 0.0,
    diameter: float = 0.0,
    depth: float = 0.0,
    direction: str = "Normal",
    plane_index: int = 1,
    from_plane_index: int = 0,
    to_plane_index: int = 0,
) -> dict:
    """Create a hole at coordinates (meters).

    method: 'finite' | 'through_all' | 'from_to'
        | 'through_next' | 'sync' | 'finite_ex'
        | 'from_to_ex' | 'through_next_ex'
        | 'through_all_ex' | 'sync_ex' | 'multi_body'
        | 'sync_multi_body'

    Parameters (used per method):
        x, y: Hole center coordinates.
        diameter: Hole diameter.
        depth: Hole depth (finite, sync, finite_ex, sync_ex,
            multi_body, sync_multi_body).
        direction: 'Normal'|'Reverse' (finite, through_next,
            finite_ex, through_next_ex, multi_body).
        plane_index: 1-based ref plane index.
        from_plane_index: 1-based ref plane (from_to,
            from_to_ex).
        to_plane_index: 1-based ref plane (from_to,
            from_to_ex).
    """
    match method:
        case "finite":
            return feature_manager.create_hole(x, y, diameter, depth, direction)
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


# =====================================================================
# Group 20: create_round (5 -> 1)
# =====================================================================


def create_round(
    method: str = "all_edges",
    radius: float = 0.0,
    face_index: int | None = None,
    radii: list | None = None,
    face_index1: int = 0,
    face_index2: int = 0,
) -> dict:
    """Round (fillet) edges of the active body.

    method: 'all_edges' | 'on_face' | 'variable' | 'blend'
        | 'surface_blend'

    Parameters (used per method):
        radius: Fillet radius (all_edges, on_face, blend,
            surface_blend).
        face_index: Face index (on_face, variable).
        radii: List of radii (variable).
        face_index1: First face (blend, surface_blend).
        face_index2: Second face (blend, surface_blend).
    """
    match method:
        case "all_edges":
            return feature_manager.create_round(radius)
        case "on_face":
            return feature_manager.create_round_on_face(radius, face_index)
        case "variable":
            return feature_manager.create_variable_round(radii, face_index)
        case "blend":
            return feature_manager.create_round_blend(face_index1, face_index2, radius)
        case "surface_blend":
            return feature_manager.create_round_surface_blend(face_index1, face_index2, radius)
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 21: create_chamfer (5 -> 1)
# =====================================================================


def create_chamfer(
    method: str = "equal",
    distance: float = 0.0,
    face_index: int = 0,
    distance1: float = 0.0,
    distance2: float = 0.0,
    angle: float = 0.0,
) -> dict:
    """Chamfer edges of the active body.

    method: 'equal' | 'on_face' | 'unequal'
        | 'unequal_on_face' | 'angle'

    Parameters (used per method):
        distance: Chamfer distance (equal, on_face, angle).
        face_index: Face index (on_face, unequal,
            unequal_on_face, angle).
        distance1: First distance (unequal, unequal_on_face).
        distance2: Second distance (unequal, unequal_on_face).
        angle: Chamfer angle in degrees (angle).
    """
    match method:
        case "equal":
            return feature_manager.create_chamfer(distance)
        case "on_face":
            return feature_manager.create_chamfer_on_face(distance, face_index)
        case "unequal":
            return feature_manager.create_chamfer_unequal(distance1, distance2, face_index)
        case "unequal_on_face":
            return feature_manager.create_chamfer_unequal_on_face(distance1, distance2, face_index)
        case "angle":
            return feature_manager.create_chamfer_angle(distance, angle, face_index)
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 22: create_blend (3 -> 1)
# =====================================================================


def create_blend(
    method: str = "basic",
    radius: float = 0.0,
    face_index: int | None = None,
    radius1: float = 0.0,
    radius2: float = 0.0,
    face_index1: int = 0,
    face_index2: int = 0,
) -> dict:
    """Create a blend (face-to-face fillet).

    method: 'basic' | 'variable' | 'surface'

    Parameters (used per method):
        radius: Blend radius (basic).
        face_index: Face index (basic, variable).
        radius1: Start radius (variable).
        radius2: End radius (variable).
        face_index1: First face (surface).
        face_index2: Second face (surface).
    """
    match method:
        case "basic":
            return feature_manager.create_blend(radius, face_index)
        case "variable":
            return feature_manager.create_blend_variable(radius1, radius2, face_index)
        case "surface":
            return feature_manager.create_blend_surface(face_index1, face_index2)
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 23: delete_topology (4 -> 1)
# =====================================================================


def delete_topology(
    type: str = "hole",
    max_diameter: float = 1.0,
    hole_type: str = "All",
    face_index: int = 0,
    face_indices: list[int] | None = None,
) -> dict:
    """Delete topology features (holes, blends, faces).

    type: 'hole' | 'hole_by_face' | 'blend' | 'faces'

    Parameters (used per type):
        max_diameter: Maximum hole diameter (hole).
        hole_type: Hole type filter (hole).
        face_index: Face index (hole_by_face, blend).
        face_indices: List of face indices (faces).
    """
    match type:
        case "hole":
            return feature_manager.create_delete_hole(max_diameter, hole_type)
        case "hole_by_face":
            return feature_manager.delete_hole_by_face(face_index)
        case "blend":
            return feature_manager.create_delete_blend(face_index)
        case "faces":
            return feature_manager.delete_faces(face_indices)
        case _:
            return {"error": f"Unknown type: {type}"}


# =====================================================================
# Group 24: create_ref_plane (17 -> 1)
# =====================================================================


def create_ref_plane(
    method: str = "offset",
    parent_plane_index: int = 1,
    distance: float = 0.0,
    normal_side: str = "Normal",
    angle: float = 0.0,
    x1: float = 0.0,
    y1: float = 0.0,
    z1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    z2: float = 0.0,
    x3: float = 0.0,
    y3: float = 0.0,
    z3: float = 0.0,
    plane1_index: int = 0,
    plane2_index: int = 0,
    curve_end: str = "End",
    pivot_plane_index: int = 2,
    ratio: float = 0.0,
    distance_along: float = 0.0,
    face_index: int = 0,
    keypoint_type: str = "End",
    curve_edge_index: int = 0,
    orientation_plane_index: int = 0,
    normal_side_int: int = 2,
) -> dict:
    """Create a reference plane.

    method: 'offset' | 'angle' | 'three_points' | 'midplane'
        | 'normal_to_curve' | 'normal_at_distance'
        | 'normal_at_arc_ratio' | 'normal_at_distance_along'
        | 'parallel_by_tangent' | 'normal_at_keypoint'
        | 'tangent_cylinder_angle'
        | 'tangent_cylinder_keypoint'
        | 'tangent_surface_keypoint'
        | 'normal_at_distance_v2'
        | 'normal_at_arc_ratio_v2'
        | 'normal_at_distance_along_v2'
        | 'tangent_parallel'

    Parameters (used per method):
        parent_plane_index: 1-based parent plane (offset, angle,
            parallel_by_tangent, tangent_cylinder_angle,
            tangent_cylinder_keypoint,
            tangent_surface_keypoint, tangent_parallel).
        distance: Offset distance (offset, normal_at_distance).
        normal_side: 'Normal' or 'Reverse' (offset, angle,
            parallel_by_tangent).
        angle: Angle in degrees (angle,
            tangent_cylinder_angle).
        x1..z3: 9 coordinates (three_points).
        plane1_index, plane2_index: Planes (midplane).
        curve_end: 'Start' or 'End' (normal_to_curve,
            normal_at_distance, normal_at_arc_ratio,
            normal_at_distance_along).
        pivot_plane_index: Pivot plane (normal_to_curve,
            normal_at_distance, normal_at_arc_ratio,
            normal_at_distance_along).
        ratio: Arc-length ratio 0-1 (normal_at_arc_ratio).
        distance_along: Distance along curve
            (normal_at_distance_along).
        face_index: Face index (parallel_by_tangent,
            tangent_cylinder_angle,
            tangent_cylinder_keypoint,
            tangent_surface_keypoint, tangent_parallel).
        keypoint_type: 'Start' or 'End' (normal_at_keypoint,
            tangent_cylinder_keypoint,
            tangent_surface_keypoint).
        curve_edge_index: Edge index (v2 variants).
        orientation_plane_index: Orientation plane (v2 variants).
        normal_side_int: Integer normal side (v2 variants,
            tangent_parallel).
    """
    match method:
        case "offset":
            return feature_manager.create_ref_plane_by_offset(
                parent_plane_index, distance, normal_side
            )
        case "angle":
            return feature_manager.create_ref_plane_by_angle(parent_plane_index, angle, normal_side)
        case "three_points":
            return feature_manager.create_ref_plane_by_3_points(x1, y1, z1, x2, y2, z2, x3, y3, z3)
        case "midplane":
            return feature_manager.create_ref_plane_midplane(plane1_index, plane2_index)
        case "normal_to_curve":
            return feature_manager.create_ref_plane_normal_to_curve(curve_end, pivot_plane_index)
        case "normal_at_distance":
            return feature_manager.create_ref_plane_normal_at_distance(
                distance, curve_end, pivot_plane_index
            )
        case "normal_at_arc_ratio":
            return feature_manager.create_ref_plane_normal_at_arc_ratio(
                ratio, curve_end, pivot_plane_index
            )
        case "normal_at_distance_along":
            return feature_manager.create_ref_plane_normal_at_distance_along(
                distance_along,
                curve_end,
                pivot_plane_index,
            )
        case "parallel_by_tangent":
            return feature_manager.create_ref_plane_parallel_by_tangent(
                parent_plane_index,
                face_index,
                normal_side,
            )
        case "normal_at_keypoint":
            return feature_manager.create_ref_plane_normal_at_keypoint(
                keypoint_type, pivot_plane_index
            )
        case "tangent_cylinder_angle":
            return feature_manager.create_ref_plane_tangent_cylinder_angle(
                face_index, angle, parent_plane_index
            )
        case "tangent_cylinder_keypoint":
            return feature_manager.create_ref_plane_tangent_cylinder_keypoint(
                face_index,
                keypoint_type,
                parent_plane_index,
            )
        case "tangent_surface_keypoint":
            return feature_manager.create_ref_plane_tangent_surface_keypoint(
                face_index,
                keypoint_type,
                parent_plane_index,
            )
        case "normal_at_distance_v2":
            return feature_manager.create_ref_plane_normal_at_distance_v2(
                curve_edge_index,
                orientation_plane_index,
                distance,
                normal_side_int,
            )
        case "normal_at_arc_ratio_v2":
            return feature_manager.create_ref_plane_normal_at_arc_ratio_v2(
                curve_edge_index,
                orientation_plane_index,
                ratio,
                normal_side_int,
            )
        case "normal_at_distance_along_v2":
            return feature_manager.create_ref_plane_normal_at_distance_along_v2(
                curve_edge_index,
                orientation_plane_index,
                distance_along,
                normal_side_int,
            )
        case "tangent_parallel":
            return feature_manager.create_ref_plane_tangent_parallel(
                parent_plane_index,
                face_index,
                normal_side_int,
            )
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 25: create_flange (8 -> 1)
# =====================================================================


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
) -> dict:
    """Create a flange on an edge (sheet metal).

    method: 'basic' | 'by_match_face' | 'sync' | 'by_face'
        | 'with_bend_calc' | 'sync_with_bend_calc'
        | 'match_face_with_bend' | 'by_face_with_bend'

    Parameters (used per method):
        face_index: Face index.
        edge_index: Edge index.
        flange_length: Length of flange.
        side: 'Right' or 'Left' (basic, by_match_face,
            with_bend_calc, match_face_with_bend, by_face,
            by_face_with_bend).
        inside_radius: Bend inside radius (basic,
            by_match_face, match_face_with_bend).
        bend_angle: Bend angle (basic).
        bend_deduction: Bend deduction value (with_bend_calc,
            sync_with_bend_calc, match_face_with_bend,
            by_face_with_bend).
        ref_face_index: Reference face (by_face,
            by_face_with_bend).
        bend_radius: Bend radius (by_face, by_face_with_bend).
    """
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
                bend_deduction,
            )
        case "by_face_with_bend":
            return feature_manager.create_flange_by_face_with_bend(
                face_index,
                edge_index,
                ref_face_index,
                flange_length,
                side,
                bend_radius,
                bend_deduction,
            )
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 26: create_contour_flange (5 -> 1)
# =====================================================================


def create_contour_flange(
    method: str = "ex",
    thickness: float = 0.0,
    bend_radius: float = 0.001,
    direction: str = "Normal",
    face_index: int = 0,
    edge_index: int = 0,
    bend_deduction: float = 0.0,
) -> dict:
    """Create a contour flange (sheet metal).

    method: 'ex' | 'sync' | 'sync_with_bend' | 'v3'
        | 'sync_ex'

    Parameters (used per method):
        thickness: Material thickness.
        bend_radius: Bend radius.
        direction: Direction.
        face_index: Face index (sync, sync_with_bend,
            sync_ex).
        edge_index: Edge index (sync, sync_with_bend,
            sync_ex).
        bend_deduction: Bend deduction (sync_with_bend,
            sync_ex).
    """
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
                bend_deduction,
            )
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 27: create_sheet_metal_base (4 -> 1)
# =====================================================================


def create_sheet_metal_base(
    type: str = "flange",
    thickness: float = 0.0,
    width: float | None = None,
    bend_radius: float | None = None,
    relief_type: str = "Default",
) -> dict:
    """Create a sheet metal base feature.

    type: 'flange' | 'tab' | 'contour_advanced'
        | 'tab_multi_profile'

    Parameters (used per type):
        thickness: Material thickness.
        width: Flange/tab width (flange, tab).
        bend_radius: Bend radius (flange, contour_advanced).
        relief_type: Relief type string (contour_advanced).
    """
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


# =====================================================================
# Group 28: create_lofted_flange (3 -> 1)
# =====================================================================


def create_lofted_flange(
    method: str = "basic",
    thickness: float = 0.0,
    bend_radius: float = 0.0,
) -> dict:
    """Create a lofted flange (sheet metal).

    method: 'basic' | 'advanced' | 'ex'

    Parameters:
        thickness: Material thickness.
        bend_radius: Bend radius (advanced).
    """
    match method:
        case "basic":
            return feature_manager.create_lofted_flange(thickness)
        case "advanced":
            return feature_manager.create_lofted_flange_advanced(thickness, bend_radius)
        case "ex":
            return feature_manager.create_lofted_flange_ex(thickness)
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 29: create_bend (2 -> 1)
# =====================================================================


def create_bend(
    method: str = "basic",
    bend_angle: float = 90.0,
    direction: str = "Normal",
    moving_side: str = "Right",
    bend_deduction: float = 0.0,
) -> dict:
    """Create a bend feature on sheet metal.

    method: 'basic' | 'with_calc'

    Parameters:
        bend_angle: Bend angle in degrees.
        direction: Direction.
        moving_side: 'Right' or 'Left'.
        bend_deduction: Bend deduction value (with_calc).
    """
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


# =====================================================================
# Group 30: create_slot (5 -> 1)
# =====================================================================


def create_slot(
    method: str = "basic",
    width: float = 0.0,
    depth: float = 0.0,
    direction: str = "Normal",
) -> dict:
    """Create a slot feature.

    method: 'basic' | 'ex' | 'sync' | 'multi_body'
        | 'sync_multi_body'

    Parameters:
        width: Slot width.
        depth: Slot depth (ex, sync, multi_body,
            sync_multi_body).
        direction: Direction (basic, ex, multi_body,
            sync_multi_body).
    """
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


# =====================================================================
# Group 31: create_thread (2 -> 1)
# =====================================================================


def create_thread(
    method: str = "basic",
    face_index: int = 0,
    thread_diameter: float = 0.0,
    thread_depth: float = 0.0,
) -> dict:
    """Create a thread on a cylindrical face.

    method: 'basic' (cosmetic thread) | 'physical' (modeled thread geometry)

    Parameters:
        face_index: 0-based cylindrical face index.
        thread_diameter: Thread diameter in meters (0 = auto-detect from face).
        thread_depth: Thread depth in meters (0 = full depth).
    """
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


# =====================================================================
# Group 32: create_drawn_cutout (2 -> 1)
# =====================================================================


def create_drawn_cutout(
    method: str = "basic",
    depth: float = 0.0,
    direction: str = "Normal",
) -> dict:
    """Create a drawn cutout (sheet metal).

    method: 'basic' | 'ex'

    Parameters:
        depth: Cut depth.
        direction: Direction.
    """
    match method:
        case "basic":
            return feature_manager.create_drawn_cutout(depth, direction)
        case "ex":
            return feature_manager.create_drawn_cutout_ex(depth, direction)
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 33: create_dimple (2 -> 1)
# =====================================================================


def create_dimple(
    method: str = "basic",
    depth: float = 0.0,
    direction: str = "Normal",
    punch_tool_diameter: float = 0.01,
) -> dict:
    """Create a dimple feature.

    method: 'basic' | 'ex'

    Parameters:
        depth: Dimple depth.
        direction: Direction.
        punch_tool_diameter: Punch tool diameter (ex).
    """
    match method:
        case "basic":
            return feature_manager.create_dimple(depth, direction)
        case "ex":
            return feature_manager.create_dimple_ex(depth, direction, punch_tool_diameter)
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 34: create_louver (2 -> 1)
# =====================================================================


def create_louver(
    method: str = "basic",
    depth: float = 0.0,
) -> dict:
    """Create a louver feature.

    method: 'basic' | 'sync'

    Parameters:
        depth: Louver depth.
    """
    match method:
        case "basic":
            return feature_manager.create_louver(depth)
        case "sync":
            return feature_manager.create_louver_sync(depth)
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 35: create_pattern (10 -> 1)
# =====================================================================


def create_pattern(
    method: str = "rectangular_ex",
    feature_index: int = 0,
    feature_name: str = "",
    x_count: int = 1,
    y_count: int = 1,
    x_gap: float = 0.0,
    y_gap: float = 0.0,
    x_spacing: float = 0.0,
    y_spacing: float = 0.0,
    count: int = 1,
    angle: float = 0.0,
    radius: float = 0.0,
    axis_face_index: int = 0,
    fill_region_face_index: int = 0,
    x_offsets: list[float] | None = None,
    y_offsets: list[float] | None = None,
    curve_edge_index: int = 0,
    spacing: float = 0.0,
) -> dict:
    """Create a pattern of a feature.

    method: 'rectangular_ex' | 'rectangular' | 'circular'
        | 'circular_ex' | 'duplicate' | 'by_fill'
        | 'by_table' | 'by_table_sync' | 'by_fill_ex'
        | 'by_curve_ex' | 'user_defined'

    Parameters (used per method):
        feature_index: 0-based feature index (rectangular,
            circular).
        feature_name: Feature name string (rectangular_ex,
            circular_ex, duplicate, by_fill, by_table,
            by_table_sync, by_fill_ex, by_curve_ex,
            user_defined).
        x_count, y_count: Pattern counts (rectangular,
            rectangular_ex).
        x_gap, y_gap: Gap between instances (rectangular).
        x_spacing, y_spacing: Spacing (rectangular_ex,
            by_fill, by_fill_ex).
        count: Instance count (circular, circular_ex,
            by_curve_ex).
        angle: Pattern angle (circular, circular_ex).
        radius: Pattern radius (circular).
        axis_face_index: Axis face (circular_ex).
        fill_region_face_index: Fill region face (by_fill,
            by_fill_ex).
        x_offsets, y_offsets: Offset lists (by_table,
            by_table_sync).
        curve_edge_index: Curve edge (by_curve_ex).
        spacing: Spacing along curve (by_curve_ex).
    """
    match method:
        case "rectangular_ex":
            return feature_manager.create_pattern_rectangular_ex(
                feature_name,
                x_count,
                y_count,
                x_spacing,
                y_spacing,
            )
        case "rectangular":
            return feature_manager.create_pattern_rectangular(
                feature_index,
                x_count,
                y_count,
                x_gap,
                y_gap,
            )
        case "circular":
            return feature_manager.create_pattern_circular(feature_index, count, angle, radius)
        case "circular_ex":
            return feature_manager.create_pattern_circular_ex(
                feature_name,
                count,
                angle,
                axis_face_index,
            )
        case "duplicate":
            return feature_manager.create_pattern_duplicate(feature_name)
        case "by_fill":
            return feature_manager.create_pattern_by_fill(
                feature_name,
                fill_region_face_index,
                x_spacing,
                y_spacing,
            )
        case "by_table":
            return feature_manager.create_pattern_by_table(feature_name, x_offsets, y_offsets)
        case "by_table_sync":
            return feature_manager.create_pattern_by_table_sync(feature_name, x_offsets, y_offsets)
        case "by_fill_ex":
            return feature_manager.create_pattern_by_fill_ex(
                feature_name,
                fill_region_face_index,
                x_spacing,
                y_spacing,
            )
        case "by_curve_ex":
            return feature_manager.create_pattern_by_curve_ex(
                feature_name,
                curve_edge_index,
                count,
                spacing,
            )
        case "user_defined":
            return feature_manager.create_user_defined_pattern(feature_name)
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 36: create_mirror (2 -> 1)
# =====================================================================


def create_mirror(
    method: str = "basic",
    feature_name: str = "",
    mirror_plane_index: int = 0,
    new_file_name: str = "",
    link_to_original: bool = True,
) -> dict:
    """Mirror a feature across a reference plane, or save as mirror part.

    method: 'basic' | 'sync_ex' | 'save_as_part'

    Parameters (used per method):
        feature_name: Name of the feature to mirror (basic,
            sync_ex).
        mirror_plane_index: 1-based ref plane index (basic,
            sync_ex, save_as_part).
        new_file_name: Full path for mirrored part .par
            (save_as_part).
        link_to_original: Maintain link to original
            (save_as_part).
    """
    match method:
        case "basic":
            return feature_manager.create_mirror(feature_name, mirror_plane_index)
        case "sync_ex":
            return feature_manager.create_mirror_sync_ex(feature_name, mirror_plane_index)
        case "save_as_part":
            return feature_manager.save_as_mirror_part(
                new_file_name, mirror_plane_index or 3, link_to_original
            )
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 37: create_thin_wall (2 -> 1)
# =====================================================================


def create_thin_wall(
    method: str = "basic",
    thickness: float = 0.0,
    open_face_indices: list[int] | None = None,
) -> dict:
    """Convert a solid body to a thin wall (shell).

    method: 'basic' | 'with_open_faces'

    Parameters:
        thickness: Shell wall thickness.
        open_face_indices: Faces to leave open
            (with_open_faces).
    """
    match method:
        case "basic":
            return feature_manager.create_shell(thickness)
        case "with_open_faces":
            return feature_manager.create_shell(thickness, open_face_indices)
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 38: face_operation (2 -> 1)
# =====================================================================


def face_operation(
    type: str = "rotate_by_points",
    face_index: int = 0,
    vertex1_index: int = 0,
    vertex2_index: int = 0,
    edge_index: int = 0,
    angle: float = 0.0,
) -> dict:
    """Perform a face operation (rotate).

    type: 'rotate_by_points' | 'rotate_by_edge'

    Parameters (used per type):
        face_index: Face to rotate.
        vertex1_index, vertex2_index: Axis vertices
            (rotate_by_points).
        edge_index: Axis edge (rotate_by_edge).
        angle: Rotation angle in degrees.
    """
    match type:
        case "rotate_by_points":
            return feature_manager.create_face_rotate_by_points(
                face_index,
                vertex1_index,
                vertex2_index,
                angle,
            )
        case "rotate_by_edge":
            return feature_manager.create_face_rotate_by_edge(face_index, edge_index, angle)
        case _:
            return {"error": f"Unknown type: {type}"}


# =====================================================================
# Group 39: add_body (5 -> 1)
# =====================================================================


def add_body(
    method: str = "basic",
    body_type: str = "Solid",
    tag: str = "",
) -> dict:
    """Add a body to the part.

    method: 'basic' | 'by_mesh' | 'feature' | 'construction'
        | 'by_tag'

    Parameters (used per method):
        body_type: Body type string (basic).
        tag: Tag reference string (by_tag).
    """
    match method:
        case "basic":
            return feature_manager.add_body(body_type)
        case "by_mesh":
            return feature_manager.add_body_by_mesh()
        case "feature":
            return feature_manager.add_body_feature()
        case "construction":
            return feature_manager.add_by_construction()
        case "by_tag":
            return feature_manager.add_body_by_tag(tag)
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 40: simplify (4 -> 1)
# =====================================================================


def simplify(
    method: str = "auto",
) -> dict:
    """Simplify the model.

    method: 'auto' | 'enclosure' | 'duplicate'
        | 'local_enclosure'
    """
    match method:
        case "auto":
            return feature_manager.auto_simplify()
        case "enclosure":
            return feature_manager.simplify_enclosure()
        case "duplicate":
            return feature_manager.simplify_duplicate()
        case "local_enclosure":
            return feature_manager.local_simplify_enclosure()
        case _:
            return {"error": f"Unknown method: {method}"}


# =====================================================================
# Group 41: manage_feature (6 -> 1)
# =====================================================================


def manage_feature(
    action: str = "delete",
    index: int = 0,
    target_index: int = 0,
    after: bool = True,
    new_name: str = "",
    feature_name: str = "",
    target_type: str = "",
) -> dict:
    """Manage features in the feature tree.

    action: 'delete' | 'suppress' | 'unsuppress' | 'reorder'
        | 'rename' | 'convert'

    Parameters (used per action):
        index: Feature index (delete, suppress, unsuppress,
            reorder, rename).
        target_index: Target position index (reorder).
        after: Insert after target (reorder).
        new_name: New feature name (rename).
        feature_name: Feature name (convert).
        target_type: Target type string (convert).
    """
    match action:
        case "delete":
            return feature_manager.delete_feature(index)
        case "suppress":
            return feature_manager.feature_suppress(index)
        case "unsuppress":
            return feature_manager.feature_unsuppress(index)
        case "reorder":
            return feature_manager.feature_reorder(index, target_index, after)
        case "rename":
            return feature_manager.feature_rename(index, new_name)
        case "convert":
            return feature_manager.convert_feature_type(feature_name, target_type)
        case _:
            return {"error": f"Unknown action: {action}"}


# =====================================================================
# Group 42: sheet_metal_misc (5 -> 1)
# =====================================================================


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
) -> dict:
    """Miscellaneous sheet metal operations.

    action: 'hem' | 'jog' | 'close_corner' | 'multi_edge_flange'
        | 'convert'

    Parameters (used per action):
        face_index: Face index (hem, close_corner,
            multi_edge_flange).
        edge_index: Edge index (hem, close_corner).
        hem_width: Hem width (hem).
        bend_radius: Bend radius (hem).
        hem_type: 'Closed'|'Open'|etc. (hem).
        jog_offset: Jog offset distance (jog).
        jog_angle: Jog angle in degrees (jog).
        direction: Direction (jog).
        moving_side: 'Right' or 'Left' (jog).
        closure_type: 'Close'|'Overlap'|etc. (close_corner).
        edge_indices: List of edge indices (multi_edge_flange).
        flange_length: Flange length (multi_edge_flange).
        side: 'Right' or 'Left' (multi_edge_flange).
        thickness: Material thickness (convert).
    """
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


# =====================================================================
# Group 43: create_stamped (2 -> 1)
# =====================================================================


def create_stamped(
    type: str = "bead",
    depth: float = 0.0,
) -> dict:
    """Create a stamped feature (bead or gusset).

    type: 'bead' | 'gusset'

    Parameters:
        depth: Stamping depth.
    """
    match type:
        case "bead":
            return feature_manager.create_bead(depth)
        case "gusset":
            return feature_manager.create_gusset(depth)
        case _:
            return {"error": f"Unknown type: {type}"}


# =====================================================================
# Group 44: create_surface_mark (2 -> 1)
# =====================================================================


def create_surface_mark(
    type: str = "emboss",
    face_indices: list | None = None,
    clearance: float = 0.001,
    thickness: float = 0.0,
    thicken: bool = False,
    default_side: bool = True,
) -> dict:
    """Create a surface mark (emboss or etch).

    type: 'emboss' | 'etch'

    Parameters (used per type):
        face_indices: Faces to emboss onto (emboss).
        clearance: Clearance distance (emboss).
        thickness: Emboss thickness (emboss).
        thicken: Whether to thicken (emboss).
        default_side: Use default side (emboss).
    """
    match type:
        case "emboss":
            return feature_manager.create_emboss(
                face_indices or [], clearance, thickness, thicken, default_side
            )
        case "etch":
            return feature_manager.create_etch()
        case _:
            return {"error": f"Unknown type: {type}"}


# =====================================================================
# Group 45: create_reinforcement (2 -> 1)
# =====================================================================


def create_reinforcement(
    type: str = "rib",
    thickness: float = 0.0,
    direction: str = "Normal",
) -> dict:
    """Create a reinforcement feature (rib or lip).

    type: 'rib' | 'lip'

    Parameters:
        thickness: Wall thickness.
        direction: Direction (rib only).
    """
    match type:
        case "rib":
            return feature_manager.create_rib(thickness, direction)
        case "lip":
            return feature_manager.create_lip(thickness)
        case _:
            return {"error": f"Unknown type: {type}"}


# =====================================================================
# Standalone tools (unique signatures, not merged)
# =====================================================================


def create_web_network() -> dict:
    """Create a web network (sheet metal)."""
    return feature_manager.create_web_network()


def create_split() -> dict:
    """Split the body using the active profile."""
    return feature_manager.create_split()


def create_draft_angle(
    face_index: int,
    angle: float,
    plane_index: int = 1,
) -> dict:
    """Add a draft angle to a face."""
    return feature_manager.create_draft_angle(face_index, angle, plane_index)


def create_bounded_surface(
    want_end_caps: bool = True,
    periodic: bool = False,
) -> dict:
    """Create a bounded (blue) surface from accumulated profiles.

    Args:
        want_end_caps: Whether to add end caps.
        periodic: Whether the surface is periodic.
    """
    return feature_manager.create_bounded_surface(want_end_caps, periodic)


# =====================================================================
# Registration
# =====================================================================


def register(mcp):
    """Register feature tools with the MCP server."""
    # Composite tools (45 groups)
    mcp.tool()(create_extrude)
    mcp.tool()(create_revolve)
    mcp.tool()(create_extruded_cutout)
    mcp.tool()(create_revolved_cutout)
    mcp.tool()(create_normal_cutout)
    mcp.tool()(create_lofted_cutout)
    mcp.tool()(create_swept_cutout)
    mcp.tool()(create_helix)
    mcp.tool()(create_helix_cutout)
    mcp.tool()(create_loft)
    mcp.tool()(create_sweep)
    mcp.tool()(create_extruded_surface)
    mcp.tool()(create_revolved_surface)
    mcp.tool()(create_lofted_surface)
    mcp.tool()(create_swept_surface)
    mcp.tool()(thicken)
    mcp.tool()(create_primitive)
    mcp.tool()(create_primitive_cutout)
    mcp.tool()(create_hole)
    mcp.tool()(create_round)
    mcp.tool()(create_chamfer)
    mcp.tool()(create_blend)
    mcp.tool()(delete_topology)
    mcp.tool()(create_ref_plane)
    mcp.tool()(create_flange)
    mcp.tool()(create_contour_flange)
    mcp.tool()(create_sheet_metal_base)
    mcp.tool()(create_lofted_flange)
    mcp.tool()(create_bend)
    mcp.tool()(create_slot)
    mcp.tool()(create_thread)
    mcp.tool()(create_drawn_cutout)
    mcp.tool()(create_dimple)
    mcp.tool()(create_louver)
    mcp.tool()(create_pattern)
    mcp.tool()(create_mirror)
    mcp.tool()(create_thin_wall)
    mcp.tool()(face_operation)
    mcp.tool()(add_body)
    mcp.tool()(simplify)
    mcp.tool()(manage_feature)
    mcp.tool()(sheet_metal_misc)
    mcp.tool()(create_stamped)
    mcp.tool()(create_surface_mark)
    mcp.tool()(create_reinforcement)
    # Standalone tools
    mcp.tool()(create_web_network)
    mcp.tool()(create_split)
    mcp.tool()(create_draft_angle)
    mcp.tool()(create_bounded_surface)
