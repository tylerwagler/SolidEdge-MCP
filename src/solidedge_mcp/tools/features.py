"""Feature modeling tools for Solid Edge MCP."""

from solidedge_mcp.managers import diagnose_feature, feature_manager

# === Extrusions ===


def create_extrude(distance: float, direction: str = "Normal") -> dict:
    """Create an extruded protrusion from the active sketch profile."""
    return feature_manager.create_extrude(distance, direction)


def create_extrude_infinite(direction: str = "Normal") -> dict:
    """Create an infinite extrusion (extends through all)."""
    return feature_manager.create_extrude_infinite(direction)


def create_extrude_through_next(direction: str = "Normal") -> dict:
    """Create an extrusion that extends to the next face encountered."""
    return feature_manager.create_extrude_through_next(direction)


def create_extrude_from_to(from_plane_index: int, to_plane_index: int) -> dict:
    """Create an extrusion between two reference planes (1-based indices)."""
    return feature_manager.create_extrude_from_to(from_plane_index, to_plane_index)


def create_extrude_thin_wall(
    distance: float, wall_thickness: float, direction: str = "Normal"
) -> dict:
    """Create a thin-walled extrusion."""
    return feature_manager.create_extrude_thin_wall(distance, wall_thickness, direction)


# === Revolves ===


def create_revolve(angle: float = 360.0) -> dict:
    """Create a revolved protrusion around the set axis."""
    return feature_manager.create_revolve(angle)


def create_revolve_finite(angle: float, axis_type: str = "CenterLine") -> dict:
    """Create a finite revolve feature."""
    return feature_manager.create_revolve_finite(angle, axis_type)


def create_revolve_sync(angle: float) -> dict:
    """Create a synchronous revolve feature."""
    return feature_manager.create_revolve_sync(angle)


def create_revolve_finite_sync(angle: float) -> dict:
    """Create a finite synchronous revolve feature."""
    return feature_manager.create_revolve_finite_sync(angle)


def create_revolve_thin_wall(angle: float, wall_thickness: float) -> dict:
    """Create a thin-walled revolve feature."""
    return feature_manager.create_revolve_thin_wall(angle, wall_thickness)


# === Cutouts ===


def create_extruded_cutout(distance: float, direction: str = "Normal") -> dict:
    """Create an extruded cutout (removes material)."""
    return feature_manager.create_extruded_cutout(distance, direction)


def create_extruded_cutout_through_all(direction: str = "Normal") -> dict:
    """Create an extruded cutout through all material."""
    return feature_manager.create_extruded_cutout_through_all(direction)


def create_extruded_cutout_through_next(direction: str = "Normal") -> dict:
    """Create an extruded cutout through the next face."""
    return feature_manager.create_extruded_cutout_through_next(direction)


def create_revolved_cutout(angle: float = 360.0) -> dict:
    """Create a revolved cutout around set axis."""
    return feature_manager.create_revolved_cutout(angle)


def create_normal_cutout(distance: float, direction: str = "Normal") -> dict:
    """Create a normal cutout perpendicular to a face."""
    return feature_manager.create_normal_cutout(distance, direction)


def create_normal_cutout_through_all(direction: str = "Normal") -> dict:
    """Create a normal cutout through all material."""
    return feature_manager.create_normal_cutout_through_all(direction)


def create_lofted_cutout(profile_indices: list[int] | None = None) -> dict:
    """Create a lofted cutout between multiple profiles."""
    return feature_manager.create_lofted_cutout(profile_indices)


def create_swept_cutout() -> dict:
    """Create a swept cutout along a path."""
    return feature_manager.create_swept_cutout()


def create_extruded_cutout_from_to(from_plane_index: int, to_plane_index: int) -> dict:
    """Create an extruded cutout between two reference planes (1-based indices)."""
    return feature_manager.create_extruded_cutout_from_to(from_plane_index, to_plane_index)


def create_drawn_cutout(depth: float, direction: str = "Normal") -> dict:
    """Create a drawn cutout (sheet metal)."""
    return feature_manager.create_drawn_cutout(depth, direction)


# === Loft / Sweep / Helix ===


def create_loft(profile_indices: list | None = None) -> dict:
    """Create a loft feature between multiple profiles."""
    return feature_manager.create_loft(profile_indices)


def create_loft_thin_wall(wall_thickness: float, profile_indices: list | None = None) -> dict:
    """Create a thin-walled loft feature."""
    return feature_manager.create_loft_thin_wall(wall_thickness, profile_indices)


def create_sweep(path_profile_index: int | None = None) -> dict:
    """Create a sweep feature along a path."""
    return feature_manager.create_sweep(path_profile_index)


def create_sweep_thin_wall(wall_thickness: float, path_profile_index: int | None = None) -> dict:
    """Create a thin-walled sweep feature."""
    return feature_manager.create_sweep_thin_wall(wall_thickness, path_profile_index)


def create_helix(
    pitch: float, height: float, revolutions: float | None = None, direction: str = "Right"
) -> dict:
    """Create a helical feature (springs, threads)."""
    return feature_manager.create_helix(pitch, height, revolutions, direction)


def create_helix_sync(pitch: float, height: float, revolutions: float | None = None) -> dict:
    """Create a synchronous helical feature."""
    return feature_manager.create_helix_sync(pitch, height, revolutions)


def create_helix_thin_wall(
    pitch: float, height: float, wall_thickness: float, revolutions: float = None
) -> dict:
    """Create a thin-walled helix feature."""
    return feature_manager.create_helix_thin_wall(pitch, height, wall_thickness, revolutions)


def create_helix_sync_thin_wall(
    pitch: float, height: float, wall_thickness: float, revolutions: float = None
) -> dict:
    """Create a synchronous thin-walled helix feature."""
    return feature_manager.create_helix_sync_thin_wall(pitch, height, wall_thickness, revolutions)


def create_helix_cutout(
    pitch: float, height: float, revolutions: float = None, direction: str = "Right"
) -> dict:
    """Create a helical cutout."""
    return feature_manager.create_helix_cutout(pitch, height, revolutions, direction)


# === Surfaces ===


def create_extruded_surface(
    distance: float, direction: str = "Normal", end_caps: bool = True
) -> dict:
    """Create an extruded surface."""
    return feature_manager.create_extruded_surface(distance, direction, end_caps)


def create_revolved_surface(angle: float = 360, want_end_caps: bool = False) -> dict:
    """Create a revolved surface."""
    return feature_manager.create_revolved_surface(angle, want_end_caps)


def create_lofted_surface(want_end_caps: bool = False) -> dict:
    """Create a lofted surface."""
    return feature_manager.create_lofted_surface(want_end_caps)


def create_swept_surface(
    path_profile_index: int | None = None, want_end_caps: bool = False
) -> dict:
    """Create a swept surface."""
    return feature_manager.create_swept_surface(path_profile_index, want_end_caps)


def thicken_surface(thickness: float, direction: str = "Both") -> dict:
    """Thicken a surface to create a solid."""
    return feature_manager.thicken_surface(thickness, direction)


# === Primitives ===


def create_box_by_center(
    x: float, y: float, z: float, length: float, width: float, height: float, plane_index: int = 1
) -> dict:
    """Create a box by center point and dimensions.

    plane_index: 1=Top/XZ, 2=Front/XY, 3=Right/YZ.
    """
    return feature_manager.create_box_by_center(x, y, z, length, width, height, plane_index)


def create_box_by_two_points(
    x1: float, y1: float, z1: float, x2: float, y2: float, z2: float, plane_index: int = 1
) -> dict:
    """Create a box by two opposite corners.

    plane_index: 1=Top/XZ, 2=Front/XY, 3=Right/YZ.
    """
    return feature_manager.create_box_by_two_points(x1, y1, z1, x2, y2, z2, plane_index)


def create_box_by_three_points(
    x1: float,
    y1: float,
    z1: float,
    x2: float,
    y2: float,
    z2: float,
    x3: float,
    y3: float,
    z3: float,
    plane_index: int = 1,
) -> dict:
    """Create a box by three points.

    plane_index: 1=Top/XZ, 2=Front/XY, 3=Right/YZ.
    """
    return feature_manager.create_box_by_three_points(
        x1, y1, z1, x2, y2, z2, x3, y3, z3, plane_index
    )


def create_cylinder(
    base_center_x: float,
    base_center_y: float,
    base_center_z: float,
    radius: float,
    height: float,
    plane_index: int = 1,
) -> dict:
    """Create a cylinder primitive.

    plane_index: 1=Top/XZ, 2=Front/XY, 3=Right/YZ.
    """
    return feature_manager.create_cylinder(
        base_center_x, base_center_y, base_center_z, radius, height, plane_index
    )


def create_sphere(
    center_x: float, center_y: float, center_z: float, radius: float, plane_index: int = 1
) -> dict:
    """Create a sphere primitive.

    plane_index: 1=Top/XZ, 2=Front/XY, 3=Right/YZ.
    """
    return feature_manager.create_sphere(center_x, center_y, center_z, radius, plane_index)


# === Primitive Cutouts ===


def create_box_cutout(
    x1: float, y1: float, z1: float, x2: float, y2: float, z2: float, plane_index: int = 1
) -> dict:
    """Create a box-shaped cutout by two opposite corners.

    plane_index: 1=Top/XZ, 2=Front/XY, 3=Right/YZ.
    """
    return feature_manager.create_box_cutout_by_two_points(x1, y1, z1, x2, y2, z2, plane_index)


def create_cylinder_cutout(
    center_x: float,
    center_y: float,
    center_z: float,
    radius: float,
    height: float,
    plane_index: int = 1,
) -> dict:
    """Create a cylindrical cutout.

    plane_index: 1=Top/XZ, 2=Front/XY, 3=Right/YZ.
    """
    return feature_manager.create_cylinder_cutout(
        center_x, center_y, center_z, radius, height, plane_index
    )


def create_sphere_cutout(
    center_x: float, center_y: float, center_z: float, radius: float, plane_index: int = 1
) -> dict:
    """Create a spherical cutout.

    plane_index: 1=Top/XZ, 2=Front/XY, 3=Right/YZ.
    """
    return feature_manager.create_sphere_cutout(center_x, center_y, center_z, radius, plane_index)


# === Holes ===


def create_hole(
    x: float,
    y: float,
    diameter: float,
    depth: float = 0.0,
    hole_type: str = "Simple",
    plane_index: int = 1,
    direction: str = "Normal",
) -> dict:
    """Create a hole at coordinates (meters)."""
    return feature_manager.create_hole(x, y, diameter, depth, direction)


def create_hole_through_all(
    x: float, y: float, diameter: float, plane_index: int = 1, direction: str = "Normal"
) -> dict:
    """Create a hole through all material."""
    return feature_manager.create_hole_through_all(x, y, diameter, plane_index, direction)


def create_delete_hole(max_diameter: float = 1.0, hole_type: str = "All") -> dict:
    """Delete/fill holes by size criteria."""
    return feature_manager.create_delete_hole(max_diameter, hole_type)


def create_delete_hole_by_face(face_index: int) -> dict:
    """Delete a specific hole by its face."""
    return feature_manager.delete_hole_by_face(face_index)


# === Rounds / Chamfers ===


def create_round(radius: float) -> dict:
    """Round (fillet) all edges of the active body."""
    return feature_manager.create_round(radius)


def create_round_on_face(radius: float, face_index: int) -> dict:
    """Round all edges of a specific face."""
    return feature_manager.create_round_on_face(radius, face_index)


def create_variable_round(radii: list, face_index: int | None = None) -> dict:
    """Create a variable-radius round."""
    return feature_manager.create_variable_round(radii, face_index)


def create_chamfer(distance: float) -> dict:
    """Chamfer all edges of the active body."""
    return feature_manager.create_chamfer(distance)


def create_chamfer_on_face(distance: float, face_index: int) -> dict:
    """Chamfer all edges of a specific face."""
    return feature_manager.create_chamfer_on_face(distance, face_index)


def create_chamfer_unequal(distance1: float, distance2: float, face_index: int = 0) -> dict:
    """Create an unequal chamfer."""
    return feature_manager.create_chamfer_unequal(distance1, distance2, face_index)


def create_chamfer_angle(distance: float, angle: float, face_index: int = 0) -> dict:
    """Create a chamfer by distance and angle."""
    return feature_manager.create_chamfer_angle(distance, angle, face_index)


# === Blends / Draft / Mirror ===


def create_blend(radius: float, face_index: int | None = None) -> dict:
    """Create a blend (face-to-face fillet)."""
    return feature_manager.create_blend(radius, face_index)


def create_delete_blend(face_index: int) -> dict:
    """Delete a blend feature."""
    return feature_manager.create_delete_blend(face_index)


def create_draft_angle(face_index: int, angle: float, plane_index: int = 1) -> dict:
    """Add a draft angle to a face."""
    return feature_manager.create_draft_angle(face_index, angle, plane_index)


def create_mirror(feature_name: str, mirror_plane_index: int) -> dict:
    """Mirror a feature across a reference plane."""
    return feature_manager.create_mirror(feature_name, mirror_plane_index)


# === Thin Wall (Shell) ===


def create_thin_wall(thickness: float) -> dict:
    """Convert a solid body to a thin wall (shell)."""
    return feature_manager.create_shell(thickness)


def create_thin_wall_with_open_faces(thickness: float, open_face_indices: list[int]) -> dict:
    """Create a thin wall with specified open faces."""
    return feature_manager.create_shell(thickness, open_face_indices)


# === Part Features (rib, lip, slot, split, thread, etc.) ===


def create_rib(thickness: float, direction: str = "Normal") -> dict:
    """Create a rib feature."""
    return feature_manager.create_rib(thickness, direction)


def create_lip(thickness: float) -> dict:
    """Create a lip feature."""
    return feature_manager.create_lip(thickness)


def create_slot(width: float, direction: str = "Normal") -> dict:
    """Create a slot feature."""
    return feature_manager.create_slot(width, direction)


def create_split() -> dict:
    """Split the body using the active profile."""
    return feature_manager.create_split()


def create_thread(face_index: int, depth: float) -> dict:
    """Create a thread on a cylindrical face."""
    return feature_manager.create_thread(face_index, depth)


def create_emboss(
    face_indices: list,
    clearance: float = 0.001,
    thickness: float = 0.0,
    thicken: bool = False,
    default_side: bool = True,
) -> dict:
    """Create an emboss feature."""
    return feature_manager.create_emboss(face_indices, clearance, thickness, thicken, default_side)


def create_etch() -> dict:
    """Create an etch feature from the active profile."""
    return feature_manager.create_etch()


def create_flange(
    face_index: int,
    edge_index: int,
    flange_length: float,
    side: str = "Right",
    inside_radius: float = None,
    bend_angle: float = None,
) -> dict:
    """Create a flange on an edge."""
    return feature_manager.create_flange(
        face_index, edge_index, flange_length, side, inside_radius, bend_angle
    )


def create_dimple(depth: float, direction: str = "Normal") -> dict:
    """Create a dimple feature."""
    return feature_manager.create_dimple(depth, direction)


def create_bead(depth: float) -> dict:
    """Create a bead feature."""
    return feature_manager.create_bead(depth)


def create_louver(depth: float) -> dict:
    """Create a louver feature."""
    return feature_manager.create_louver(depth)


def create_gusset(depth: float) -> dict:
    """Create a gusset feature."""
    return feature_manager.create_gusset(depth)


# === Patterns ===


def create_pattern_rectangular(
    feature_index: int, x_count: int, y_count: int, x_gap: float, y_gap: float
) -> dict:
    """Create a rectangular pattern of a feature."""
    return feature_manager.create_pattern_rectangular(feature_index, x_count, y_count, x_gap, y_gap)


def create_pattern_circular(feature_index: int, count: int, angle: float, radius: float) -> dict:
    """Create a circular pattern of a feature."""
    return feature_manager.create_pattern_circular(feature_index, count, angle, radius)


# === Face Operations ===


def delete_faces(face_indices: list[int]) -> dict:
    """Delete specified faces from the body."""
    return feature_manager.delete_faces(face_indices)


def create_face_rotate_by_points(
    face_index: int, vertex1_index: int, vertex2_index: int, angle: float
) -> dict:
    """Rotate a face around an axis defined by two vertices."""
    return feature_manager.create_face_rotate_by_points(
        face_index, vertex1_index, vertex2_index, angle
    )


def create_face_rotate_by_edge(face_index: int, edge_index: int, angle: float) -> dict:
    """Rotate a face around an edge."""
    return feature_manager.create_face_rotate_by_edge(face_index, edge_index, angle)


# === Reference Planes ===


def create_ref_plane_by_offset(
    parent_plane_index: int, distance: float, normal_side: str = "Normal"
) -> dict:
    """Create a reference plane offset from an existing plane."""
    return feature_manager.create_ref_plane_by_offset(parent_plane_index, distance, normal_side)


def create_ref_plane_by_angle(
    parent_plane_index: int, angle: float, normal_side: str = "Normal"
) -> dict:
    """Create a reference plane at an angle to an existing plane."""
    return feature_manager.create_ref_plane_by_angle(parent_plane_index, angle, normal_side)


def create_ref_plane_by_3_points(
    x1: float,
    y1: float,
    z1: float,
    x2: float,
    y2: float,
    z2: float,
    x3: float,
    y3: float,
    z3: float,
) -> dict:
    """Create a reference plane defined by three points."""
    return feature_manager.create_ref_plane_by_3_points(x1, y1, z1, x2, y2, z2, x3, y3, z3)


def create_ref_plane_midplane(plane1_index: int, plane2_index: int) -> dict:
    """Create a midplane between two existing planes."""
    return feature_manager.create_ref_plane_midplane(plane1_index, plane2_index)


def create_ref_plane_normal_to_curve(curve_end: str = "End", pivot_plane_index: int = 2) -> dict:
    """Create a reference plane normal to a curve endpoint."""
    return feature_manager.create_ref_plane_normal_to_curve(curve_end, pivot_plane_index)


def create_ref_plane_normal_at_distance(
    distance: float, curve_end: str = "End", pivot_plane_index: int = 2
) -> dict:
    """Create a reference plane normal to a curve at a distance from an endpoint."""
    return feature_manager.create_ref_plane_normal_at_distance(
        distance, curve_end, pivot_plane_index
    )


def create_ref_plane_normal_at_arc_ratio(
    ratio: float, curve_end: str = "End", pivot_plane_index: int = 2
) -> dict:
    """Create a reference plane normal to a curve at an arc-length ratio (0.0 to 1.0)."""
    return feature_manager.create_ref_plane_normal_at_arc_ratio(ratio, curve_end, pivot_plane_index)


def create_ref_plane_normal_at_distance_along(
    distance_along: float, curve_end: str = "End", pivot_plane_index: int = 2
) -> dict:
    """Create a reference plane normal to a curve at a distance along the curve."""
    return feature_manager.create_ref_plane_normal_at_distance_along(
        distance_along, curve_end, pivot_plane_index
    )


def create_ref_plane_parallel_by_tangent(
    parent_plane_index: int, face_index: int, normal_side: str = "Normal"
) -> dict:
    """Create a reference plane parallel to a parent plane and tangent to a face."""
    return feature_manager.create_ref_plane_parallel_by_tangent(
        parent_plane_index, face_index, normal_side
    )


# === Body Operations ===


def add_body(body_type: str = "Solid") -> dict:
    """Add a body to the part."""
    return feature_manager.add_body(body_type)


def add_body_by_mesh() -> dict:
    """Add a body by mesh facets."""
    return feature_manager.add_body_by_mesh()


def add_body_feature() -> dict:
    """Add a body feature."""
    return feature_manager.add_body_feature()


def add_by_construction() -> dict:
    """Add a construction body."""
    return feature_manager.add_by_construction()


def add_body_by_tag(tag: str) -> dict:
    """Add a body by tag reference."""
    return feature_manager.add_body_by_tag(tag)


# === Simplification ===


def auto_simplify() -> dict:
    """Auto-simplify the model (reduce complexity)."""
    return feature_manager.auto_simplify()


def simplify_enclosure() -> dict:
    """Create a simplified enclosure around the model."""
    return feature_manager.simplify_enclosure()


def simplify_duplicate() -> dict:
    """Create a simplified duplicate of the model."""
    return feature_manager.simplify_duplicate()


def local_simplify_enclosure() -> dict:
    """Create a local simplified enclosure."""
    return feature_manager.local_simplify_enclosure()


# === Feature Management ===


def list_features() -> dict:
    """List all features in the active model."""
    return feature_manager.list_features()


def get_feature_info(index: int) -> dict:
    """Get detailed info about a feature by index."""
    return feature_manager.get_feature_info(index)


def delete_feature(index: int) -> dict:
    """Delete a feature by index."""
    return feature_manager.delete_feature(index)


def feature_suppress(index: int) -> dict:
    """Suppress a feature by index."""
    return feature_manager.feature_suppress(index)


def feature_unsuppress(index: int) -> dict:
    """Unsuppress a feature by index."""
    return feature_manager.feature_unsuppress(index)


def feature_reorder(feature_index: int, target_index: int, after: bool = True) -> dict:
    """Reorder a feature in the feature tree."""
    return feature_manager.feature_reorder(feature_index, target_index, after)


def feature_rename(index: int, new_name: str) -> dict:
    """Rename a feature by index."""
    return feature_manager.feature_rename(index, new_name)


def convert_feature_type(feature_name: str, target_type: str) -> dict:
    """Convert a feature between ordered and synchronous types."""
    return feature_manager.convert_feature_type(feature_name, target_type)


# === Sheet Metal (lofted flange, web network, advanced SM) ===


def create_base_flange(width: float, thickness: float, bend_radius: float | None = None) -> dict:
    """Create a base contour flange (sheet metal)."""
    return feature_manager.create_base_flange(width, thickness, bend_radius)


def create_base_tab(thickness: float, width: float | None = None) -> dict:
    """Create a base tab (sheet metal)."""
    return feature_manager.create_base_tab(thickness, width)


def create_lofted_flange(thickness: float) -> dict:
    """Create a lofted flange (sheet metal)."""
    return feature_manager.create_lofted_flange(thickness)


def create_web_network() -> dict:
    """Create a web network (sheet metal)."""
    return feature_manager.create_web_network()


def create_base_contour_flange_advanced(
    thickness: float, bend_radius: float, relief_type: str = "Default"
) -> dict:
    """Create a base contour flange with bend deduction (advanced sheet metal)."""
    return feature_manager.create_base_contour_flange_advanced(thickness, bend_radius, relief_type)


def create_base_tab_multi_profile(thickness: float) -> dict:
    """Create a base tab with multiple profiles (sheet metal)."""
    return feature_manager.create_base_tab_multi_profile(thickness)


def create_lofted_flange_advanced(thickness: float, bend_radius: float) -> dict:
    """Create a lofted flange with bend deduction (advanced sheet metal)."""
    return feature_manager.create_lofted_flange_advanced(thickness, bend_radius)


def create_lofted_flange_ex(thickness: float) -> dict:
    """Create an extended lofted flange (sheet metal)."""
    return feature_manager.create_lofted_flange_ex(thickness)


# === Symmetric Extrusion ===


def create_extrude_symmetric(distance: float) -> dict:
    """Create a symmetric extrusion (extends equally in both directions)."""
    return feature_manager.create_extrude_symmetric(distance)


# === Unequal Chamfer on Face ===


def create_chamfer_unequal_on_face(distance1: float, distance2: float, face_index: int) -> dict:
    """Create an unequal chamfer on all edges of a specific face."""
    return feature_manager.create_chamfer_unequal_on_face(distance1, distance2, face_index)


# === Batch 4: Additional Ref Plane Variants ===


def create_ref_plane_normal_at_keypoint(
    keypoint_type: str = "End", pivot_plane_index: int = 2
) -> dict:
    """Create a reference plane normal to a curve at a keypoint (Start or End)."""
    return feature_manager.create_ref_plane_normal_at_keypoint(keypoint_type, pivot_plane_index)


def create_ref_plane_tangent_cylinder_angle(
    face_index: int, angle: float, parent_plane_index: int = 1
) -> dict:
    """Create a reference plane tangent to a cylindrical/conical face at an angle.

    face_index: 0-based face index. angle: rotation angle in degrees.
    """
    return feature_manager.create_ref_plane_tangent_cylinder_angle(
        face_index, angle, parent_plane_index
    )


def create_ref_plane_tangent_cylinder_keypoint(
    face_index: int, keypoint_type: str = "End", parent_plane_index: int = 1
) -> dict:
    """Create a reference plane tangent to a cylindrical/conical face at a keypoint.

    face_index: 0-based face index. keypoint_type: 'Start' or 'End'.
    """
    return feature_manager.create_ref_plane_tangent_cylinder_keypoint(
        face_index, keypoint_type, parent_plane_index
    )


def create_ref_plane_tangent_surface_keypoint(
    face_index: int, keypoint_type: str = "End", parent_plane_index: int = 1
) -> dict:
    """Create a reference plane tangent to a curved surface at a keypoint.

    face_index: 0-based face index. keypoint_type: 'Start' or 'End'.
    """
    return feature_manager.create_ref_plane_tangent_surface_keypoint(
        face_index, keypoint_type, parent_plane_index
    )


# === Batch 4: Additional Surface Creation ===


def create_extruded_surface_from_to(from_plane_index: int, to_plane_index: int) -> dict:
    """Create an extruded surface between two reference planes (1-based indices)."""
    return feature_manager.create_extruded_surface_from_to(from_plane_index, to_plane_index)


def create_extruded_surface_by_keypoint(keypoint_type: str = "End") -> dict:
    """Create an extruded surface up to a keypoint extent."""
    return feature_manager.create_extruded_surface_by_keypoint(keypoint_type)


def create_extruded_surface_by_curves(distance: float, direction: str = "Normal") -> dict:
    """Create an extruded surface by curves (extrude along curve path)."""
    return feature_manager.create_extruded_surface_by_curves(distance, direction)


def create_revolved_surface_sync(angle: float = 360.0, want_end_caps: bool = False) -> dict:
    """Create a synchronous revolved construction surface."""
    return feature_manager.create_revolved_surface_sync(angle, want_end_caps)


def create_revolved_surface_by_keypoint(
    keypoint_type: str = "End", want_end_caps: bool = False
) -> dict:
    """Create a revolved surface up to a keypoint extent."""
    return feature_manager.create_revolved_surface_by_keypoint(keypoint_type, want_end_caps)


def create_lofted_surface_v2(want_end_caps: bool = False) -> dict:
    """Create a lofted surface using the extended Add2 method with OutputSurfaceType support."""
    return feature_manager.create_lofted_surface_v2(want_end_caps)


def create_swept_surface_ex(
    path_profile_index: int | None = None, want_end_caps: bool = False
) -> dict:
    """Create a swept surface using the extended AddEx method."""
    return feature_manager.create_swept_surface_ex(path_profile_index, want_end_caps)


def create_extruded_surface_full(
    distance: float,
    direction: str = "Normal",
    treatment_type: str = "None",
    draft_angle: float = 0.0,
) -> dict:
    """Create an extruded surface with full treatment parameters (crown, draft)."""
    return feature_manager.create_extruded_surface_full(
        distance, direction, treatment_type, draft_angle
    )


# === Batch 5: Protrusion & Cutout Variants ===


def create_extrude_through_next_v2(direction: str = "Normal") -> dict:
    """Create an extrusion through the next face (multi-profile collection API)."""
    return feature_manager.create_extrude_through_next_v2(direction)


def create_extrude_from_to_v2(from_plane_index: int, to_plane_index: int) -> dict:
    """Create an extrusion between two reference planes (multi-profile collection API)."""
    return feature_manager.create_extrude_from_to_v2(from_plane_index, to_plane_index)


def create_extrude_by_keypoint(direction: str = "Normal") -> dict:
    """Create an extrusion up to a keypoint extent."""
    return feature_manager.create_extrude_by_keypoint(direction)


def create_revolve_by_keypoint() -> dict:
    """Create a revolve up to a keypoint extent."""
    return feature_manager.create_revolve_by_keypoint()


def create_revolve_full(angle: float = 360.0, treatment_type: str = "None") -> dict:
    """Create a revolve with full treatment parameters (crown, draft)."""
    return feature_manager.create_revolve_full(angle, treatment_type)


def create_extruded_cutout_from_to_v2(from_plane_index: int, to_plane_index: int) -> dict:
    """Create an extruded cutout between two ref planes (multi-profile collection API)."""
    return feature_manager.create_extruded_cutout_from_to_v2(from_plane_index, to_plane_index)


def create_extruded_cutout_by_keypoint(direction: str = "Normal") -> dict:
    """Create an extruded cutout up to a keypoint extent."""
    return feature_manager.create_extruded_cutout_by_keypoint(direction)


def create_revolved_cutout_sync(angle: float = 360.0) -> dict:
    """Create a synchronous revolved cutout."""
    return feature_manager.create_revolved_cutout_sync(angle)


def create_revolved_cutout_by_keypoint() -> dict:
    """Create a revolved cutout up to a keypoint extent."""
    return feature_manager.create_revolved_cutout_by_keypoint()


def create_normal_cutout_from_to(from_plane_index: int, to_plane_index: int) -> dict:
    """Create a normal cutout between two reference planes."""
    return feature_manager.create_normal_cutout_from_to(from_plane_index, to_plane_index)


def create_normal_cutout_through_next(direction: str = "Normal") -> dict:
    """Create a normal cutout through the next face."""
    return feature_manager.create_normal_cutout_through_next(direction)


def create_normal_cutout_by_keypoint(direction: str = "Normal") -> dict:
    """Create a normal cutout up to a keypoint extent."""
    return feature_manager.create_normal_cutout_by_keypoint(direction)


def create_lofted_cutout_full(profile_indices: list[int] | None = None) -> dict:
    """Create a lofted cutout with full extent and guide curve parameters."""
    return feature_manager.create_lofted_cutout_full(profile_indices)


def create_swept_cutout_multi_body(path_profile_index: int | None = None) -> dict:
    """Create a swept cutout that spans multiple bodies."""
    return feature_manager.create_swept_cutout_multi_body(path_profile_index)


def create_helix_from_to(from_plane_index: int, to_plane_index: int, pitch: float) -> dict:
    """Create a helix protrusion between two reference planes."""
    return feature_manager.create_helix_from_to(from_plane_index, to_plane_index, pitch)


def create_helix_from_to_thin_wall(
    from_plane_index: int, to_plane_index: int, pitch: float, wall_thickness: float
) -> dict:
    """Create a thin-walled helix protrusion between two reference planes."""
    return feature_manager.create_helix_from_to_thin_wall(
        from_plane_index, to_plane_index, pitch, wall_thickness
    )


def create_helix_cutout_sync(
    pitch: float, height: float, revolutions: float = None, direction: str = "Right"
) -> dict:
    """Create a synchronous helical cutout."""
    return feature_manager.create_helix_cutout_sync(pitch, height, revolutions, direction)


def create_helix_cutout_from_to(from_plane_index: int, to_plane_index: int, pitch: float) -> dict:
    """Create a helical cutout between two reference planes."""
    return feature_manager.create_helix_cutout_from_to(from_plane_index, to_plane_index, pitch)


# === Batch 6: Rounds, Chamfers, Holes Extended ===


def create_round_blend(face_index1: int, face_index2: int, radius: float) -> dict:
    """Create a round blend between two faces."""
    return feature_manager.create_round_blend(face_index1, face_index2, radius)


def create_round_surface_blend(face_index1: int, face_index2: int, radius: float) -> dict:
    """Create a round surface blend between two faces."""
    return feature_manager.create_round_surface_blend(face_index1, face_index2, radius)


def create_hole_from_to(
    x: float, y: float, diameter: float, from_plane_index: int, to_plane_index: int
) -> dict:
    """Create a hole between two reference planes (circular cutout workaround)."""
    return feature_manager.create_hole_from_to(x, y, diameter, from_plane_index, to_plane_index)


def create_hole_through_next(
    x: float, y: float, diameter: float, direction: str = "Normal", plane_index: int = 1
) -> dict:
    """Create a hole through the next face (circular cutout workaround)."""
    return feature_manager.create_hole_through_next(x, y, diameter, direction, plane_index)


def create_hole_sync(
    x: float, y: float, diameter: float, depth: float, plane_index: int = 1
) -> dict:
    """Create a synchronous hole feature."""
    return feature_manager.create_hole_sync(x, y, diameter, depth, plane_index)


def create_hole_finite_ex(
    x: float,
    y: float,
    diameter: float,
    depth: float,
    direction: str = "Normal",
    plane_index: int = 1,
) -> dict:
    """Create a finite hole using the extended API."""
    return feature_manager.create_hole_finite_ex(x, y, diameter, depth, direction, plane_index)


def create_hole_from_to_ex(
    x: float, y: float, diameter: float, from_plane_index: int, to_plane_index: int
) -> dict:
    """Create a hole between two planes using the extended Holes API."""
    return feature_manager.create_hole_from_to_ex(x, y, diameter, from_plane_index, to_plane_index)


def create_hole_through_next_ex(
    x: float, y: float, diameter: float, direction: str = "Normal", plane_index: int = 1
) -> dict:
    """Create a hole through the next face using the extended Holes API."""
    return feature_manager.create_hole_through_next_ex(x, y, diameter, direction, plane_index)


def create_hole_through_all_ex(x: float, y: float, diameter: float, plane_index: int = 1) -> dict:
    """Create a hole through all material using the extended Holes API."""
    return feature_manager.create_hole_through_all_ex(x, y, diameter, plane_index)


def create_hole_sync_ex(
    x: float, y: float, diameter: float, depth: float, plane_index: int = 1
) -> dict:
    """Create a synchronous hole using the extended API."""
    return feature_manager.create_hole_sync_ex(x, y, diameter, depth, plane_index)


def create_hole_multi_body(
    x: float,
    y: float,
    diameter: float,
    depth: float,
    plane_index: int = 1,
    direction: str = "Normal",
) -> dict:
    """Create a hole that spans multiple bodies."""
    return feature_manager.create_hole_multi_body(x, y, diameter, depth, plane_index, direction)


def create_hole_sync_multi_body(
    x: float, y: float, diameter: float, depth: float, plane_index: int = 1
) -> dict:
    """Create a synchronous hole that spans multiple bodies."""
    return feature_manager.create_hole_sync_multi_body(x, y, diameter, depth, plane_index)


def create_blend_variable(radius1: float, radius2: float, face_index: int | None = None) -> dict:
    """Create a variable-radius blend feature."""
    return feature_manager.create_blend_variable(radius1, radius2, face_index)


def create_blend_surface(face_index1: int, face_index2: int) -> dict:
    """Create a surface blend between two faces."""
    return feature_manager.create_blend_surface(face_index1, face_index2)


# === Batch 7: Sheet Metal Completion ===


def create_flange_by_match_face(
    face_index: int,
    edge_index: int,
    flange_length: float,
    side: str = "Right",
    inside_radius: float = 0.001,
) -> dict:
    """Create a flange by matching an existing face edge."""
    return feature_manager.create_flange_by_match_face(
        face_index, edge_index, flange_length, side, inside_radius
    )


def create_flange_sync(
    face_index: int,
    edge_index: int,
    flange_length: float,
    inside_radius: float = 0.001,
) -> dict:
    """Create a synchronous flange feature on an edge."""
    return feature_manager.create_flange_sync(face_index, edge_index, flange_length, inside_radius)


def create_flange_by_face(
    face_index: int,
    edge_index: int,
    ref_face_index: int,
    flange_length: float,
    side: str = "Right",
    bend_radius: float = 0.001,
) -> dict:
    """Create a flange by face reference."""
    return feature_manager.create_flange_by_face(
        face_index, edge_index, ref_face_index, flange_length, side, bend_radius
    )


def create_flange_with_bend_calc(
    face_index: int,
    edge_index: int,
    flange_length: float,
    side: str = "Right",
    bend_deduction: float = 0.0,
) -> dict:
    """Create a flange with bend deduction/allowance calculation."""
    return feature_manager.create_flange_with_bend_calc(
        face_index, edge_index, flange_length, side, bend_deduction
    )


def create_flange_sync_with_bend_calc(
    face_index: int,
    edge_index: int,
    flange_length: float,
    bend_deduction: float = 0.0,
) -> dict:
    """Create a synchronous flange with bend deduction/allowance."""
    return feature_manager.create_flange_sync_with_bend_calc(
        face_index, edge_index, flange_length, bend_deduction
    )


def create_contour_flange_ex(
    thickness: float,
    bend_radius: float = 0.001,
    direction: str = "Normal",
) -> dict:
    """Create an extended contour flange from the active profile."""
    return feature_manager.create_contour_flange_ex(thickness, bend_radius, direction)


def create_contour_flange_sync(
    face_index: int,
    edge_index: int,
    thickness: float,
    bend_radius: float = 0.001,
    direction: str = "Normal",
) -> dict:
    """Create a synchronous contour flange with a reference edge."""
    return feature_manager.create_contour_flange_sync(
        face_index, edge_index, thickness, bend_radius, direction
    )


def create_contour_flange_sync_with_bend(
    face_index: int,
    edge_index: int,
    thickness: float,
    bend_radius: float = 0.001,
    direction: str = "Normal",
    bend_deduction: float = 0.0,
) -> dict:
    """Create a synchronous contour flange with bend deduction/allowance."""
    return feature_manager.create_contour_flange_sync_with_bend(
        face_index, edge_index, thickness, bend_radius, direction, bend_deduction
    )


def create_hem(
    face_index: int,
    edge_index: int,
    hem_width: float = 0.005,
    bend_radius: float = 0.001,
    hem_type: str = "Closed",
) -> dict:
    """Create a hem feature on a sheet metal edge."""
    return feature_manager.create_hem(face_index, edge_index, hem_width, bend_radius, hem_type)


def create_jog(
    jog_offset: float = 0.005,
    jog_angle: float = 90.0,
    direction: str = "Normal",
    moving_side: str = "Right",
) -> dict:
    """Create a jog feature on sheet metal from the active profile."""
    return feature_manager.create_jog(jog_offset, jog_angle, direction, moving_side)


def create_close_corner(
    face_index: int,
    edge_index: int,
    closure_type: str = "Close",
) -> dict:
    """Create a close corner feature on sheet metal."""
    return feature_manager.create_close_corner(face_index, edge_index, closure_type)


def create_multi_edge_flange(
    face_index: int,
    edge_indices: list[int],
    flange_length: float,
    side: str = "Right",
) -> dict:
    """Create a multi-edge flange on multiple edges simultaneously."""
    return feature_manager.create_multi_edge_flange(face_index, edge_indices, flange_length, side)


def create_bend_with_calc(
    bend_angle: float = 90.0,
    direction: str = "Normal",
    moving_side: str = "Right",
    bend_deduction: float = 0.0,
) -> dict:
    """Create a bend feature with bend deduction/allowance from the active profile."""
    return feature_manager.create_bend_with_calc(bend_angle, direction, moving_side, bend_deduction)


def convert_part_to_sheet_metal(thickness: float = 0.001) -> dict:
    """Convert the current part to a sheet metal document."""
    return feature_manager.convert_part_to_sheet_metal(thickness)


def create_dimple_ex(
    depth: float,
    direction: str = "Normal",
    punch_tool_diameter: float = 0.01,
) -> dict:
    """Create an extended dimple feature with punch tool diameter control."""
    return feature_manager.create_dimple_ex(depth, direction, punch_tool_diameter)


# === Batch 9: Part Feature Variants & Pattern Variants ===


def create_thread_ex(face_index: int, depth: float, pitch: float = 0.001) -> dict:
    """Create an extended thread on a cylindrical face with depth and pitch control."""
    return feature_manager.create_thread_ex(face_index, depth, pitch)


def create_slot_ex(width: float, depth: float, direction: str = "Normal") -> dict:
    """Create an extended slot feature with width and depth."""
    return feature_manager.create_slot_ex(width, depth, direction)


def create_slot_sync(width: float, depth: float) -> dict:
    """Create a synchronous slot feature."""
    return feature_manager.create_slot_sync(width, depth)


def create_drawn_cutout_ex(depth: float, direction: str = "Normal") -> dict:
    """Create an extended drawn cutout feature (sheet metal)."""
    return feature_manager.create_drawn_cutout_ex(depth, direction)


def create_louver_sync(depth: float) -> dict:
    """Create a synchronous louver feature (sheet metal)."""
    return feature_manager.create_louver_sync(depth)


def create_thicken_sync(thickness: float, direction: str = "Both") -> dict:
    """Create a synchronous thicken feature."""
    return feature_manager.create_thicken_sync(thickness, direction)


def create_mirror_sync_ex(feature_name: str, mirror_plane_index: int) -> dict:
    """Create a synchronous mirror copy using the extended AddSyncEx method."""
    return feature_manager.create_mirror_sync_ex(feature_name, mirror_plane_index)


def create_pattern_rectangular_ex(
    feature_name: str, x_count: int, y_count: int, x_spacing: float, y_spacing: float
) -> dict:
    """Create a rectangular pattern using the extended Ex API."""
    return feature_manager.create_pattern_rectangular_ex(
        feature_name, x_count, y_count, x_spacing, y_spacing
    )


def create_pattern_circular_ex(
    feature_name: str, count: int, angle: float, axis_face_index: int
) -> dict:
    """Create a circular pattern using the extended Ex API."""
    return feature_manager.create_pattern_circular_ex(feature_name, count, angle, axis_face_index)


def create_pattern_duplicate(feature_name: str) -> dict:
    """Create a duplicate pattern of a feature."""
    return feature_manager.create_pattern_duplicate(feature_name)


def create_pattern_by_fill(
    feature_name: str, fill_region_face_index: int, x_spacing: float, y_spacing: float
) -> dict:
    """Create a fill pattern of a feature within a face region."""
    return feature_manager.create_pattern_by_fill(
        feature_name, fill_region_face_index, x_spacing, y_spacing
    )


def create_pattern_by_table(
    feature_name: str, x_offsets: list[float], y_offsets: list[float]
) -> dict:
    """Create a table-driven pattern of a feature at specific X/Y offsets."""
    return feature_manager.create_pattern_by_table(feature_name, x_offsets, y_offsets)


# === Batch 10: Reference Planes v2, Sync Variants, Single-Profile,
#     MultiBody Cutouts, Full Treatment, Sheet Metal, Pattern Ex + Slots ===


# Group 1: Reference Planes (4 tools)


def create_ref_plane_normal_at_distance_v2(
    curve_edge_index: int,
    orientation_plane_index: int,
    distance: float,
    normal_side: int = 2,
) -> dict:
    """Create a ref plane normal to a curve at a distance (v2 with explicit edge/plane args)."""
    return feature_manager.create_ref_plane_normal_at_distance_v2(
        curve_edge_index, orientation_plane_index, distance, normal_side
    )


def create_ref_plane_normal_at_arc_ratio_v2(
    curve_edge_index: int,
    orientation_plane_index: int,
    ratio: float,
    normal_side: int = 2,
) -> dict:
    """Create a reference plane normal to a curve at an arc-length ratio (v2 with explicit args)."""
    return feature_manager.create_ref_plane_normal_at_arc_ratio_v2(
        curve_edge_index, orientation_plane_index, ratio, normal_side
    )


def create_ref_plane_normal_at_distance_along_v2(
    curve_edge_index: int,
    orientation_plane_index: int,
    distance_along: float,
    normal_side: int = 2,
) -> dict:
    """Create a reference plane normal to a curve at a distance along the curve (v2)."""
    return feature_manager.create_ref_plane_normal_at_distance_along_v2(
        curve_edge_index, orientation_plane_index, distance_along, normal_side
    )


def create_ref_plane_tangent_parallel(
    parent_plane_index: int, face_index: int, normal_side: int = 2
) -> dict:
    """Create a reference plane parallel to a parent plane and tangent to a face (v2)."""
    return feature_manager.create_ref_plane_tangent_parallel(
        parent_plane_index, face_index, normal_side
    )


# Group 2: Sync Variants (5 tools)


def create_revolve_by_keypoint_sync() -> dict:
    """Create a synchronous revolve up to a keypoint extent."""
    return feature_manager.create_revolve_by_keypoint_sync()


def create_helix_from_to_sync(from_plane_index: int, to_plane_index: int, pitch: float) -> dict:
    """Create a synchronous helix protrusion between two reference planes."""
    return feature_manager.create_helix_from_to_sync(from_plane_index, to_plane_index, pitch)


def create_helix_from_to_sync_thin_wall(
    from_plane_index: int, to_plane_index: int, pitch: float, wall_thickness: float
) -> dict:
    """Create a synchronous thin-walled helix between two reference planes."""
    return feature_manager.create_helix_from_to_sync_thin_wall(
        from_plane_index, to_plane_index, pitch, wall_thickness
    )


def create_helix_cutout_from_to_sync(
    from_plane_index: int, to_plane_index: int, pitch: float
) -> dict:
    """Create a synchronous helical cutout between two reference planes."""
    return feature_manager.create_helix_cutout_from_to_sync(from_plane_index, to_plane_index, pitch)


def create_pattern_by_table_sync(
    feature_name: str, x_offsets: list[float], y_offsets: list[float]
) -> dict:
    """Create a synchronous table-driven pattern of a feature."""
    return feature_manager.create_pattern_by_table_sync(feature_name, x_offsets, y_offsets)


# Group 3: Single-Profile Variants (3 tools)


def create_extrude_from_to_single(from_plane_index: int, to_plane_index: int) -> dict:
    """Create an extrusion from-to using single-profile API (AddFromTo)."""
    return feature_manager.create_extrude_from_to_single(from_plane_index, to_plane_index)


def create_extrude_through_next_single(direction: str = "Normal") -> dict:
    """Create an extrusion through next face using single-profile API (AddThroughNext)."""
    return feature_manager.create_extrude_through_next_single(direction)


def create_extruded_cutout_through_next_single(direction: str = "Normal") -> dict:
    """Create an extruded cutout through next face using single-profile API (AddThroughNext)."""
    return feature_manager.create_extruded_cutout_through_next_single(direction)


# Group 4: MultiBody Cutout Variants (4 tools)


def create_extruded_cutout_multi_body(distance: float, direction: str = "Normal") -> dict:
    """Create an extruded cutout that spans multiple bodies."""
    return feature_manager.create_extruded_cutout_multi_body(distance, direction)


def create_extruded_cutout_from_to_multi_body(from_plane_index: int, to_plane_index: int) -> dict:
    """Create an extruded cutout from-to that spans multiple bodies."""
    return feature_manager.create_extruded_cutout_from_to_multi_body(
        from_plane_index, to_plane_index
    )


def create_extruded_cutout_through_all_multi_body(direction: str = "Normal") -> dict:
    """Create an extruded cutout through-all that spans multiple bodies."""
    return feature_manager.create_extruded_cutout_through_all_multi_body(direction)


def create_revolved_cutout_multi_body(angle: float = 360.0) -> dict:
    """Create a revolved cutout that spans multiple bodies."""
    return feature_manager.create_revolved_cutout_multi_body(angle)


# Group 5: Full Treatment Variants (4 tools)


def create_revolved_cutout_full(angle: float = 360.0) -> dict:
    """Create a revolved cutout with full treatment parameters."""
    return feature_manager.create_revolved_cutout_full(angle)


def create_revolved_cutout_full_sync(angle: float = 360.0) -> dict:
    """Create a synchronous revolved cutout with full treatment parameters."""
    return feature_manager.create_revolved_cutout_full_sync(angle)


def create_revolved_surface_full(angle: float = 360.0, want_end_caps: bool = False) -> dict:
    """Create a revolved surface with full treatment parameters."""
    return feature_manager.create_revolved_surface_full(angle, want_end_caps)


def create_revolved_surface_full_sync(angle: float = 360.0, want_end_caps: bool = False) -> dict:
    """Create a synchronous revolved surface with full treatment parameters."""
    return feature_manager.create_revolved_surface_full_sync(angle, want_end_caps)


# Group 6: Sheet Metal (5 tools)


def create_flange_match_face_with_bend(
    face_index: int,
    edge_index: int,
    flange_length: float,
    side: str = "Right",
    inside_radius: float = 0.001,
    bend_deduction: float = 0.0,
) -> dict:
    """Create a flange by matching a face with bend deduction/allowance."""
    return feature_manager.create_flange_match_face_with_bend(
        face_index, edge_index, flange_length, side, inside_radius, bend_deduction
    )


def create_flange_by_face_with_bend(
    face_index: int,
    edge_index: int,
    ref_face_index: int,
    flange_length: float,
    side: str = "Right",
    bend_radius: float = 0.001,
    bend_deduction: float = 0.0,
) -> dict:
    """Create a flange by face reference with bend deduction/allowance."""
    return feature_manager.create_flange_by_face_with_bend(
        face_index, edge_index, ref_face_index, flange_length, side, bend_radius, bend_deduction
    )


def create_contour_flange_v3(
    thickness: float, bend_radius: float = 0.001, direction: str = "Normal"
) -> dict:
    """Create a contour flange using the Add3 method with extended parameters."""
    return feature_manager.create_contour_flange_v3(thickness, bend_radius, direction)


def create_contour_flange_sync_ex(
    face_index: int,
    edge_index: int,
    thickness: float,
    bend_radius: float = 0.001,
    direction: str = "Normal",
    bend_deduction: float = 0.0,
) -> dict:
    """Create a synchronous contour flange with extended parameters (AddSyncEx)."""
    return feature_manager.create_contour_flange_sync_ex(
        face_index, edge_index, thickness, bend_radius, direction, bend_deduction
    )


def create_bend(
    bend_angle: float = 90.0,
    direction: str = "Normal",
    moving_side: str = "Right",
) -> dict:
    """Create a bend feature on sheet metal from the active profile."""
    return feature_manager.create_bend(bend_angle, direction, moving_side)


# Group 7: Pattern Ex + Slots (4 tools)


def create_pattern_by_fill_ex(
    feature_name: str, fill_region_face_index: int, x_spacing: float, y_spacing: float
) -> dict:
    """Create a fill pattern using the extended AddByFillEx method."""
    return feature_manager.create_pattern_by_fill_ex(
        feature_name, fill_region_face_index, x_spacing, y_spacing
    )


def create_pattern_by_curve_ex(
    feature_name: str, curve_edge_index: int, count: int, spacing: float
) -> dict:
    """Create a curve-driven pattern using the extended AddByCurveEx method."""
    return feature_manager.create_pattern_by_curve_ex(
        feature_name, curve_edge_index, count, spacing
    )


def create_slot_multi_body(width: float, depth: float, direction: str = "Normal") -> dict:
    """Create a slot feature that spans multiple bodies."""
    return feature_manager.create_slot_multi_body(width, depth, direction)


def create_slot_sync_multi_body(width: float, depth: float, direction: str = "Normal") -> dict:
    """Create a synchronous slot feature that spans multiple bodies."""
    return feature_manager.create_slot_sync_multi_body(width, depth, direction)


# === Diagnostics ===


def diagnose_feature_tool(index: int) -> dict:
    """Diagnose properties and methods on a specific feature."""
    return diagnose_feature(feature_manager.doc_manager.get_active_document(), index)


# === Registration ===


def register(mcp):
    """Register feature tools with the MCP server."""
    # Extrusions
    mcp.tool()(create_extrude)
    mcp.tool()(create_extrude_infinite)
    mcp.tool()(create_extrude_through_next)
    mcp.tool()(create_extrude_from_to)
    mcp.tool()(create_extrude_thin_wall)
    # Revolves
    mcp.tool()(create_revolve)
    mcp.tool()(create_revolve_finite)
    mcp.tool()(create_revolve_sync)
    mcp.tool()(create_revolve_finite_sync)
    mcp.tool()(create_revolve_thin_wall)
    # Cutouts
    mcp.tool()(create_extruded_cutout)
    mcp.tool()(create_extruded_cutout_through_all)
    mcp.tool()(create_extruded_cutout_through_next)
    mcp.tool()(create_revolved_cutout)
    mcp.tool()(create_normal_cutout)
    mcp.tool()(create_normal_cutout_through_all)
    mcp.tool()(create_extruded_cutout_from_to)
    mcp.tool()(create_lofted_cutout)
    mcp.tool()(create_swept_cutout)
    mcp.tool()(create_drawn_cutout)
    # Loft / Sweep / Helix
    mcp.tool()(create_loft)
    mcp.tool()(create_loft_thin_wall)
    mcp.tool()(create_sweep)
    mcp.tool()(create_sweep_thin_wall)
    mcp.tool()(create_helix)
    mcp.tool()(create_helix_sync)
    mcp.tool()(create_helix_thin_wall)
    mcp.tool()(create_helix_sync_thin_wall)
    mcp.tool()(create_helix_cutout)
    # Surfaces
    mcp.tool()(create_extruded_surface)
    mcp.tool()(create_revolved_surface)
    mcp.tool()(create_lofted_surface)
    mcp.tool()(create_swept_surface)
    mcp.tool()(thicken_surface)
    # Primitives
    mcp.tool()(create_box_by_center)
    mcp.tool()(create_box_by_two_points)
    mcp.tool()(create_box_by_three_points)
    mcp.tool()(create_cylinder)
    mcp.tool()(create_sphere)
    # Primitive cutouts
    mcp.tool()(create_box_cutout)
    mcp.tool()(create_cylinder_cutout)
    mcp.tool()(create_sphere_cutout)
    # Holes
    mcp.tool()(create_hole)
    mcp.tool()(create_hole_through_all)
    mcp.tool()(create_delete_hole)
    mcp.tool()(create_delete_hole_by_face)
    # Rounds / Chamfers
    mcp.tool()(create_round)
    mcp.tool()(create_round_on_face)
    mcp.tool()(create_variable_round)
    mcp.tool()(create_chamfer)
    mcp.tool()(create_chamfer_on_face)
    mcp.tool()(create_chamfer_unequal)
    mcp.tool()(create_chamfer_angle)
    # Blends / Draft / Mirror
    mcp.tool()(create_blend)
    mcp.tool()(create_delete_blend)
    mcp.tool()(create_draft_angle)
    mcp.tool()(create_mirror)
    # Thin Wall
    mcp.tool()(create_thin_wall)
    mcp.tool()(create_thin_wall_with_open_faces)
    # Part Features
    mcp.tool()(create_rib)
    mcp.tool()(create_lip)
    mcp.tool()(create_slot)
    mcp.tool()(create_split)
    mcp.tool()(create_thread)
    mcp.tool()(create_emboss)
    mcp.tool()(create_etch)
    mcp.tool()(create_flange)
    mcp.tool()(create_dimple)
    mcp.tool()(create_bead)
    mcp.tool()(create_louver)
    mcp.tool()(create_gusset)
    # Patterns
    mcp.tool()(create_pattern_rectangular)
    mcp.tool()(create_pattern_circular)
    # Face Operations
    mcp.tool()(delete_faces)
    mcp.tool()(create_face_rotate_by_points)
    mcp.tool()(create_face_rotate_by_edge)
    # Reference Planes
    mcp.tool()(create_ref_plane_by_offset)
    mcp.tool()(create_ref_plane_by_angle)
    mcp.tool()(create_ref_plane_by_3_points)
    mcp.tool()(create_ref_plane_midplane)
    mcp.tool()(create_ref_plane_normal_to_curve)
    mcp.tool()(create_ref_plane_normal_at_distance)
    mcp.tool()(create_ref_plane_normal_at_arc_ratio)
    mcp.tool()(create_ref_plane_normal_at_distance_along)
    mcp.tool()(create_ref_plane_parallel_by_tangent)
    # Batch 4: Additional Ref Plane Variants
    mcp.tool()(create_ref_plane_normal_at_keypoint)
    mcp.tool()(create_ref_plane_tangent_cylinder_angle)
    mcp.tool()(create_ref_plane_tangent_cylinder_keypoint)
    mcp.tool()(create_ref_plane_tangent_surface_keypoint)
    # Batch 4: Additional Surface Creation
    mcp.tool()(create_extruded_surface_from_to)
    mcp.tool()(create_extruded_surface_by_keypoint)
    mcp.tool()(create_extruded_surface_by_curves)
    mcp.tool()(create_revolved_surface_sync)
    mcp.tool()(create_revolved_surface_by_keypoint)
    mcp.tool()(create_lofted_surface_v2)
    mcp.tool()(create_swept_surface_ex)
    mcp.tool()(create_extruded_surface_full)
    # Body Operations
    mcp.tool()(add_body)
    mcp.tool()(add_body_by_mesh)
    mcp.tool()(add_body_feature)
    mcp.tool()(add_by_construction)
    mcp.tool()(add_body_by_tag)
    # Simplification
    mcp.tool()(auto_simplify)
    mcp.tool()(simplify_enclosure)
    mcp.tool()(simplify_duplicate)
    mcp.tool()(local_simplify_enclosure)
    # Feature Management
    mcp.tool()(list_features)
    mcp.tool()(get_feature_info)
    mcp.tool()(delete_feature)
    mcp.tool()(feature_suppress)
    mcp.tool()(feature_unsuppress)
    mcp.tool()(feature_reorder)
    mcp.tool()(feature_rename)
    mcp.tool()(convert_feature_type)
    # Sheet Metal
    mcp.tool()(create_base_flange)
    mcp.tool()(create_base_tab)
    mcp.tool()(create_lofted_flange)
    mcp.tool()(create_web_network)
    mcp.tool()(create_base_contour_flange_advanced)
    mcp.tool()(create_base_tab_multi_profile)
    mcp.tool()(create_lofted_flange_advanced)
    mcp.tool()(create_lofted_flange_ex)
    # Symmetric Extrusion
    mcp.tool()(create_extrude_symmetric)
    # Unequal Chamfer on Face
    mcp.tool()(create_chamfer_unequal_on_face)
    # Batch 5: Protrusion & Cutout Variants
    mcp.tool()(create_extrude_through_next_v2)
    mcp.tool()(create_extrude_from_to_v2)
    mcp.tool()(create_extrude_by_keypoint)
    mcp.tool()(create_revolve_by_keypoint)
    mcp.tool()(create_revolve_full)
    mcp.tool()(create_extruded_cutout_from_to_v2)
    mcp.tool()(create_extruded_cutout_by_keypoint)
    mcp.tool()(create_revolved_cutout_sync)
    mcp.tool()(create_revolved_cutout_by_keypoint)
    mcp.tool()(create_normal_cutout_from_to)
    mcp.tool()(create_normal_cutout_through_next)
    mcp.tool()(create_normal_cutout_by_keypoint)
    mcp.tool()(create_lofted_cutout_full)
    mcp.tool()(create_swept_cutout_multi_body)
    mcp.tool()(create_helix_from_to)
    mcp.tool()(create_helix_from_to_thin_wall)
    mcp.tool()(create_helix_cutout_sync)
    mcp.tool()(create_helix_cutout_from_to)
    # Batch 6: Rounds, Chamfers, Holes Extended
    mcp.tool()(create_round_blend)
    mcp.tool()(create_round_surface_blend)
    mcp.tool()(create_hole_from_to)
    mcp.tool()(create_hole_through_next)
    mcp.tool()(create_hole_sync)
    mcp.tool()(create_hole_finite_ex)
    mcp.tool()(create_hole_from_to_ex)
    mcp.tool()(create_hole_through_next_ex)
    mcp.tool()(create_hole_through_all_ex)
    mcp.tool()(create_hole_sync_ex)
    mcp.tool()(create_hole_multi_body)
    mcp.tool()(create_hole_sync_multi_body)
    mcp.tool()(create_blend_variable)
    mcp.tool()(create_blend_surface)
    # Batch 7: Sheet Metal Completion
    mcp.tool()(create_flange_by_match_face)
    mcp.tool()(create_flange_sync)
    mcp.tool()(create_flange_by_face)
    mcp.tool()(create_flange_with_bend_calc)
    mcp.tool()(create_flange_sync_with_bend_calc)
    mcp.tool()(create_contour_flange_ex)
    mcp.tool()(create_contour_flange_sync)
    mcp.tool()(create_contour_flange_sync_with_bend)
    mcp.tool()(create_hem)
    mcp.tool()(create_jog)
    mcp.tool()(create_close_corner)
    mcp.tool()(create_multi_edge_flange)
    mcp.tool()(create_bend_with_calc)
    mcp.tool()(convert_part_to_sheet_metal)
    mcp.tool()(create_dimple_ex)
    # Batch 9: Part Feature Variants & Pattern Variants
    mcp.tool()(create_thread_ex)
    mcp.tool()(create_slot_ex)
    mcp.tool()(create_slot_sync)
    mcp.tool()(create_drawn_cutout_ex)
    mcp.tool()(create_louver_sync)
    mcp.tool()(create_thicken_sync)
    mcp.tool()(create_mirror_sync_ex)
    mcp.tool()(create_pattern_rectangular_ex)
    mcp.tool()(create_pattern_circular_ex)
    mcp.tool()(create_pattern_duplicate)
    mcp.tool()(create_pattern_by_fill)
    mcp.tool()(create_pattern_by_table)
    # Batch 10: Reference Planes v2, Sync Variants, Single-Profile,
    #   MultiBody Cutouts, Full Treatment, Sheet Metal, Pattern Ex + Slots
    # Group 1: Reference Planes (4)
    mcp.tool()(create_ref_plane_normal_at_distance_v2)
    mcp.tool()(create_ref_plane_normal_at_arc_ratio_v2)
    mcp.tool()(create_ref_plane_normal_at_distance_along_v2)
    mcp.tool()(create_ref_plane_tangent_parallel)
    # Group 2: Sync Variants (5)
    mcp.tool()(create_revolve_by_keypoint_sync)
    mcp.tool()(create_helix_from_to_sync)
    mcp.tool()(create_helix_from_to_sync_thin_wall)
    mcp.tool()(create_helix_cutout_from_to_sync)
    mcp.tool()(create_pattern_by_table_sync)
    # Group 3: Single-Profile Variants (3)
    mcp.tool()(create_extrude_from_to_single)
    mcp.tool()(create_extrude_through_next_single)
    mcp.tool()(create_extruded_cutout_through_next_single)
    # Group 4: MultiBody Cutout Variants (4)
    mcp.tool()(create_extruded_cutout_multi_body)
    mcp.tool()(create_extruded_cutout_from_to_multi_body)
    mcp.tool()(create_extruded_cutout_through_all_multi_body)
    mcp.tool()(create_revolved_cutout_multi_body)
    # Group 5: Full Treatment Variants (4)
    mcp.tool()(create_revolved_cutout_full)
    mcp.tool()(create_revolved_cutout_full_sync)
    mcp.tool()(create_revolved_surface_full)
    mcp.tool()(create_revolved_surface_full_sync)
    # Group 6: Sheet Metal (5)
    mcp.tool()(create_flange_match_face_with_bend)
    mcp.tool()(create_flange_by_face_with_bend)
    mcp.tool()(create_contour_flange_v3)
    mcp.tool()(create_contour_flange_sync_ex)
    mcp.tool()(create_bend)
    # Group 7: Pattern Ex + Slots (4)
    mcp.tool()(create_pattern_by_fill_ex)
    mcp.tool()(create_pattern_by_curve_ex)
    mcp.tool()(create_slot_multi_body)
    mcp.tool()(create_slot_sync_multi_body)
    # Diagnostics
    mcp.tool(name="diagnose_feature")(diagnose_feature_tool)
