"""Feature modeling tools for Solid Edge MCP."""

from typing import Optional, List
from solidedge_mcp.managers import feature_manager, diagnose_feature


# === Extrusions ===

def create_extrude(distance: float, direction: str = "Normal") -> dict:
    """Create an extruded protrusion from the active sketch profile."""
    return feature_manager.create_extrude(distance, direction)

def create_extrude_infinite(direction: str = "Normal") -> dict:
    """Create an infinite extrusion (extends through all)."""
    return feature_manager.create_extrude_infinite(direction)

def create_extrude_thin_wall(distance: float, wall_thickness: float, direction: str = "Normal") -> dict:
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

def create_lofted_cutout(profile_indices: Optional[List[int]] = None) -> dict:
    """Create a lofted cutout between multiple profiles."""
    return feature_manager.create_lofted_cutout(profile_indices)

def create_swept_cutout() -> dict:
    """Create a swept cutout along a path."""
    return feature_manager.create_swept_cutout()

def create_drawn_cutout(depth: float, direction: str = "Normal") -> dict:
    """Create a drawn cutout (sheet metal)."""
    return feature_manager.create_drawn_cutout(depth, direction)


# === Loft / Sweep / Helix ===

def create_loft(profile_indices: Optional[list] = None) -> dict:
    """Create a loft feature between multiple profiles."""
    return feature_manager.create_loft(profile_indices)

def create_loft_thin_wall(wall_thickness: float, profile_indices: Optional[list] = None) -> dict:
    """Create a thin-walled loft feature."""
    return feature_manager.create_loft_thin_wall(wall_thickness, profile_indices)

def create_sweep(path_profile_index: Optional[int] = None) -> dict:
    """Create a sweep feature along a path."""
    return feature_manager.create_sweep(path_profile_index)

def create_sweep_thin_wall(wall_thickness: float, path_profile_index: Optional[int] = None) -> dict:
    """Create a thin-walled sweep feature."""
    return feature_manager.create_sweep_thin_wall(wall_thickness, path_profile_index)

def create_helix(pitch: float, height: float, revolutions: Optional[float] = None, direction: str = "Right") -> dict:
    """Create a helical feature (springs, threads)."""
    return feature_manager.create_helix(pitch, height, revolutions, direction)

def create_helix_sync(pitch: float, height: float, revolutions: Optional[float] = None) -> dict:
    """Create a synchronous helical feature."""
    return feature_manager.create_helix_sync(pitch, height, revolutions)

def create_helix_thin_wall(pitch: float, height: float, wall_thickness: float, revolutions: float = None) -> dict:
    """Create a thin-walled helix feature."""
    return feature_manager.create_helix_thin_wall(pitch, height, wall_thickness, revolutions)

def create_helix_sync_thin_wall(pitch: float, height: float, wall_thickness: float, revolutions: float = None) -> dict:
    """Create a synchronous thin-walled helix feature."""
    return feature_manager.create_helix_sync_thin_wall(pitch, height, wall_thickness, revolutions)

def create_helix_cutout(pitch: float, height: float, revolutions: float = None, direction: str = "Right") -> dict:
    """Create a helical cutout."""
    return feature_manager.create_helix_cutout(pitch, height, revolutions, direction)


# === Surfaces ===

def create_extruded_surface(distance: float, direction: str = "Normal", end_caps: bool = True) -> dict:
    """Create an extruded surface."""
    return feature_manager.create_extruded_surface(distance, direction, end_caps)

def create_revolved_surface(angle: float = 360, want_end_caps: bool = False) -> dict:
    """Create a revolved surface."""
    return feature_manager.create_revolved_surface(angle, want_end_caps)

def create_lofted_surface(want_end_caps: bool = False) -> dict:
    """Create a lofted surface."""
    return feature_manager.create_lofted_surface(want_end_caps)

def create_swept_surface(path_profile_index: Optional[int] = None, want_end_caps: bool = False) -> dict:
    """Create a swept surface."""
    return feature_manager.create_swept_surface(path_profile_index, want_end_caps)

def thicken_surface(thickness: float, direction: str = "Both") -> dict:
    """Thicken a surface to create a solid."""
    return feature_manager.thicken_surface(thickness, direction)


# === Primitives ===

def create_box_by_center(x: float, y: float, z: float, length: float, width: float, height: float) -> dict:
    """Create a box primitive by center point and dimensions."""
    return feature_manager.create_box_by_center(x, y, z, length, width, height)

def create_box_by_two_points(x1: float, y1: float, z1: float, x2: float, y2: float, z2: float) -> dict:
    """Create a box primitive by two opposite corners."""
    return feature_manager.create_box_by_two_points(x1, y1, z1, x2, y2, z2)

def create_box_by_three_points(x1: float, y1: float, z1: float, x2: float, y2: float, z2: float, x3: float, y3: float, z3: float) -> dict:
    """Create a box primitive by three points."""
    return feature_manager.create_box_by_three_points(x1, y1, z1, x2, y2, z2, x3, y3, z3)

def create_cylinder(base_center_x: float, base_center_y: float, base_center_z: float, radius: float, height: float) -> dict:
    """Create a cylinder primitive."""
    return feature_manager.create_cylinder(base_center_x, base_center_y, base_center_z, radius, height)

def create_sphere(center_x: float, center_y: float, center_z: float, radius: float) -> dict:
    """Create a sphere primitive."""
    return feature_manager.create_sphere(center_x, center_y, center_z, radius)


# === Primitive Cutouts ===

def create_box_cutout(x1: float, y1: float, z1: float, x2: float, y2: float, z2: float) -> dict:
    """Create a box-shaped cutout by two opposite corners."""
    return feature_manager.create_box_cutout_by_two_points(x1, y1, z1, x2, y2, z2)

def create_cylinder_cutout(center_x: float, center_y: float, center_z: float, radius: float, height: float) -> dict:
    """Create a cylindrical cutout."""
    return feature_manager.create_cylinder_cutout(center_x, center_y, center_z, radius, height)

def create_sphere_cutout(center_x: float, center_y: float, center_z: float, radius: float) -> dict:
    """Create a spherical cutout."""
    return feature_manager.create_sphere_cutout(center_x, center_y, center_z, radius)


# === Holes ===

def create_hole(x: float, y: float, diameter: float, depth: float = 0.0, hole_type: str = "Simple", plane_index: int = 1, direction: str = "Normal") -> dict:
    """Create a hole at coordinates (meters)."""
    return feature_manager.create_hole(x, y, diameter, depth, direction)

def create_hole_through_all(x: float, y: float, diameter: float, plane_index: int = 1, direction: str = "Normal") -> dict:
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

def create_variable_round(radii: list, face_index: Optional[int] = None) -> dict:
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

def create_blend(radius: float, face_index: Optional[int] = None) -> dict:
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

def create_thin_wall_with_open_faces(thickness: float, open_face_indices: List[int]) -> dict:
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

def create_emboss(face_indices: list, clearance: float = 0.001, thickness: float = 0.0, thicken: bool = False, default_side: bool = True) -> dict:
    """Create an emboss feature."""
    return feature_manager.create_emboss(face_indices, clearance, thickness, thicken, default_side)

def create_etch() -> dict:
    """Create an etch feature from the active profile."""
    return feature_manager.create_etch()

def create_flange(face_index: int, edge_index: int, flange_length: float, side: str = "Right", inside_radius: float = None, bend_angle: float = None) -> dict:
    """Create a flange on an edge."""
    return feature_manager.create_flange(face_index, edge_index, flange_length, side, inside_radius, bend_angle)

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

def create_pattern_rectangular(feature_index: int, x_count: int, y_count: int, x_gap: float, y_gap: float) -> dict:
    """Create a rectangular pattern of a feature."""
    return feature_manager.create_pattern_rectangular(feature_index, x_count, y_count, x_gap, y_gap)

def create_pattern_circular(feature_index: int, count: int, angle: float, radius: float) -> dict:
    """Create a circular pattern of a feature."""
    return feature_manager.create_pattern_circular(feature_index, count, angle, radius)


# === Face Operations ===

def delete_faces(face_indices: List[int]) -> dict:
    """Delete specified faces from the body."""
    return feature_manager.delete_faces(face_indices)

def create_face_rotate_by_points(face_index: int, vertex1_index: int, vertex2_index: int, angle: float) -> dict:
    """Rotate a face around an axis defined by two vertices."""
    return feature_manager.create_face_rotate_by_points(face_index, vertex1_index, vertex2_index, angle)

def create_face_rotate_by_edge(face_index: int, edge_index: int, angle: float) -> dict:
    """Rotate a face around an edge."""
    return feature_manager.create_face_rotate_by_edge(face_index, edge_index, angle)


# === Reference Planes ===

def create_ref_plane_by_offset(parent_plane_index: int, distance: float, normal_side: str = "Normal") -> dict:
    """Create a reference plane offset from an existing plane."""
    return feature_manager.create_ref_plane_by_offset(parent_plane_index, distance, normal_side)

def create_ref_plane_by_angle(parent_plane_index: int, angle: float, normal_side: str = "Normal") -> dict:
    """Create a reference plane at an angle to an existing plane."""
    return feature_manager.create_ref_plane_by_angle(parent_plane_index, angle, normal_side)

def create_ref_plane_by_3_points(x1: float, y1: float, z1: float, x2: float, y2: float, z2: float, x3: float, y3: float, z3: float) -> dict:
    """Create a reference plane defined by three points."""
    return feature_manager.create_ref_plane_by_3_points(x1, y1, z1, x2, y2, z2, x3, y3, z3)

def create_ref_plane_midplane(plane1_index: int, plane2_index: int) -> dict:
    """Create a midplane between two existing planes."""
    return feature_manager.create_ref_plane_midplane(plane1_index, plane2_index)

def create_ref_plane_normal_to_curve(curve_end: str = "End", pivot_plane_index: int = 2) -> dict:
    """Create a reference plane normal to a curve endpoint."""
    return feature_manager.create_ref_plane_normal_to_curve(curve_end, pivot_plane_index)


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

def create_base_flange(width: float, thickness: float, bend_radius: Optional[float] = None) -> dict:
    """Create a base contour flange (sheet metal)."""
    return feature_manager.create_base_flange(width, thickness, bend_radius)

def create_base_tab(thickness: float, width: Optional[float] = None) -> dict:
    """Create a base tab (sheet metal)."""
    return feature_manager.create_base_tab(thickness, width)

def create_lofted_flange(thickness: float) -> dict:
    """Create a lofted flange (sheet metal)."""
    return feature_manager.create_lofted_flange(thickness)

def create_web_network() -> dict:
    """Create a web network (sheet metal)."""
    return feature_manager.create_web_network()

def create_base_contour_flange_advanced(thickness: float, bend_radius: float, relief_type: str = "Default") -> dict:
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
    # Diagnostics
    mcp.tool(name="diagnose_feature")(diagnose_feature_tool)
