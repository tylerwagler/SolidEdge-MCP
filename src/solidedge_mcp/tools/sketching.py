"""Sketching tools for Solid Edge MCP."""

from solidedge_mcp.managers import sketch_manager


def create_sketch(plane: str = "Top") -> dict:
    """Create a new 2D sketch on a reference plane.

    Plane: 'Top', 'Front', 'Right', 'XY', 'XZ', 'YZ'.
    """
    return sketch_manager.create_sketch(plane)


def draw_line(x1: float, y1: float, x2: float, y2: float) -> dict:
    """Draw a line in the active sketch. Coordinates in meters."""
    return sketch_manager.draw_line(x1, y1, x2, y2)


def draw_circle(center_x: float, center_y: float, radius: float) -> dict:
    """Draw a circle in the active sketch. Coordinates and radius in meters."""
    return sketch_manager.draw_circle(center_x, center_y, radius)


def draw_rectangle(x1: float, y1: float, x2: float, y2: float) -> dict:
    """Draw a rectangle in the active sketch defined by two diagonal corners (meters)."""
    return sketch_manager.draw_rectangle(x1, y1, x2, y2)


def draw_arc(
    center_x: float, center_y: float, radius: float, start_angle: float, end_angle: float
) -> dict:
    """Draw an arc by center, radius, and angles (degrees)."""
    return sketch_manager.draw_arc(center_x, center_y, radius, start_angle, end_angle)


def draw_polygon(center_x: float, center_y: float, radius: float, sides: int) -> dict:
    """Draw a regular polygon."""
    return sketch_manager.draw_polygon(center_x, center_y, radius, sides)


def draw_ellipse(
    center_x: float, center_y: float, major_radius: float, minor_radius: float, angle: float = 0.0
) -> dict:
    """Draw an ellipse. Angle in degrees."""
    return sketch_manager.draw_ellipse(center_x, center_y, major_radius, minor_radius, angle)


def draw_spline(points: list) -> dict:
    """Draw a B-spline curve through a list of [x, y] points."""
    return sketch_manager.draw_spline(points)


def draw_arc_by_3_points(
    start_x: float, start_y: float, center_x: float, center_y: float, end_x: float, end_y: float
) -> dict:
    """Draw an arc defined by start, center, and end points."""
    return sketch_manager.draw_arc_by_3_points(start_x, start_y, center_x, center_y, end_x, end_y)


def draw_circle_by_2_points(x1: float, y1: float, x2: float, y2: float) -> dict:
    """Draw a circle defined by two diametrically opposite points."""
    return sketch_manager.draw_circle_by_2_points(x1, y1, x2, y2)


def draw_circle_by_3_points(
    x1: float, y1: float, x2: float, y2: float, x3: float, y3: float
) -> dict:
    """Draw a circle through three points."""
    return sketch_manager.draw_circle_by_3_points(x1, y1, x2, y2, x3, y3)


def mirror_spline(
    axis_x1: float, axis_y1: float, axis_x2: float, axis_y2: float, copy: bool = True
) -> dict:
    """Mirror B-spline curves across a line defined by two points."""
    return sketch_manager.mirror_spline(axis_x1, axis_y1, axis_x2, axis_y2, copy)


def set_profile_visibility(visible: bool = False) -> dict:
    """Show or hide the active sketch profile."""
    return sketch_manager.hide_profile(visible)


def draw_point(x: float, y: float) -> dict:
    """Draw a construction point in the active sketch."""
    return sketch_manager.draw_point(x, y)


def add_constraint(constraint_type: str, elements: list) -> dict:
    """Add a geometric constraint (Horizontal, Vertical, Parallel, etc.) to sketch elements."""
    return sketch_manager.add_constraint(constraint_type, elements)


def add_keypoint_constraint(
    element1_type: str,
    element1_index: int,
    keypoint1: int,
    element2_type: str,
    element2_index: int,
    keypoint2: int,
) -> dict:
    """Add a keypoint constraint connecting two sketch elements at specific points."""
    return sketch_manager.add_keypoint_constraint(
        element1_type, element1_index, keypoint1, element2_type, element2_index, keypoint2
    )


def set_axis_of_revolution(x1: float, y1: float, x2: float, y2: float) -> dict:
    """Draw an axis of revolution for revolve operations."""
    return sketch_manager.set_axis_of_revolution(x1, y1, x2, y2)


def create_sketch_on_plane(plane_index: int) -> dict:
    """Create a new 2D sketch on a reference plane by index (1-based)."""
    return sketch_manager.create_sketch_on_plane_index(plane_index)


def close_sketch() -> dict:
    """Close/finish the active sketch."""
    return sketch_manager.close_sketch()


def get_sketch_info() -> dict:
    """Get information about geometry counts in the active sketch."""
    return sketch_manager.get_sketch_info()


def sketch_fillet(radius: float) -> dict:
    """Add fillet (round) to sketch corners."""
    return sketch_manager.sketch_fillet(radius)


def sketch_chamfer(distance: float) -> dict:
    """Add chamfer to sketch corners."""
    return sketch_manager.sketch_chamfer(distance)


def sketch_offset(distance: float) -> dict:
    """Offset the entire active sketch profile by a distance."""
    return sketch_manager.sketch_offset(distance)


def sketch_rotate(center_x: float, center_y: float, angle_degrees: float) -> dict:
    """Rotate the entire active sketch profile around a center point."""
    return sketch_manager.sketch_rotate(center_x, center_y, angle_degrees)


def sketch_scale(center_x: float, center_y: float, scale_factor: float) -> dict:
    """Scale the entire active sketch profile relative to a center point."""
    return sketch_manager.sketch_scale(center_x, center_y, scale_factor)


def sketch_mirror(axis: str = "X") -> dict:
    """Mirror sketch geometry about an axis ('X' or 'Y')."""
    return sketch_manager.sketch_mirror(axis)


def get_sketch_matrix() -> dict:
    """Get the sketch coordinate system matrix (2D-to-3D transformation)."""
    return sketch_manager.get_sketch_matrix()


def clean_sketch_geometry(
    clean_points: bool = True,
    clean_splines: bool = True,
    clean_identical: bool = True,
    clean_small: bool = True,
    small_tolerance: float = 0.0001,
) -> dict:
    """Clean up duplicate, small, or invalid geometry in the active sketch."""
    return sketch_manager.clean_sketch_geometry(
        clean_points, clean_splines, clean_identical, clean_small, small_tolerance
    )


def draw_construction_line(x1: float, y1: float, x2: float, y2: float) -> dict:
    """Draw a construction line in the active sketch."""
    return sketch_manager.draw_construction_line(x1, y1, x2, y2)


def project_edge(face_index: int, edge_index: int) -> dict:
    """Project a 3D body edge into the active sketch."""
    return sketch_manager.project_edge(face_index, edge_index)


def include_edge(face_index: int, edge_index: int) -> dict:
    """Include a 3D body edge in the active sketch (associative)."""
    return sketch_manager.include_edge(face_index, edge_index)


def project_ref_plane(plane_index: int) -> dict:
    """Project a reference plane into the active sketch as construction geometry."""
    return sketch_manager.project_ref_plane(plane_index)


def offset_sketch_2d(offset_side_x: float, offset_side_y: float, offset_distance: float) -> dict:
    """Offset the active sketch profile in 2D. Side params control offset direction."""
    return sketch_manager.offset_sketch_2d(offset_side_x, offset_side_y, offset_distance)


def get_sketch_constraints() -> dict:
    """Get information about constraints in the active sketch."""
    return sketch_manager.get_sketch_constraints()


def project_silhouette_edges() -> dict:
    """Project silhouette edges of the body onto the active sketch."""
    return sketch_manager.project_silhouette_edges()


def include_region_faces(face_indices: list[int]) -> dict:
    """Include faces as regions in the active sketch."""
    return sketch_manager.include_region_faces(face_indices)


def chain_locate(x: float, y: float, tolerance: float = 0.001) -> dict:
    """Find a chain of connected sketch elements at a location."""
    return sketch_manager.chain_locate(x, y, tolerance)


def convert_to_curve() -> dict:
    """Convert sketch geometry to a curve."""
    return sketch_manager.convert_to_curve()


def sketch_paste() -> dict:
    """Paste clipboard content into the active sketch."""
    return sketch_manager.sketch_paste()


def get_ordered_geometry() -> dict:
    """Get the ordered geometry elements from the active sketch."""
    return sketch_manager.get_ordered_geometry()


def register(mcp):
    """Register sketching tools with the MCP server."""
    mcp.tool()(create_sketch)
    mcp.tool()(draw_line)
    mcp.tool()(draw_circle)
    mcp.tool()(draw_rectangle)
    mcp.tool()(draw_arc)
    mcp.tool()(draw_polygon)
    mcp.tool()(draw_ellipse)
    mcp.tool()(draw_spline)
    mcp.tool()(draw_arc_by_3_points)
    mcp.tool()(draw_circle_by_2_points)
    mcp.tool()(draw_circle_by_3_points)
    mcp.tool()(mirror_spline)
    mcp.tool()(set_profile_visibility)
    mcp.tool()(draw_point)
    mcp.tool()(add_constraint)
    mcp.tool()(add_keypoint_constraint)
    mcp.tool()(set_axis_of_revolution)
    mcp.tool()(create_sketch_on_plane)
    mcp.tool()(close_sketch)
    mcp.tool()(get_sketch_info)
    mcp.tool()(sketch_fillet)
    mcp.tool()(sketch_chamfer)
    mcp.tool()(sketch_offset)
    mcp.tool()(sketch_rotate)
    mcp.tool()(sketch_scale)
    mcp.tool()(sketch_mirror)
    mcp.tool()(draw_construction_line)
    mcp.tool()(project_edge)
    mcp.tool()(include_edge)
    mcp.tool()(project_ref_plane)
    mcp.tool()(offset_sketch_2d)
    mcp.tool()(get_sketch_constraints)
    mcp.tool()(get_sketch_matrix)
    mcp.tool()(clean_sketch_geometry)
    # Batch 9: Sketch Extensions
    mcp.tool()(project_silhouette_edges)
    mcp.tool()(include_region_faces)
    mcp.tool()(chain_locate)
    mcp.tool()(convert_to_curve)
    # Batch 10: Sketch Paste & Ordered Geometry
    mcp.tool()(sketch_paste)
    mcp.tool()(get_ordered_geometry)
