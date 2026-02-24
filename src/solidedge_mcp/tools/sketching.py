"""Sketching tools for Solid Edge MCP."""

from typing import Any

from solidedge_mcp.backends.validation import validate_numerics
from solidedge_mcp.managers import sketch_manager

# === Composite: manage_sketch ===


def manage_sketch(
    action: str = "create",
    plane: str = "Top",
    plane_index: int = 1,
    x1: float = 0.0,
    y1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    visible: bool = False,
) -> dict[str, Any]:
    """Create, close, or configure a 2D sketch.

    action: 'create' | 'close' | 'create_on_plane'
      | 'set_axis' | 'set_visibility' | 'get_geometry'

    Named planes: 'Top','Front','Right','XY','XZ','YZ'.
    Coordinates in meters.
    """
    err = validate_numerics(x1=x1, y1=y1, x2=x2, y2=y2)
    if err:
        return err
    match action:
        case "create":
            return sketch_manager.create_sketch(plane)
        case "close":
            return sketch_manager.close_sketch()
        case "create_on_plane":
            return sketch_manager.create_sketch_on_plane_index(plane_index)
        case "set_axis":
            return sketch_manager.set_axis_of_revolution(x1, y1, x2, y2)
        case "set_visibility":
            return sketch_manager.hide_profile(visible)
        case "get_geometry":
            return sketch_manager.get_ordered_geometry()
        case _:
            return {"error": f"Unknown action: {action}"}


# === Composite: draw ===


def draw(
    shape: str = "line",
    x1: float = 0.0,
    y1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    x3: float = 0.0,
    y3: float = 0.0,
    center_x: float = 0.0,
    center_y: float = 0.0,
    radius: float = 0.0,
    major_radius: float = 0.0,
    minor_radius: float = 0.0,
    start_angle: float = 0.0,
    end_angle: float = 360.0,
    angle: float = 0.0,
    sides: int = 6,
    points: list[list[float]] | None = None,
    x: float = 0.0,
    y: float = 0.0,
) -> dict[str, Any]:
    """Draw geometry in the active sketch.

    shape: 'line' | 'circle' | 'rectangle' | 'arc' | 'polygon'
           | 'ellipse' | 'spline' | 'arc_3pt' | 'circle_2pt'
           | 'circle_3pt' | 'point' | 'construction_line'

    Coordinates in meters. Angles in degrees.
    """
    err = validate_numerics(
        x1=x1, y1=y1, x2=x2, y2=y2, x3=x3, y3=y3,
        center_x=center_x, center_y=center_y, radius=radius,
        major_radius=major_radius, minor_radius=minor_radius,
        start_angle=start_angle, end_angle=end_angle, angle=angle,
        x=x, y=y,
    )
    if err:
        return err
    match shape:
        case "line":
            return sketch_manager.draw_line(x1, y1, x2, y2)
        case "circle":
            return sketch_manager.draw_circle(center_x, center_y, radius)
        case "rectangle":
            return sketch_manager.draw_rectangle(x1, y1, x2, y2)
        case "arc":
            return sketch_manager.draw_arc(
                center_x, center_y, radius, start_angle, end_angle
            )
        case "polygon":
            return sketch_manager.draw_polygon(
                center_x, center_y, radius, sides
            )
        case "ellipse":
            return sketch_manager.draw_ellipse(
                center_x, center_y, major_radius, minor_radius, angle
            )
        case "spline":
            return sketch_manager.draw_spline(points or [])
        case "arc_3pt":
            return sketch_manager.draw_arc_by_3_points(
                x1, y1, center_x, center_y, x2, y2
            )
        case "circle_2pt":
            return sketch_manager.draw_circle_by_2_points(x1, y1, x2, y2)
        case "circle_3pt":
            return sketch_manager.draw_circle_by_3_points(
                x1, y1, x2, y2, x3, y3
            )
        case "point":
            return sketch_manager.draw_point(x, y)
        case "construction_line":
            return sketch_manager.draw_construction_line(x1, y1, x2, y2)
        case _:
            return {"error": f"Unknown shape: {shape}"}


# === Composite: sketch_modify (common operations) ===


def sketch_modify(
    action: str,
    radius: float = 0.0,
    distance: float = 0.0,
    center_x: float = 0.0,
    center_y: float = 0.0,
    angle_degrees: float = 0.0,
    scale_factor: float = 1.0,
    axis: str = "X",
) -> dict[str, Any]:
    """Modify sketch geometry with common operations.

    action: 'fillet' | 'chamfer' | 'offset' | 'rotate' | 'scale'
      | 'mirror' | 'paste'

    Distances in meters. angle_degrees in degrees.
    """
    err = validate_numerics(
        radius=radius, distance=distance, center_x=center_x,
        center_y=center_y, angle_degrees=angle_degrees,
        scale_factor=scale_factor,
    )
    if err:
        return err
    match action:
        case "fillet":
            return sketch_manager.sketch_fillet(radius)
        case "chamfer":
            return sketch_manager.sketch_chamfer(distance)
        case "offset":
            return sketch_manager.sketch_offset(distance)
        case "rotate":
            return sketch_manager.sketch_rotate(
                center_x, center_y, angle_degrees
            )
        case "scale":
            return sketch_manager.sketch_scale(
                center_x, center_y, scale_factor
            )
        case "mirror":
            return sketch_manager.sketch_mirror(axis)
        case "paste":
            return sketch_manager.sketch_paste()
        case _:
            return {"error": f"Unknown action: {action}"}


# === Composite: sketch_advanced_modify (specialized operations) ===


def sketch_advanced_modify(
    action: str,
    axis_x1: float = 0.0,
    axis_y1: float = 0.0,
    axis_x2: float = 0.0,
    axis_y2: float = 0.0,
    copy: bool = True,
    offset_side_x: float = 0.0,
    offset_side_y: float = 0.0,
    offset_distance: float = 0.0,
    clean_points: bool = True,
    clean_splines: bool = True,
    clean_identical: bool = True,
    clean_small: bool = True,
    small_tolerance: float = 0.0001,
) -> dict[str, Any]:
    """Specialized sketch modification operations.

    action: 'mirror_spline' | 'offset_2d' | 'clean'

    Coordinates and distances in meters.
    """
    err = validate_numerics(
        axis_x1=axis_x1, axis_y1=axis_y1, axis_x2=axis_x2,
        axis_y2=axis_y2, offset_side_x=offset_side_x,
        offset_side_y=offset_side_y, offset_distance=offset_distance,
        small_tolerance=small_tolerance,
    )
    if err:
        return err
    match action:
        case "mirror_spline":
            return sketch_manager.mirror_spline(
                axis_x1, axis_y1, axis_x2, axis_y2, copy
            )
        case "offset_2d":
            return sketch_manager.offset_sketch_2d(
                offset_side_x, offset_side_y, offset_distance
            )
        case "clean":
            return sketch_manager.clean_sketch_geometry(
                clean_points, clean_splines, clean_identical,
                clean_small, small_tolerance,
            )
        case _:
            return {"error": f"Unknown action: {action}"}


# === Composite: sketch_constraint ===


def sketch_constraint(
    type: str = "geometric",
    constraint_type: str = "",
    elements: list[Any] | None = None,
    element1_type: str = "",
    element1_index: int = 0,
    keypoint1: int = 0,
    element2_type: str = "",
    element2_index: int = 0,
    keypoint2: int = 0,
) -> dict[str, Any]:
    """Add a constraint to sketch elements.

    type: 'geometric' | 'keypoint'

    geometric: constraint_type (Horizontal, Vertical, etc.) + elements list.
    keypoint: connect two elements at specific keypoints.
    """
    match type:
        case "geometric":
            return sketch_manager.add_constraint(
                constraint_type, elements or []
            )
        case "keypoint":
            return sketch_manager.add_keypoint_constraint(
                element1_type, element1_index, keypoint1,
                element2_type, element2_index, keypoint2,
            )
        case _:
            return {"error": f"Unknown constraint type: {type}"}


# === Composite: sketch_project ===


def sketch_project(
    source: str,
    face_index: int = 0,
    edge_index: int = 0,
    plane_index: int = 1,
    face_indices: list[int] | None = None,
    x: float = 0.0,
    y: float = 0.0,
    tolerance: float = 0.001,
) -> dict[str, Any]:
    """Project external geometry into the active sketch.

    source: 'edge' | 'include_edge' | 'ref_plane' | 'silhouette'
            | 'region_faces' | 'chain' | 'to_curve'
    """
    err = validate_numerics(x=x, y=y, tolerance=tolerance)
    if err:
        return err
    match source:
        case "edge":
            return sketch_manager.project_edge(face_index, edge_index)
        case "include_edge":
            return sketch_manager.include_edge(face_index, edge_index)
        case "ref_plane":
            return sketch_manager.project_ref_plane(plane_index)
        case "silhouette":
            return sketch_manager.project_silhouette_edges()
        case "region_faces":
            return sketch_manager.include_region_faces(
                face_indices or []
            )
        case "chain":
            return sketch_manager.chain_locate(x, y, tolerance)
        case "to_curve":
            return sketch_manager.convert_to_curve()
        case _:
            return {"error": f"Unknown source: {source}"}


# === Registration ===


def register(mcp: Any) -> None:
    """Register sketching tools with the MCP server."""
    mcp.tool()(manage_sketch)
    mcp.tool()(draw)
    mcp.tool()(sketch_modify)
    mcp.tool()(sketch_advanced_modify)
    mcp.tool()(sketch_constraint)
    mcp.tool()(sketch_project)
