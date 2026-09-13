"""Sketching tools for Solid Edge MCP."""

from typing import Any, Literal

from solidedge_mcp.backends.validation import validate_numerics
from solidedge_mcp.managers import sketch_manager
from solidedge_mcp.tools._registry import register_tool

SketchElementType = Literal["line", "circle", "arc", "ellipse", "spline"]

# === Composite: manage_sketch ===


def manage_sketch(
    action: Literal[
        "create",
        "close",
        "create_on_plane",
        "set_axis",
        "set_visibility",
        "get_geometry",
    ] = "create",
    plane: Literal["Top", "Front", "Right", "XY", "XZ", "YZ"] = "Top",
    plane_index: int = 1,
    x1: float = 0.0,
    y1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    visible: bool = False,
    closed: bool = True,
) -> dict[str, Any]:
    """Create, close, or configure the active 2D sketch.

    create: on named plane (Top=XY, Right=YZ, Front=XZ).
    create_on_plane: 1-based plane_index (1=Top/XY, 2=Right/YZ, 3=Front/XZ,
    4+ = user planes). close: finish the profile; closed=True validates a
    closed region (needed for solids) and returns 'validation_code' (0=clean).
    set_axis: revolve axis from (x1,y1) to (x2,y2) in meters.
    set_visibility: visible. get_geometry: ordered element list.
    """
    err = validate_numerics(x1=x1, y1=y1, x2=x2, y2=y2)
    if err:
        return err
    match action:
        case "create":
            return sketch_manager.create_sketch(plane=plane)
        case "close":
            return sketch_manager.close_sketch(closed=closed)
        case "create_on_plane":
            return sketch_manager.create_sketch_on_plane_index(plane_index=plane_index)
        case "set_axis":
            return sketch_manager.set_axis_of_revolution(x1=x1, y1=y1, x2=x2, y2=y2)
        case "set_visibility":
            return sketch_manager.hide_profile(visible=visible)
        case "get_geometry":
            return sketch_manager.get_ordered_geometry()
        case _:
            return {"error": f"Unknown action: {action}"}


# === Composite: draw ===


def draw(
    shape: Literal[
        "line",
        "circle",
        "rectangle",
        "arc",
        "polygon",
        "ellipse",
        "spline",
        "arc_3pt",
        "circle_2pt",
        "circle_3pt",
        "point",
        "construction_line",
    ] = "line",
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
    """Draw one element in the active sketch. Meters; angles in degrees.

    line/construction_line/rectangle/circle_2pt: (x1,y1)-(x2,y2).
    circle: center_x/y + radius. arc: center + radius + start/end_angle.
    polygon: center + radius + sides. ellipse: center + major/minor_radius + angle.
    spline: points [[x,y],...]. arc_3pt: (x1,y1) start, (x2,y2) a point the arc
    passes through, (x3,y3) end. circle_3pt: three points. point: x,y.
    """
    err = validate_numerics(
        x1=x1,
        y1=y1,
        x2=x2,
        y2=y2,
        x3=x3,
        y3=y3,
        center_x=center_x,
        center_y=center_y,
        radius=radius,
        major_radius=major_radius,
        minor_radius=minor_radius,
        start_angle=start_angle,
        end_angle=end_angle,
        angle=angle,
        x=x,
        y=y,
    )
    if err:
        return err
    match shape:
        case "line":
            return sketch_manager.draw_line(x1=x1, y1=y1, x2=x2, y2=y2)
        case "circle":
            return sketch_manager.draw_circle(center_x=center_x, center_y=center_y, radius=radius)
        case "rectangle":
            return sketch_manager.draw_rectangle(x1=x1, y1=y1, x2=x2, y2=y2)
        case "arc":
            return sketch_manager.draw_arc(
                center_x=center_x,
                center_y=center_y,
                radius=radius,
                start_angle=start_angle,
                end_angle=end_angle,
            )
        case "polygon":
            return sketch_manager.draw_polygon(
                center_x=center_x, center_y=center_y, radius=radius, sides=sides
            )
        case "ellipse":
            return sketch_manager.draw_ellipse(
                center_x=center_x,
                center_y=center_y,
                major_radius=major_radius,
                minor_radius=minor_radius,
                angle=angle,
            )
        case "spline":
            return sketch_manager.draw_spline(points=points or [])
        case "arc_3pt":
            return sketch_manager.draw_arc_by_3_points(
                start_x=x1, start_y=y1, along_x=x2, along_y=y2, end_x=x3, end_y=y3
            )
        case "circle_2pt":
            return sketch_manager.draw_circle_by_2_points(x1=x1, y1=y1, x2=x2, y2=y2)
        case "circle_3pt":
            return sketch_manager.draw_circle_by_3_points(x1=x1, y1=y1, x2=x2, y2=y2, x3=x3, y3=y3)
        case "point":
            return sketch_manager.draw_point(x=x, y=y)
        case "construction_line":
            return sketch_manager.draw_construction_line(x1=x1, y1=y1, x2=x2, y2=y2)
        case _:
            return {"error": f"Unknown shape: {shape}"}


# === Composite: sketch_modify (common operations) ===


def sketch_modify(
    action: Literal["fillet", "chamfer", "offset", "rotate", "scale", "mirror", "paste"],
    radius: float = 0.0,
    distance: float = 0.0,
    center_x: float = 0.0,
    center_y: float = 0.0,
    angle_degrees: float = 0.0,
    scale_factor: float = 1.0,
    axis: Literal["X", "Y"] = "X",
) -> dict[str, Any]:
    """Modify all geometry in the active sketch. Meters.

    fillet: radius. chamfer/offset: distance. rotate: center_x/y + angle_degrees.
    scale: center_x/y + scale_factor. mirror: axis X or Y (adds mirrored copies).
    paste: paste clipboard geometry.
    """
    err = validate_numerics(
        radius=radius,
        distance=distance,
        center_x=center_x,
        center_y=center_y,
        angle_degrees=angle_degrees,
        scale_factor=scale_factor,
    )
    if err:
        return err
    match action:
        case "fillet":
            return sketch_manager.sketch_fillet(radius=radius)
        case "chamfer":
            return sketch_manager.sketch_chamfer(distance=distance)
        case "offset":
            return sketch_manager.sketch_offset(distance=distance)
        case "rotate":
            return sketch_manager.sketch_rotate(
                center_x=center_x, center_y=center_y, angle_degrees=angle_degrees
            )
        case "scale":
            return sketch_manager.sketch_scale(
                center_x=center_x, center_y=center_y, scale_factor=scale_factor
            )
        case "mirror":
            return sketch_manager.sketch_mirror(axis=axis)
        case "paste":
            return sketch_manager.sketch_paste()
        case _:
            return {"error": f"Unknown action: {action}"}


# === Composite: sketch_advanced_modify (specialized operations) ===


def sketch_advanced_modify(
    action: Literal["mirror_spline", "offset_2d", "clean"],
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
    """Specialized sketch operations. Meters.

    mirror_spline: axis (axis_x1,axis_y1)-(axis_x2,axis_y2); copy keeps original.
    offset_2d: offset toward point (offset_side_x, offset_side_y) by offset_distance.
    clean: delete stray points/splines/duplicates/tiny elements (< small_tolerance)
    per the clean_* flags.
    """
    err = validate_numerics(
        axis_x1=axis_x1,
        axis_y1=axis_y1,
        axis_x2=axis_x2,
        axis_y2=axis_y2,
        offset_side_x=offset_side_x,
        offset_side_y=offset_side_y,
        offset_distance=offset_distance,
        small_tolerance=small_tolerance,
    )
    if err:
        return err
    match action:
        case "mirror_spline":
            return sketch_manager.mirror_spline(
                axis_x1=axis_x1, axis_y1=axis_y1, axis_x2=axis_x2, axis_y2=axis_y2, copy=copy
            )
        case "offset_2d":
            return sketch_manager.offset_sketch_2d(
                offset_side_x=offset_side_x,
                offset_side_y=offset_side_y,
                offset_distance=offset_distance,
            )
        case "clean":
            return sketch_manager.clean_sketch_geometry(
                clean_points=clean_points,
                clean_splines=clean_splines,
                clean_identical=clean_identical,
                clean_small=clean_small,
                small_tolerance=small_tolerance,
            )
        case _:
            return {"error": f"Unknown action: {action}"}


# === Composite: sketch_constraint ===


def sketch_constraint(
    type: Literal["geometric", "keypoint"] = "geometric",
    constraint_type: Literal[
        "Horizontal",
        "Vertical",
        "Parallel",
        "Perpendicular",
        "Equal",
        "Concentric",
        "Tangent",
    ] = "Horizontal",
    elements: list[list[str | int]] | None = None,
    element1_type: SketchElementType = "line",
    element1_index: int = 0,
    keypoint1: int = 0,
    element2_type: SketchElementType = "line",
    element2_index: int = 0,
    keypoint2: int = 0,
) -> dict[str, Any]:
    """Add a 2D relation in the active sketch. Element indices are 1-BASED.

    geometric: constraint_type + elements as [type, index] pairs, e.g.
    [["line", 1], ["line", 2]]; type in line/circle/arc/ellipse/spline.
    Horizontal/Vertical need 1 element, the rest need 2.
    keypoint: weld element1 keypoint1 to element2 keypoint2. Keypoints are
    0=start, 1=end, 2=midpoint (lines/arcs); 0=center (circles).
    """
    match type:
        case "geometric":
            return sketch_manager.add_constraint(
                constraint_type=constraint_type, elements=elements or []
            )
        case "keypoint":
            return sketch_manager.add_keypoint_constraint(
                element1_type=element1_type,
                element1_index=element1_index,
                keypoint1=keypoint1,
                element2_type=element2_type,
                element2_index=element2_index,
                keypoint2=keypoint2,
            )
        case _:
            return {"error": f"Unknown constraint type: {type}"}


# === Composite: sketch_project ===


def sketch_project(
    source: Literal[
        "edge",
        "include_edge",
        "ref_plane",
        "silhouette",
        "region_faces",
        "chain",
        "to_curve",
    ],
    face_index: int = 0,
    edge_index: int = 0,
    plane_index: int = 1,
    face_indices: list[int] | None = None,
    x: float = 0.0,
    y: float = 0.0,
    tolerance: float = 0.001,
) -> dict[str, Any]:
    """Bring external geometry into the active sketch.

    edge/include_edge: 0-based face_index + edge_index. ref_plane: 1-based
    plane_index (1=Top/XY, 2=Right/YZ, 3=Front/XZ). silhouette: body outline.
    region_faces: 0-based face_indices. chain: pick a chain near (x,y) within
    tolerance (meters). to_curve: convert included geometry to curves.
    """
    err = validate_numerics(x=x, y=y, tolerance=tolerance)
    if err:
        return err
    match source:
        case "edge":
            return sketch_manager.project_edge(face_index=face_index, edge_index=edge_index)
        case "include_edge":
            return sketch_manager.include_edge(face_index=face_index, edge_index=edge_index)
        case "ref_plane":
            return sketch_manager.project_ref_plane(plane_index=plane_index)
        case "silhouette":
            return sketch_manager.project_silhouette_edges()
        case "region_faces":
            return sketch_manager.include_region_faces(face_indices=face_indices or [])
        case "chain":
            return sketch_manager.chain_locate(x=x, y=y, tolerance=tolerance)
        case "to_curve":
            return sketch_manager.convert_to_curve()
        case _:
            return {"error": f"Unknown source: {source}"}


# === Registration ===


def register(mcp: Any) -> None:
    """Register sketching tools with the MCP server."""
    tags = {"sketch"}
    register_tool(mcp, manage_sketch, tags=tags)
    register_tool(mcp, draw, tags=tags)
    register_tool(mcp, sketch_modify, tags=tags)
    register_tool(mcp, sketch_advanced_modify, tags=tags, destructive=True)
    register_tool(mcp, sketch_constraint, tags=tags)
    register_tool(mcp, sketch_project, tags=tags)
