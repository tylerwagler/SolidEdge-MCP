"""Query and inspection tools for Solid Edge MCP.

Read-only getters have been migrated to MCP Resources (see resources.py).
This module retains measurement tools, setters, actions, feature editing,
and B-Rep topology queries.

Composite tools use a discriminator parameter (type/action/property/target)
to dispatch to the correct backend method via match/case.
"""

from typing import Any

from solidedge_mcp.managers import query_manager

# ── Group 59: measure ──────────────────────────────────────────────


def measure(
    type: str = "distance",
    x1: float = 0.0,
    y1: float = 0.0,
    z1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    z2: float = 0.0,
    x3: float = 0.0,
    y3: float = 0.0,
    z3: float = 0.0,
) -> dict[str, Any]:
    """Measure distance or angle between 3D points.

    type: 'distance' | 'angle'

    Coordinates in meters. 'angle' returns degrees at vertex (x2,y2,z2).
    """
    match type:
        case "distance":
            return query_manager.measure_distance(x1, y1, z1, x2, y2, z2)
        case "angle":
            return query_manager.measure_angle(x1, y1, z1, x2, y2, z2, x3, y3, z3)
        case _:
            return {"error": f"Unknown type: {type}"}


# ── Group 60: manage_variable ──────────────────────────────────────


def manage_variable(
    action: str = "set",
    name: str = "",
    value: float | None = None,
    formula: str | None = None,
    units_type: str | None = None,
    new_name: str | None = None,
    pattern: str = "*",
    case_insensitive: bool = True,
) -> dict[str, Any]:
    """Manage document variables.

    action: 'set' | 'add' | 'query' | 'rename' | 'translate'
            | 'copy_clipboard' | 'add_from_clipboard' | 'set_formula'
    """
    match action:
        case "set":
            if value is None:
                return {"error": "value is required for 'set' action"}
            return query_manager.set_variable(name, value)
        case "add":
            if formula is None:
                return {"error": "formula is required for 'add' action"}
            return query_manager.add_variable(name, formula, units_type)
        case "query":
            return query_manager.query_variables(pattern, case_insensitive)
        case "rename":
            if new_name is None:
                return {"error": "new_name is required for 'rename' action"}
            return query_manager.rename_variable(name, new_name)
        case "translate":
            return query_manager.translate_variable(name)
        case "copy_clipboard":
            return query_manager.copy_variable_to_clipboard(name)
        case "add_from_clipboard":
            return query_manager.add_variable_from_clipboard(name, units_type)
        case "set_formula":
            if formula is None:
                return {"error": "formula is required for 'set_formula' action"}
            return query_manager.set_variable_formula(name, formula)
        case _:
            return {"error": f"Unknown action: {action}"}


# ── Group 61: manage_property ──────────────────────────────────────


def manage_property(
    action: str,
    name: str = "",
    value: str = "",
) -> dict[str, Any]:
    """Manage document and custom properties.

    action: 'set_document' | 'set_custom' | 'delete_custom'
    """
    match action:
        case "set_document":
            return query_manager.set_document_property(name, value)
        case "set_custom":
            return query_manager.set_custom_property(name, value)
        case "delete_custom":
            return query_manager.delete_custom_property(name)
        case _:
            return {"error": f"Unknown action: {action}"}


# ── Group 62: manage_material ──────────────────────────────────────


def manage_material(
    action: str = "set",
    material_name: str = "",
    density: float = 0.0,
) -> dict[str, Any]:
    """Manage material assignment and density.

    action: 'set' | 'set_density' | 'set_by_name' | 'get_library'

    density is in kg/m3.
    """
    match action:
        case "set":
            return query_manager.set_material(material_name)
        case "set_density":
            return query_manager.set_material_density(density)
        case "set_by_name":
            return query_manager.set_material_by_name(material_name)
        case "get_library":
            return query_manager.get_material_library()
        case _:
            return {"error": f"Unknown action: {action}"}


# ── Group 63: set_appearance ───────────────────────────────────────


def set_appearance(
    target: str,
    red: int = 0,
    green: int = 0,
    blue: int = 0,
    face_index: int = 0,
    opacity: float = 1.0,
    reflectivity: float = 0.0,
) -> dict[str, Any]:
    """Set visual appearance of the active part body or a face.

    target: 'body_color' | 'face_color' | 'opacity' | 'reflectivity'

    RGB values 0-255. opacity/reflectivity 0.0-1.0.
    """
    match target:
        case "body_color":
            return query_manager.set_body_color(red, green, blue)
        case "face_color":
            return query_manager.set_face_color(face_index, red, green, blue)
        case "opacity":
            return query_manager.set_body_opacity(opacity)
        case "reflectivity":
            return query_manager.set_body_reflectivity(reflectivity)
        case _:
            return {"error": f"Unknown target: {target}"}


# ── Group 64: manage_layer ─────────────────────────────────────────


def manage_layer(
    action: str,
    name_or_index: str | int = "",
    show: bool | None = None,
    selectable: bool | None = None,
) -> dict[str, Any]:
    """Manage layers in the active document.

    action: 'add' | 'activate' | 'set_properties' | 'delete'

    name_or_index accepts either a string name or integer index.
    """
    match action:
        case "add":
            if not isinstance(name_or_index, str):
                return {"error": "name_or_index must be a string for 'add' action"}
            return query_manager.add_layer(name_or_index)
        case "activate":
            return query_manager.activate_layer(name_or_index)
        case "set_properties":
            return query_manager.set_layer_properties(name_or_index, show, selectable)
        case "delete":
            return query_manager.delete_layer(name_or_index)
        case _:
            return {"error": f"Unknown action: {action}"}


# ── Group 65: select_set ──────────────────────────────────────────


def select_set(
    action: str,
    object_type: str = "",
    index: int = 0,
) -> dict[str, Any]:
    """Manage the document selection set.

    action: 'clear' | 'add' | 'remove' | 'all' | 'copy' | 'cut'
            | 'delete' | 'suspend_display' | 'resume_display'
            | 'refresh_display'
    """
    match action:
        case "clear":
            return query_manager.clear_select_set()
        case "add":
            return query_manager.select_add(object_type, index)
        case "remove":
            return query_manager.select_remove(index)
        case "all":
            return query_manager.select_all()
        case "copy":
            return query_manager.select_copy()
        case "cut":
            return query_manager.select_cut()
        case "delete":
            return query_manager.select_delete()
        case "suspend_display":
            return query_manager.select_suspend_display()
        case "resume_display":
            return query_manager.select_resume_display()
        case "refresh_display":
            return query_manager.select_refresh_display()
        case _:
            return {"error": f"Unknown action: {action}"}


# ── Group 66: edit_feature_extent ─────────────────────────────────


def edit_feature_extent(
    property: str,
    feature_name: str = "",
    extent_type: int = 0,
    distance: float = 0.0,
    wall_type: int = 0,
    thickness1: float = 0.0,
    thickness2: float = 0.0,
    offset: float = 0.0,
    body_indices: list[int] | None = None,
    offset_side: int = 0,
    treatment_type: int = 0,
    draft_side: int = 0,
    draft_angle: float = 0.0,
    crown_type: int = 0,
    crown_side: int = 0,
    crown_curvature_side: int = 0,
    crown_radius_or_offset: float = 0.0,
    crown_takeoff_angle: float = 0.0,
) -> dict[str, Any]:
    """Edit feature extent, thin wall, face offset, body array, and treatment properties.

    property: 'get_direction1' | 'set_direction1' | 'get_direction2' | 'set_direction2'
              | 'get_thin_wall' | 'set_thin_wall' | 'get_from_face' | 'set_from_face'
              | 'get_body_array' | 'set_body_array' | 'get_to_face' | 'set_to_face'
              | 'get_direction1_treatment' | 'apply_direction1_treatment'

    extent_type: 13=Finite, 16=ThroughAll, 44=None. offset_side: 1=igLeft, 2=igRight.
    Distances/thicknesses in meters. Angles in degrees. body_indices are 0-based.
    """
    match property:
        case "get_direction1":
            return query_manager.get_direction1_extent(feature_name)
        case "set_direction1":
            return query_manager.set_direction1_extent(feature_name, extent_type, distance)
        case "get_direction2":
            return query_manager.get_direction2_extent(feature_name)
        case "set_direction2":
            return query_manager.set_direction2_extent(feature_name, extent_type, distance)
        case "get_thin_wall":
            return query_manager.get_thin_wall_options(feature_name)
        case "set_thin_wall":
            return query_manager.set_thin_wall_options(
                feature_name,
                wall_type,
                thickness1,
                thickness2,
            )
        case "get_from_face":
            return query_manager.get_from_face_offset(feature_name)
        case "set_from_face":
            return query_manager.set_from_face_offset(feature_name, offset)
        case "get_body_array":
            return query_manager.get_body_array(feature_name)
        case "set_body_array":
            return query_manager.set_body_array(feature_name, body_indices or [])
        case "get_to_face":
            return query_manager.get_to_face_offset(feature_name)
        case "set_to_face":
            return query_manager.set_to_face_offset(feature_name, offset_side, distance)
        case "get_direction1_treatment":
            return query_manager.get_direction1_treatment(feature_name)
        case "apply_direction1_treatment":
            return query_manager.apply_direction1_treatment(
                feature_name,
                treatment_type,
                draft_side,
                draft_angle,
                crown_type,
                crown_side,
                crown_curvature_side,
                crown_radius_or_offset,
                crown_takeoff_angle,
            )
        case _:
            return {"error": f"Unknown property: {property}"}


# ── Group 67: manage_feature_tree ─────────────────────────────────


def manage_feature_tree(
    action: str,
    feature_name: str = "",
    new_name: str = "",
    mode: str = "",
) -> dict[str, Any]:
    """Manage features in the design tree.

    action: 'rename' | 'suppress' | 'unsuppress' | 'set_mode'

    mode: 'ordered' or 'synchronous' (for set_mode).
    """
    match action:
        case "rename":
            return query_manager.rename_feature(feature_name, new_name)
        case "suppress":
            return query_manager.suppress_feature(feature_name)
        case "unsuppress":
            return query_manager.unsuppress_feature(feature_name)
        case "set_mode":
            return query_manager.set_modeling_mode(mode)
        case _:
            return {"error": f"Unknown action: {action}"}


# ── Group 68: query_edge ──────────────────────────────────────────


def query_edge(
    property: str,
    face_index: int = 0,
    edge_index: int = 0,
    param: float = 0.5,
    which: str = "start",
) -> dict[str, Any]:
    """Query edge topology and geometry on a face.

    property: 'endpoints' | 'length' | 'tangent' | 'geometry' | 'curvature' | 'vertex'

    param is a 0.0-1.0 parametric position along the edge. which: 'start' or 'end'.
    """
    match property:
        case "endpoints":
            return query_manager.get_edge_endpoints(face_index, edge_index)
        case "length":
            return query_manager.get_edge_length(face_index, edge_index)
        case "tangent":
            return query_manager.get_edge_tangent(face_index, edge_index, param)
        case "geometry":
            return query_manager.get_edge_geometry(face_index, edge_index)
        case "curvature":
            return query_manager.get_edge_curvature(face_index, edge_index, param)
        case "vertex":
            return query_manager.get_vertex_point(face_index, edge_index, which)
        case _:
            return {"error": f"Unknown property: {property}"}


# ── Group 69: query_face ──────────────────────────────────────────


def query_face(
    property: str,
    face_index: int = 0,
    u: float = 0.5,
    v: float = 0.5,
) -> dict[str, Any]:
    """Query face topology and geometry.

    property: 'normal' | 'geometry' | 'loops' | 'curvature'

    u, v are parametric coordinates on the face (0.0-1.0).
    """
    match property:
        case "normal":
            return query_manager.get_face_normal(face_index, u, v)
        case "geometry":
            return query_manager.get_face_geometry(face_index)
        case "loops":
            return query_manager.get_face_loops(face_index)
        case "curvature":
            return query_manager.get_face_curvature(face_index, u, v)
        case _:
            return {"error": f"Unknown property: {property}"}


# ── Group 70: query_body ──────────────────────────────────────────


def query_body(
    property: str,
    direction_x: float = 0.0,
    direction_y: float = 0.0,
    direction_z: float = 0.0,
    origin_x: float = 0.0,
    origin_y: float = 0.0,
    origin_z: float = 0.0,
    x: float = 0.0,
    y: float = 0.0,
    z: float = 0.0,
    shell_index: int = 0,
    tolerance: float = 0.0,
) -> dict[str, Any]:
    """Query body-level topology and geometry.

    property: 'extreme_point' | 'faces_by_ray' | 'shells' | 'vertices'
              | 'shell_info' | 'point_inside' | 'user_physical_properties'
              | 'facet_data'

    All coordinates and tolerance in meters.
    """
    match property:
        case "extreme_point":
            return query_manager.get_body_extreme_point(direction_x, direction_y, direction_z)
        case "faces_by_ray":
            return query_manager.get_faces_by_ray(
                origin_x,
                origin_y,
                origin_z,
                direction_x,
                direction_y,
                direction_z,
            )
        case "shells":
            return query_manager.get_body_shells()
        case "vertices":
            return query_manager.get_body_vertices()
        case "shell_info":
            return query_manager.get_shell_info(shell_index)
        case "point_inside":
            return query_manager.is_point_inside_body(x, y, z)
        case "user_physical_properties":
            return query_manager.get_user_physical_properties()
        case "facet_data":
            return query_manager.get_body_facet_data(tolerance)
        case _:
            return {"error": f"Unknown property: {property}"}


# ── Group 71: query_bspline ───────────────────────────────────────


def query_bspline(
    type: str,
    face_index: int = 0,
    edge_index: int = 0,
) -> dict[str, Any]:
    """Query B-spline (NURBS) metadata from edges or faces.

    type: 'curve' | 'surface'
    """
    match type:
        case "curve":
            return query_manager.get_bspline_curve_info(face_index, edge_index)
        case "surface":
            return query_manager.get_bspline_surface_info(face_index)
        case _:
            return {"error": f"Unknown type: {type}"}


# ── Composite: recompute ──────────────────────────────────────────


def recompute(scope: str = "model") -> dict[str, Any]:
    """Recompute the active model or document.

    scope: 'model' | 'document'
    """
    match scope:
        case "model":
            return query_manager.recompute()
        case "document":
            return query_manager.recompute_document()
        case _:
            return {"error": f"Unknown scope: {scope}"}


# ── Registration ──────────────────────────────────────────────────


def register(mcp: Any) -> None:
    """Register query tools with the MCP server."""
    mcp.tool()(measure)
    mcp.tool()(manage_variable)
    mcp.tool()(manage_property)
    mcp.tool()(manage_material)
    mcp.tool()(set_appearance)
    mcp.tool()(manage_layer)
    mcp.tool()(select_set)
    mcp.tool()(edit_feature_extent)
    mcp.tool()(manage_feature_tree)
    mcp.tool()(query_edge)
    mcp.tool()(query_face)
    mcp.tool()(query_body)
    mcp.tool()(query_bspline)
    mcp.tool()(recompute)
