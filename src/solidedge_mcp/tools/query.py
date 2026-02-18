"""Query and inspection tools for Solid Edge MCP.

Read-only getters have been migrated to MCP Resources (see resources.py).
This module retains measurement tools, setters, actions, feature editing,
and B-Rep topology queries.

Composite tools use a discriminator parameter (type/action/property/target)
to dispatch to the correct backend method via match/case.
"""

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
) -> dict:
    """Measure distance or angle between 3D points.

    type: 'distance' (default) | 'angle'

    - distance: Euclidean distance between (x1,y1,z1) and (x2,y2,z2).
    - angle: Angle at vertex (x2,y2,z2) between rays to
      (x1,y1,z1) and (x3,y3,z3).
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
) -> dict:
    """Manage document variables.

    action: 'set' (default) | 'add' | 'query' | 'rename'
            | 'translate' | 'copy_clipboard'
            | 'add_from_clipboard' | 'set_formula'

    - set: Set variable `name` to `value` (float).
    - add: Add new variable `name` with `formula`
      (and optional `units_type`).
    - query: Query variables matching `pattern`
      (default '*', `case_insensitive`).
    - rename: Rename variable `name` to `new_name`.
    - translate: Look up variable `name` via Variables.Translate().
    - copy_clipboard: Copy variable `name` definition to clipboard.
    - add_from_clipboard: Add variable `name` from clipboard
      (optional `units_type`).
    - set_formula: Set formula of existing variable `name`
      to `formula`.
    """
    match action:
        case "set":
            return query_manager.set_variable(name, value)
        case "add":
            return query_manager.add_variable(name, formula, units_type)
        case "query":
            return query_manager.query_variables(pattern, case_insensitive)
        case "rename":
            return query_manager.rename_variable(name, new_name)
        case "translate":
            return query_manager.translate_variable(name)
        case "copy_clipboard":
            return query_manager.copy_variable_to_clipboard(name)
        case "add_from_clipboard":
            return query_manager.add_variable_from_clipboard(name, units_type)
        case "set_formula":
            return query_manager.set_variable_formula(name, formula)
        case _:
            return {"error": f"Unknown action: {action}"}


# ── Group 61: manage_property ──────────────────────────────────────


def manage_property(
    action: str,
    name: str = "",
    value: str = "",
) -> dict:
    """Manage document and custom properties.

    action: 'set_document' | 'set_custom' | 'delete_custom'

    - set_document: Set a summary/document property `name`
      to `value`.
    - set_custom: Set or create a custom property `name`
      to `value`.
    - delete_custom: Delete a custom property by `name`.
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
) -> dict:
    """Manage material assignment and density.

    action: 'set' (default) | 'set_density' | 'set_by_name'
            | 'get_library'

    - set: Assign material `material_name` to active part.
    - set_density: Set material density for mass calculations.
    - set_by_name: Look up material by name and apply it.
    - get_library: Get the full material library with names,
      densities, and properties.
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
) -> dict:
    """Set visual appearance of the active part body or face.

    target: 'body_color' | 'face_color' | 'opacity'
            | 'reflectivity'

    - body_color: Set body color using `red`, `green`, `blue`
      (0-255).
    - face_color: Set color of face at `face_index` using
      `red`, `green`, `blue`.
    - opacity: Set body opacity (0.0=transparent,
      1.0=opaque).
    - reflectivity: Set body reflectivity (0.0 to 1.0).
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
) -> dict:
    """Manage layers in the active document.

    action: 'add' | 'activate' | 'set_properties' | 'delete'

    - add: Add a new layer with `name_or_index` (str name).
    - activate: Activate a layer by name (str) or index (int).
    - set_properties: Set visibility (`show`) and selectability
      (`selectable`) on a layer identified by `name_or_index`.
    - delete: Delete a layer by name (str) or index (int).
    """
    match action:
        case "add":
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
) -> dict:
    """Manage the document selection set.

    action: 'clear' | 'add' | 'remove' | 'all' | 'copy' | 'cut'
            | 'delete' | 'suspend_display' | 'resume_display'
            | 'refresh_display'

    - clear: Clear the current selection set.
    - add: Add an object of `object_type` at `index` to selection.
    - remove: Remove object at `index` from selection set.
    - all: Select all objects in the active document.
    - copy: Copy current selection to clipboard.
    - cut: Cut current selection to clipboard.
    - delete: Delete currently selected objects.
    - suspend_display: Suspend display updates (batch perf).
    - resume_display: Resume display updates.
    - refresh_display: Refresh the display of the selection set.
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
) -> dict:
    """Edit feature extent, thin wall, face offset, body array,
    and treatment properties.

    property: 'get_direction1' | 'set_direction1'
      | 'get_direction2' | 'set_direction2'
      | 'get_thin_wall' | 'set_thin_wall'
      | 'get_from_face' | 'set_from_face'
      | 'get_body_array' | 'set_body_array'
      | 'get_to_face' | 'set_to_face'
      | 'get_direction1_treatment'
      | 'apply_direction1_treatment'

    - get_direction1: Get Direction 1 extent params from
      `feature_name`.
    - set_direction1: Set Direction 1 extent on `feature_name`
      (extent_type: 13=Finite, 16=ThroughAll, 44=None;
      distance).
    - get_direction2: Get Direction 2 extent params.
    - set_direction2: Set Direction 2 extent.
    - get_thin_wall: Get thin wall options from `feature_name`.
    - set_thin_wall: Set thin wall options (wall_type,
      thickness1, thickness2).
    - get_from_face: Get 'from face' offset data.
    - set_from_face: Set 'from face' offset distance.
    - get_body_array: Get multi-body references from feature.
    - set_body_array: Set body array using 0-based
      `body_indices`.
    - get_to_face: Get 'to face' offset data.
    - set_to_face: Set 'to face' offset (offset_side:
      1=igLeft, 2=igRight; distance).
    - get_direction1_treatment: Get Direction 1 treatment
      (crown/draft).
    - apply_direction1_treatment: Apply Direction 1 treatment
      with treatment_type, draft_side, draft_angle, crown_type,
      crown_side, crown_curvature_side,
      crown_radius_or_offset, crown_takeoff_angle.
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
            return query_manager.set_body_array(feature_name, body_indices)
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
) -> dict:
    """Manage features in the design tree.

    action: 'rename' | 'suppress' | 'unsuppress' | 'set_mode'

    - rename: Rename feature `feature_name` to `new_name`.
    - suppress: Suppress feature by `feature_name`.
    - unsuppress: Unsuppress feature by `feature_name`.
    - set_mode: Set modeling mode ('ordered' or 'synchronous').
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
) -> dict:
    """Query edge topology and geometry on a face.

    property: 'endpoints' | 'length' | 'tangent' | 'geometry'
      | 'curvature' | 'vertex'

    - endpoints: Get start/end XYZ of edge.
    - length: Get total length of the edge.
    - tangent: Get tangent vector at `param` (0.0-1.0).
    - geometry: Get underlying geometry type and data.
    - curvature: Get curvature at `param` on the edge.
    - vertex: Get XYZ of a vertex (which='start' or 'end').
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
) -> dict:
    """Query face topology and geometry.

    property: 'normal' | 'geometry' | 'loops' | 'curvature'

    - normal: Get normal vector at parametric (u, v) on face.
    - geometry: Get geometry type and data of the face
      (Plane/Cylinder/Cone/etc).
    - loops: Get loop info (outer boundary vs holes).
    - curvature: Get principal curvatures at parametric
      (u, v).
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
) -> dict:
    """Query body-level topology and geometry.

    property: 'extreme_point' | 'faces_by_ray' | 'shells'
      | 'vertices' | 'shell_info' | 'point_inside'
      | 'user_physical_properties' | 'facet_data'

    - extreme_point: Get extreme point in direction.
    - faces_by_ray: Ray-cast to find hit faces.
    - shells: List all shells with is_closed and volume.
    - vertices: Get all vertices with XYZ coordinates.
    - shell_info: Get shell topology at `shell_index`.
    - point_inside: Test if (x,y,z) is inside body.
    - user_physical_properties: Get user-overridden props.
    - facet_data: Get tessellation/mesh data (tolerance).
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
) -> dict:
    """Query B-spline (NURBS) metadata from edges or faces.

    type: 'curve' | 'surface'

    - curve: Get NURBS curve metadata (order, poles, knots,
      flags) from edge at `face_index`, `edge_index`.
    - surface: Get NURBS surface metadata (order, poles,
      knots, flags) from face at `face_index`.
    """
    match type:
        case "curve":
            return query_manager.get_bspline_curve_info(face_index, edge_index)
        case "surface":
            return query_manager.get_bspline_surface_info(face_index)
        case _:
            return {"error": f"Unknown type: {type}"}


# ── Composite: recompute ──────────────────────────────────────────


def recompute(scope: str = "model") -> dict:
    """Recompute the active model or document.

    scope: 'model' | 'document'

    - model: Recompute the active model features.
    - document: Force a full document-level recompute.
    """
    match scope:
        case "model":
            return query_manager.recompute()
        case "document":
            return query_manager.recompute_document()
        case _:
            return {"error": f"Unknown scope: {scope}"}


# ── Registration ──────────────────────────────────────────────────


def register(mcp):
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
