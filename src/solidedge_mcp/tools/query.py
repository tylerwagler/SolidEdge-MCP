"""Query and inspection tools for Solid Edge MCP.

Read-only getters have been migrated to MCP Resources (see resources.py).
This module retains measurement tools, setters, actions, feature editing,
and B-Rep topology queries.

Composite tools use a discriminator parameter (type/action/property/target)
to dispatch to the correct backend method via match/case.
"""

from typing import Any, Literal

from solidedge_mcp.backends.constants import DirectionConstants, ExtentTypeConstants
from solidedge_mcp.backends.query import DEFAULT_PAGE_LIMIT
from solidedge_mcp.managers import query_manager
from solidedge_mcp.tools._registry import register_tool

# ── Group 59: measure ──────────────────────────────────────────────


def measure(
    type: Literal["distance", "angle"] = "distance",
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
    """Measure between 3D points (meters). Read-only.

    distance: point1 to point2. angle: degrees at vertex point2 between
    point1 and point3.
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
    action: Literal[
        "set",
        "add",
        "query",
        "rename",
        "translate",
        "copy_clipboard",
        "add_from_clipboard",
        "set_formula",
    ] = "set",
    name: str = "",
    value: float | None = None,
    formula: str | None = None,
    units_type: str | None = None,
    new_name: str | None = None,
    pattern: str = "*",
    case_insensitive: bool = True,
) -> dict[str, Any]:
    """Manage document variables (Variable Table).

    set: name + value (document units, typically meters). add: name + formula
    [+ units_type]. set_formula: name + formula. rename: name + new_name.
    query: wildcard pattern [+ case_insensitive]. translate/copy_clipboard/
    add_from_clipboard: name [+ units_type].
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
    action: Literal["set_document", "set_custom", "delete_custom"],
    name: str = "",
    value: str = "",
) -> dict[str, Any]:
    """Set a document (Title, Author, ...) or custom property, or delete a custom one."""
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
    action: Literal["set", "set_density", "set_by_name", "get_library"] = "set",
    material_name: str = "",
    density: float = 0.0,
) -> dict[str, Any]:
    """Assign material to the active part.

    set/set_by_name: material_name from the material library.
    set_density: density in kg/m3. get_library: list library materials.
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
    target: Literal["body_color", "face_color", "opacity", "reflectivity"],
    red: int = 0,
    green: int = 0,
    blue: int = 0,
    face_index: int = 0,
    opacity: float = 1.0,
    reflectivity: float = 0.0,
) -> dict[str, Any]:
    """Set part body or face appearance.

    body_color: RGB 0-255. face_color: 0-based face_index + RGB.
    opacity/reflectivity: 0.0-1.0 on the body.
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
    action: Literal["add", "activate", "set_properties", "delete"],
    name_or_index: str | int = "",
    show: bool | None = None,
    selectable: bool | None = None,
) -> dict[str, Any]:
    """Manage document layers by name (str) or 0-based index (int).

    add requires a name. set_properties: show/selectable (None = unchanged).
    delete removes the layer.
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
    action: Literal[
        "clear",
        "add",
        "remove",
        "all",
        "copy",
        "cut",
        "delete",
        "suspend_display",
        "resume_display",
        "refresh_display",
    ],
    object_type: Literal["feature", "face", "plane"] = "feature",
    index: int = 0,
) -> dict[str, Any]:
    """Manipulate the document SelectSet.

    add: object_type + 0-based index. remove: 0-based index within the
    selection. cut/delete remove the selected objects from the model.
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

_EXTENT_TYPE_CONSTANTS: dict[str, int] = {
    "finite": ExtentTypeConstants.igFinite,
    "through_all": ExtentTypeConstants.igThroughAll,
    "none": ExtentTypeConstants.igNone,
}

_OFFSET_SIDE_CONSTANTS: dict[str, int] = {
    "left": DirectionConstants.igLeft,
    "right": DirectionConstants.igRight,
}

#: constant.tlb > FeaturePropertyConstants, the ExtentSide/ThicknessSide slot.
_SIDE_CONSTANTS: dict[str, int] = {
    "left": DirectionConstants.igLeft,
    "right": DirectionConstants.igRight,
    "symmetric": DirectionConstants.igSymmetric,
}


def edit_feature_extent(
    property: Literal[
        "get_direction1",
        "set_direction1",
        "get_direction2",
        "set_direction2",
        "get_thin_wall",
        "set_thin_wall",
        "get_from_face",
        "set_from_face",
        "get_body_array",
        "set_body_array",
        "get_to_face",
        "set_to_face",
        "get_direction1_treatment",
        "apply_direction1_treatment",
    ],
    feature_name: str = "",
    extent_type: Literal["finite", "through_all", "none"] = "finite",
    distance: float = 0.0,
    extent_side: Literal["left", "right", "symmetric"] = "right",
    thickness: float = 0.0,
    thickness_side: Literal["left", "right", "symmetric"] = "right",
    thin_wall: bool = True,
    add_end_caps: bool = False,
    remove_inside_material: bool = False,
    offset: float = 0.0,
    body_indices: list[int] | None = None,
    multi_body_cut: bool = True,
    offset_side: Literal["left", "right"] = "left",
    treatment_type: int = 0,
    draft_side: int = 0,
    draft_angle: float = 0.0,
    crown_type: int = 0,
    crown_side: int = 0,
    crown_curvature_side: int = 0,
    crown_radius_or_offset: float = 0.0,
    crown_takeoff_angle: float = 0.0,
) -> dict[str, Any]:
    """Get/set extent data on the named feature. Meters; angles in degrees.

    set_direction1/2: extent_type + distance (finite only) + extent_side.
    set_thin_wall: thickness + thickness_side, plus thin_wall / add_end_caps /
    remove_inside_material. set_from_face: offset (reuses the face already on
    the feature; a feature with no from-face cannot be edited here).
    set_body_array: 0-based body_indices + multi_body_cut.
    set_to_face: offset_side + distance.
    apply_direction1_treatment: treatment_type, draft_side, draft_angle,
    crown_* (raw FeaturePropertyConstants ints).
    """
    match property:
        case "get_direction1":
            return query_manager.get_direction1_extent(feature_name)
        case "set_direction1":
            return query_manager.set_direction1_extent(
                feature_name,
                _EXTENT_TYPE_CONSTANTS[extent_type],
                distance,
                _SIDE_CONSTANTS[extent_side],
            )
        case "get_direction2":
            return query_manager.get_direction2_extent(feature_name)
        case "set_direction2":
            return query_manager.set_direction2_extent(
                feature_name,
                _EXTENT_TYPE_CONSTANTS[extent_type],
                distance,
                _SIDE_CONSTANTS[extent_side],
            )
        case "get_thin_wall":
            return query_manager.get_thin_wall_options(feature_name)
        case "set_thin_wall":
            return query_manager.set_thin_wall_options(
                feature_name,
                thickness,
                _SIDE_CONSTANTS[thickness_side],
                thin_wall,
                add_end_caps,
                remove_inside_material,
            )
        case "get_from_face":
            return query_manager.get_from_face_offset(feature_name)
        case "set_from_face":
            return query_manager.set_from_face_offset(feature_name, offset)
        case "get_body_array":
            return query_manager.get_body_array(feature_name)
        case "set_body_array":
            return query_manager.set_body_array(feature_name, body_indices or [], multi_body_cut)
        case "get_to_face":
            return query_manager.get_to_face_offset(feature_name)
        case "set_to_face":
            return query_manager.set_to_face_offset(
                feature_name, _OFFSET_SIDE_CONSTANTS[offset_side], distance
            )
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
    action: Literal["rename", "suppress", "unsuppress", "set_mode"],
    feature_name: str = "",
    new_name: str = "",
    mode: Literal["ordered", "synchronous"] = "ordered",
) -> dict[str, Any]:
    """Rename, suppress or unsuppress a feature by name, or set the modeling mode.

    set_mode: mode ordered/synchronous (ignores feature_name).
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
    property: Literal["endpoints", "length", "tangent", "geometry", "curvature", "vertex"],
    face_index: int = 0,
    edge_index: int = 0,
    param: float = 0.5,
    which: Literal["start", "end"] = "start",
) -> dict[str, Any]:
    """Read edge data (read-only). 0-based face_index and edge_index within that face.

    tangent/curvature: at param 0.0-1.0 along the edge. vertex: which start/end.
    Lengths/points in meters.
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
    property: Literal["normal", "geometry", "loops", "curvature"],
    face_index: int = 0,
    u: float = 0.5,
    v: float = 0.5,
) -> dict[str, Any]:
    """Read face data (read-only). 0-based face_index.

    normal/curvature: evaluated at parametric (u, v) in 0.0-1.0.
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
    property: Literal[
        "extreme_point",
        "faces_by_ray",
        "shells",
        "vertices",
        "shell_info",
        "point_inside",
        "user_physical_properties",
        "facet_data",
        "faces",
        "edges",
        "spatial_context",
    ],
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
    offset: int = 0,
    limit: int = DEFAULT_PAGE_LIMIT,
) -> dict[str, Any]:
    """Read body-level topology (read-only). Meters.

    extreme_point: farthest point along direction_x/y/z.
    faces_by_ray: ray from origin_x/y/z along direction_x/y/z.
    shell_info: 0-based shell_index. point_inside: x,y,z.
    facet_data: tessellate at tolerance (0 = default).
    spatial_context: body count, bounding box with center, whether the body
    sits on the origin, the open sketch's plane, and the plane-to-axis map.

    faces / edges / vertices / shells are PAGED: they walk the collection one
    COM call per entity, so they return at most `limit` items starting at the
    0-based `offset` (default 200, hard ceiling 2000). The reply is
    {total, offset, limit, items, truncated}; keep raising `offset` by `limit`
    while `truncated` is true.
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
        case "faces":
            return query_manager.get_body_faces(offset, limit)
        case "edges":
            return query_manager.get_body_edges(offset, limit)
        case "shells":
            return query_manager.get_body_shells(offset, limit)
        case "vertices":
            return query_manager.get_body_vertices(offset, limit)
        case "spatial_context":
            return query_manager.get_spatial_context()
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
    type: Literal["curve", "surface"],
    face_index: int = 0,
    edge_index: int = 0,
) -> dict[str, Any]:
    """Read NURBS metadata (read-only). curve: 0-based face_index + edge_index.

    surface: 0-based face_index.
    """
    match type:
        case "curve":
            return query_manager.get_bspline_curve_info(face_index, edge_index)
        case "surface":
            return query_manager.get_bspline_surface_info(face_index)
        case _:
            return {"error": f"Unknown type: {type}"}


# ── Composite: recompute ──────────────────────────────────────────


def recompute(scope: Literal["model", "document"] = "model") -> dict[str, Any]:
    """Recompute the active model (feature tree) or the whole document."""
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
    tags = {"query"}
    register_tool(mcp, measure, tags=tags, read_only=True)
    register_tool(mcp, manage_variable, tags=tags)
    register_tool(mcp, manage_property, tags=tags, destructive=True)
    register_tool(mcp, manage_material, tags=tags)
    register_tool(mcp, set_appearance, tags=tags, idempotent=True)
    register_tool(mcp, manage_layer, tags=tags, destructive=True)
    register_tool(mcp, select_set, tags=tags, destructive=True)
    register_tool(mcp, edit_feature_extent, tags=tags)
    register_tool(mcp, manage_feature_tree, tags=tags)
    register_tool(mcp, query_edge, tags=tags, read_only=True)
    register_tool(mcp, query_face, tags=tags, read_only=True)
    register_tool(mcp, query_body, tags=tags, read_only=True)
    register_tool(mcp, query_bspline, tags=tags, read_only=True)
    register_tool(mcp, recompute, tags=tags, idempotent=True)
