"""Query and inspection tools for Solid Edge MCP."""

from solidedge_mcp.managers import query_manager

# === Measurement ===


def measure_distance(x1: float, y1: float, z1: float, x2: float, y2: float, z2: float) -> dict:
    """Measure distance between two 3D points."""
    return query_manager.measure_distance(x1, y1, z1, x2, y2, z2)


def measure_angle(
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
    """Measure angle between three 3D points (vertex at point 2)."""
    return query_manager.measure_angle(x1, y1, z1, x2, y2, z2, x3, y3, z3)


# === Mass / Physical Properties ===


def get_mass_properties(density: float = 7850) -> dict:
    """Get mass properties (density default: 7850 kg/m3 steel)."""
    return query_manager.get_mass_properties(density)


def get_center_of_gravity() -> dict:
    """Get the center of gravity of the active part."""
    return query_manager.get_center_of_gravity()


def get_moments_of_inertia() -> dict:
    """Get moments of inertia of the active part."""
    return query_manager.get_moments_of_inertia()


def get_surface_area() -> dict:
    """Get the total surface area of the active part."""
    return query_manager.get_surface_area()


def get_volume() -> dict:
    """Get the volume of the active part."""
    return query_manager.get_volume()


def set_material_density(density: float) -> dict:
    """Set the material density for mass property calculations."""
    return query_manager.set_material_density(density)


# === Geometry / Topology ===


def get_bounding_box() -> dict:
    """Get the bounding box of the model."""
    return query_manager.get_bounding_box()


def get_body_faces() -> dict:
    """Get all faces on the body with geometry info."""
    return query_manager.get_body_faces()


def get_body_edges() -> dict:
    """Get edge information from the model body."""
    return query_manager.get_body_edges()


def get_face_info(face_index: int) -> dict:
    """Get detailed information about a specific face."""
    return query_manager.get_face_info(face_index)


def get_face_count() -> dict:
    """Get the total face count on the body."""
    return query_manager.get_face_count()


def get_face_area(face_index: int) -> dict:
    """Get the area of a specific face."""
    return query_manager.get_face_area(face_index)


def get_edge_count() -> dict:
    """Get total edge count on the model body."""
    return query_manager.get_edge_count()


def get_edge_info(face_index: int, edge_index: int) -> dict:
    """Get detailed info about a specific edge."""
    return query_manager.get_edge_info(face_index, edge_index)


def get_solid_bodies() -> dict:
    """Report all solid bodies in the active part document."""
    return query_manager.get_solid_bodies()


def get_body_facet_data(tolerance: float = 0.0) -> dict:
    """Get tessellation/mesh data from the model body."""
    return query_manager.get_body_facet_data(tolerance)


# === Features / Design Tree ===


def get_feature_count() -> dict:
    """Get the total count of features."""
    return query_manager.get_feature_count()


def get_feature_dimensions(feature_name: str) -> dict:
    """Get dimensions/parameters of a named feature."""
    return query_manager.get_feature_dimensions(feature_name)


def get_design_edgebar_features() -> dict:
    """Get the full feature tree from DesignEdgebarFeatures."""
    return query_manager.get_design_edgebar_features()


def rename_feature(old_name: str, new_name: str) -> dict:
    """Rename a feature in the design tree (by name)."""
    return query_manager.rename_feature(old_name, new_name)


def suppress_feature(feature_name: str) -> dict:
    """Suppress a feature by name."""
    return query_manager.suppress_feature(feature_name)


def unsuppress_feature(feature_name: str) -> dict:
    """Unsuppress a feature by name."""
    return query_manager.unsuppress_feature(feature_name)


# === Document Properties ===


def get_document_properties() -> dict:
    """Get document properties (Title, Subject, Author, etc.)."""
    return query_manager.get_document_properties()


def set_document_property(name: str, value: str) -> dict:
    """Set a summary/document property."""
    return query_manager.set_document_property(name, value)


def get_ref_planes() -> dict:
    """List all reference planes."""
    return query_manager.get_ref_planes()


# === Variables ===


def get_variables() -> dict:
    """Get all variables (dimensions, parameters) in the document."""
    return query_manager.get_variables()


def get_variable(name: str) -> dict:
    """Get a specific variable by name."""
    return query_manager.get_variable(name)


def set_variable(name: str, value: float) -> dict:
    """Set a variable's value."""
    return query_manager.set_variable(name, value)


def add_variable(name: str, formula: str, units_type: str = None) -> dict:
    """Add a new user variable."""
    return query_manager.add_variable(name, formula, units_type)


def query_variables(pattern: str = "*", case_insensitive: bool = True) -> dict:
    """Query variables matching a pattern."""
    return query_manager.query_variables(pattern, case_insensitive)


def get_variable_formula(name: str) -> dict:
    """Get the formula of a variable by name."""
    return query_manager.get_variable_formula(name)


def rename_variable(old_name: str, new_name: str) -> dict:
    """Rename a variable's display name."""
    return query_manager.rename_variable(old_name, new_name)


def get_variable_names(name: str) -> dict:
    """Get both DisplayName and SystemName of a variable."""
    return query_manager.get_variable_names(name)


def translate_variable(name: str) -> dict:
    """Translate (look up) a variable by name via Variables.Translate()."""
    return query_manager.translate_variable(name)


def copy_variable_to_clipboard(name: str) -> dict:
    """Copy a variable definition to the clipboard for pasting into another document."""
    return query_manager.copy_variable_to_clipboard(name)


def add_variable_from_clipboard(name: str, units_type: str = None) -> dict:
    """Add a variable from the clipboard (previously copied with copy_variable_to_clipboard)."""
    return query_manager.add_variable_from_clipboard(name, units_type)


# === Custom Properties ===


def get_custom_properties() -> dict:
    """Get all custom properties."""
    return query_manager.get_custom_properties()


def set_custom_property(name: str, value: str) -> dict:
    """Set or create a custom property."""
    return query_manager.set_custom_property(name, value)


def delete_custom_property(name: str) -> dict:
    """Delete a custom property."""
    return query_manager.delete_custom_property(name)


# === Materials ===


def get_material_list() -> dict:
    """Get the list of available materials."""
    return query_manager.get_material_list()


def get_material_table() -> dict:
    """Get the full material table with properties."""
    return query_manager.get_material_table()


def get_material_property(material_name: str, property_index: int) -> dict:
    """Get a specific property of a material."""
    return query_manager.get_material_property(material_name, property_index)


def set_material(material_name: str) -> dict:
    """Assign a material to the active part."""
    return query_manager.set_material(material_name)


# === Appearance ===


def set_body_color(red: int, green: int, blue: int) -> dict:
    """Set the body color of the active part."""
    return query_manager.set_body_color(red, green, blue)


def get_body_color() -> dict:
    """Get the current body color of the active part."""
    return query_manager.get_body_color()


def set_face_color(face_index: int, red: int, green: int, blue: int) -> dict:
    """Set the color of a specific face."""
    return query_manager.set_face_color(face_index, red, green, blue)


def set_body_opacity(opacity: float) -> dict:
    """Set the body opacity (0.0 = fully transparent, 1.0 = fully opaque)."""
    return query_manager.set_body_opacity(opacity)


def set_body_reflectivity(reflectivity: float) -> dict:
    """Set the body reflectivity (0.0 to 1.0)."""
    return query_manager.set_body_reflectivity(reflectivity)


# === Layers ===


def get_layers() -> dict:
    """Get all layers in the active document."""
    return query_manager.get_layers()


def add_layer(name: str) -> dict:
    """Add a new layer to the active document."""
    return query_manager.add_layer(name)


def activate_layer(name_or_index) -> dict:
    """Activate a layer by name (str) or 0-based index (int)."""
    return query_manager.activate_layer(name_or_index)


def set_layer_properties(name_or_index, show: bool = None, selectable: bool = None) -> dict:
    """Set layer visibility and selectability properties."""
    return query_manager.set_layer_properties(name_or_index, show, selectable)


# === Modeling Mode ===


def get_modeling_mode() -> dict:
    """Get the current modeling mode (Ordered vs Synchronous)."""
    return query_manager.get_modeling_mode()


def set_modeling_mode(mode: str) -> dict:
    """Set the modeling mode ('ordered' or 'synchronous')."""
    return query_manager.set_modeling_mode(mode)


# === Select Set ===


def get_select_set() -> dict:
    """Get the current selection set."""
    return query_manager.get_select_set()


def clear_select_set() -> dict:
    """Clear the current selection set."""
    return query_manager.clear_select_set()


def select_add(object_type: str, index: int) -> dict:
    """Add an object to the selection set."""
    return query_manager.select_add(object_type, index)


def select_remove(index: int) -> dict:
    """Remove an object from the selection set by 0-based index."""
    return query_manager.select_remove(index)


def select_all() -> dict:
    """Select all objects in the active document."""
    return query_manager.select_all()


def select_copy() -> dict:
    """Copy the current selection to the clipboard."""
    return query_manager.select_copy()


def select_cut() -> dict:
    """Cut the current selection to the clipboard."""
    return query_manager.select_cut()


def select_delete() -> dict:
    """Delete the currently selected objects."""
    return query_manager.select_delete()


def select_suspend_display() -> dict:
    """Suspend display updates for the selection set (for batch perf)."""
    return query_manager.select_suspend_display()


def select_resume_display() -> dict:
    """Resume display updates for the selection set."""
    return query_manager.select_resume_display()


def select_refresh_display() -> dict:
    """Refresh the display of the selection set."""
    return query_manager.select_refresh_display()


# === Feature Analysis ===


def get_feature_status(feature_name: str) -> dict:
    """Get the status of a feature (OK, suppressed, failed, etc.)."""
    return query_manager.get_feature_status(feature_name)


def get_feature_profiles(feature_name: str) -> dict:
    """Get the sketch profiles associated with a feature."""
    return query_manager.get_feature_profiles(feature_name)


def get_vertex_count() -> dict:
    """Get the total vertex count on the model body."""
    return query_manager.get_vertex_count()


# === Recompute ===


def recompute() -> dict:
    """Recompute the active document and model."""
    return query_manager.recompute()


# === Variable Formulas ===


def set_variable_formula(name: str, formula: str) -> dict:
    """Set the formula of an existing variable by name."""
    return query_manager.set_variable_formula(name, formula)


# === Layer Management ===


def delete_layer(name_or_index) -> dict:
    """Delete a layer by name (str) or 0-based index (int)."""
    return query_manager.delete_layer(name_or_index)


# === Feature Parents ===


def get_feature_parents(feature_name: str) -> dict:
    """Get parent geometry/features of a named feature."""
    return query_manager.get_feature_parents(feature_name)


# === Feature Editing ===


def get_direction1_extent(feature_name: str) -> dict:
    """Get Direction 1 extent parameters (type, distance) from a feature."""
    return query_manager.get_direction1_extent(feature_name)


def set_direction1_extent(feature_name: str, extent_type: int, distance: float = 0.0) -> dict:
    """Set Direction 1 extent on a feature (13=Finite, 16=ThroughAll, 44=None)."""
    return query_manager.set_direction1_extent(feature_name, extent_type, distance)


def get_direction2_extent(feature_name: str) -> dict:
    """Get Direction 2 extent parameters (type, distance) from a feature."""
    return query_manager.get_direction2_extent(feature_name)


def set_direction2_extent(feature_name: str, extent_type: int, distance: float = 0.0) -> dict:
    """Set Direction 2 extent on a feature (13=Finite, 16=ThroughAll, 44=None)."""
    return query_manager.set_direction2_extent(feature_name, extent_type, distance)


def get_thin_wall_options(feature_name: str) -> dict:
    """Get thin wall options (type, thicknesses) from a feature."""
    return query_manager.get_thin_wall_options(feature_name)


def set_thin_wall_options(
    feature_name: str, wall_type: int, thickness1: float, thickness2: float = 0.0
) -> dict:
    """Set thin wall options on a feature."""
    return query_manager.set_thin_wall_options(feature_name, wall_type, thickness1, thickness2)


def get_from_face_offset(feature_name: str) -> dict:
    """Get the 'from face' offset data from a feature."""
    return query_manager.get_from_face_offset(feature_name)


def set_from_face_offset(feature_name: str, offset: float) -> dict:
    """Set the 'from face' offset distance on a feature."""
    return query_manager.set_from_face_offset(feature_name, offset)


def get_body_array(feature_name: str) -> dict:
    """Get the body array (multi-body references) from a feature."""
    return query_manager.get_body_array(feature_name)


def set_body_array(feature_name: str, body_indices: list[int]) -> dict:
    """Set the body array on a feature using 0-based model indices."""
    return query_manager.set_body_array(feature_name, body_indices)


# === Material Library (Batch 10) ===


def get_material_library() -> dict:
    """Get the full material library with names, densities, and properties."""
    return query_manager.get_material_library()


def set_material_by_name(material_name: str) -> dict:
    """Look up a material by name and apply it to the active document."""
    return query_manager.set_material_by_name(material_name)


# === B-Rep Topology Queries ===


def get_edge_endpoints(face_index: int, edge_index: int) -> dict:
    """Get start/end XYZ coordinates of an edge on a face."""
    return query_manager.get_edge_endpoints(face_index, edge_index)


def get_edge_length(face_index: int, edge_index: int) -> dict:
    """Get the total length of an edge on a face."""
    return query_manager.get_edge_length(face_index, edge_index)


def get_edge_tangent(face_index: int, edge_index: int, param: float = 0.5) -> dict:
    """Get tangent vector at a parameter on an edge."""
    return query_manager.get_edge_tangent(face_index, edge_index, param)


def get_edge_geometry(face_index: int, edge_index: int) -> dict:
    """Get underlying geometry type and data of an edge (Line/Circle/Ellipse/BSpline)."""
    return query_manager.get_edge_geometry(face_index, edge_index)


def get_face_normal(face_index: int, u: float = 0.5, v: float = 0.5) -> dict:
    """Get normal vector at a parametric point on a face."""
    return query_manager.get_face_normal(face_index, u, v)


def get_face_geometry(face_index: int) -> dict:
    """Get geometry type and data of a face (Plane/Cylinder/Cone/etc)."""
    return query_manager.get_face_geometry(face_index)


def get_face_loops(face_index: int) -> dict:
    """Get loop info for a face (outer boundary vs holes)."""
    return query_manager.get_face_loops(face_index)


def get_face_curvature(face_index: int, u: float = 0.5, v: float = 0.5) -> dict:
    """Get principal curvatures at a parametric point on a face."""
    return query_manager.get_face_curvature(face_index, u, v)


def get_vertex_point(face_index: int, edge_index: int, which: str = "start") -> dict:
    """Get XYZ coordinates of a vertex (start or end) on an edge."""
    return query_manager.get_vertex_point(face_index, edge_index, which)


def get_body_extreme_point(direction_x: float, direction_y: float, direction_z: float) -> dict:
    """Get extreme point of body in a direction vector."""
    return query_manager.get_body_extreme_point(direction_x, direction_y, direction_z)


def get_faces_by_ray(
    origin_x: float,
    origin_y: float,
    origin_z: float,
    direction_x: float,
    direction_y: float,
    direction_z: float,
) -> dict:
    """Ray-cast query to find faces hit by a ray."""
    return query_manager.get_faces_by_ray(
        origin_x, origin_y, origin_z, direction_x, direction_y, direction_z
    )


def get_shell_info(shell_index: int = 0) -> dict:
    """Get shell topology info (is_closed, volume, face/edge counts)."""
    return query_manager.get_shell_info(shell_index)


def is_point_inside_body(x: float, y: float, z: float) -> dict:
    """Test if a 3D point is inside the solid body."""
    return query_manager.is_point_inside_body(x, y, z)


def get_body_shells() -> dict:
    """List all shells in the body with is_closed and volume."""
    return query_manager.get_body_shells()


def get_body_vertices() -> dict:
    """Get all vertices of the body with XYZ coordinates."""
    return query_manager.get_body_vertices()


def get_bspline_curve_info(face_index: int, edge_index: int) -> dict:
    """Get NURBS curve metadata (order, poles, knots, flags) from an edge."""
    return query_manager.get_bspline_curve_info(face_index, edge_index)


def get_bspline_surface_info(face_index: int) -> dict:
    """Get NURBS surface metadata (order, poles, knots, flags) from a face."""
    return query_manager.get_bspline_surface_info(face_index)


def get_edge_curvature(face_index: int, edge_index: int, param: float = 0.5) -> dict:
    """Get curvature at a parameter on an edge."""
    return query_manager.get_edge_curvature(face_index, edge_index, param)


# === Registration ===


def register(mcp):
    """Register query tools with the MCP server."""
    # Measurement
    mcp.tool()(measure_distance)
    mcp.tool()(measure_angle)
    # Mass / Physical Properties
    mcp.tool()(get_mass_properties)
    mcp.tool()(get_center_of_gravity)
    mcp.tool()(get_moments_of_inertia)
    mcp.tool()(get_surface_area)
    mcp.tool()(get_volume)
    mcp.tool()(set_material_density)
    # Geometry / Topology
    mcp.tool()(get_bounding_box)
    mcp.tool()(get_body_faces)
    mcp.tool()(get_body_edges)
    mcp.tool()(get_face_info)
    mcp.tool()(get_face_count)
    mcp.tool()(get_face_area)
    mcp.tool()(get_edge_count)
    mcp.tool()(get_edge_info)
    mcp.tool()(get_solid_bodies)
    mcp.tool()(get_body_facet_data)
    # Features / Design Tree
    mcp.tool()(get_feature_count)
    mcp.tool()(get_feature_dimensions)
    mcp.tool()(get_design_edgebar_features)
    mcp.tool()(rename_feature)
    mcp.tool()(suppress_feature)
    mcp.tool()(unsuppress_feature)
    # Document Properties
    mcp.tool()(get_document_properties)
    mcp.tool()(set_document_property)
    mcp.tool()(get_ref_planes)
    # Variables
    mcp.tool()(get_variables)
    mcp.tool()(get_variable)
    mcp.tool()(set_variable)
    mcp.tool()(add_variable)
    mcp.tool()(query_variables)
    mcp.tool()(get_variable_formula)
    mcp.tool()(rename_variable)
    mcp.tool()(get_variable_names)
    mcp.tool()(translate_variable)
    mcp.tool()(copy_variable_to_clipboard)
    mcp.tool()(add_variable_from_clipboard)
    # Custom Properties
    mcp.tool()(get_custom_properties)
    mcp.tool()(set_custom_property)
    mcp.tool()(delete_custom_property)
    # Materials
    mcp.tool()(get_material_list)
    mcp.tool()(get_material_table)
    mcp.tool()(get_material_property)
    mcp.tool()(set_material)
    # Appearance
    mcp.tool()(set_body_color)
    mcp.tool()(get_body_color)
    mcp.tool()(set_face_color)
    mcp.tool()(set_body_opacity)
    mcp.tool()(set_body_reflectivity)
    # Layers
    mcp.tool()(get_layers)
    mcp.tool()(add_layer)
    mcp.tool()(activate_layer)
    mcp.tool()(set_layer_properties)
    # Modeling Mode
    mcp.tool()(get_modeling_mode)
    mcp.tool()(set_modeling_mode)
    # Select Set
    mcp.tool()(get_select_set)
    mcp.tool()(clear_select_set)
    mcp.tool()(select_add)
    mcp.tool()(select_remove)
    mcp.tool()(select_all)
    mcp.tool()(select_copy)
    mcp.tool()(select_cut)
    mcp.tool()(select_delete)
    mcp.tool()(select_suspend_display)
    mcp.tool()(select_resume_display)
    mcp.tool()(select_refresh_display)
    # Feature Analysis
    mcp.tool()(get_feature_status)
    mcp.tool()(get_feature_profiles)
    mcp.tool()(get_vertex_count)
    # Recompute
    mcp.tool()(recompute)
    # Variable Formulas
    mcp.tool()(set_variable_formula)
    # Layer Management
    mcp.tool()(delete_layer)
    # Feature Parents
    mcp.tool()(get_feature_parents)
    # Feature Editing
    mcp.tool()(get_direction1_extent)
    mcp.tool()(set_direction1_extent)
    mcp.tool()(get_direction2_extent)
    mcp.tool()(set_direction2_extent)
    mcp.tool()(get_thin_wall_options)
    mcp.tool()(set_thin_wall_options)
    mcp.tool()(get_from_face_offset)
    mcp.tool()(set_from_face_offset)
    mcp.tool()(get_body_array)
    mcp.tool()(set_body_array)
    # Material Library (Batch 10)
    mcp.tool()(get_material_library)
    mcp.tool()(set_material_by_name)
    # B-Rep Topology Queries
    mcp.tool()(get_edge_endpoints)
    mcp.tool()(get_edge_length)
    mcp.tool()(get_edge_tangent)
    mcp.tool()(get_edge_geometry)
    mcp.tool()(get_face_normal)
    mcp.tool()(get_face_geometry)
    mcp.tool()(get_face_loops)
    mcp.tool()(get_face_curvature)
    mcp.tool()(get_vertex_point)
    mcp.tool()(get_body_extreme_point)
    mcp.tool()(get_faces_by_ray)
    mcp.tool()(get_shell_info)
    mcp.tool()(is_point_inside_body)
    mcp.tool()(get_body_shells)
    mcp.tool()(get_body_vertices)
    mcp.tool()(get_bspline_curve_info)
    mcp.tool()(get_bspline_surface_info)
    mcp.tool()(get_edge_curvature)
