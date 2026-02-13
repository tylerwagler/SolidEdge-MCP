"""Query and inspection tools for Solid Edge MCP."""

from solidedge_mcp.managers import query_manager


# === Measurement ===

def measure_distance(x1: float, y1: float, z1: float, x2: float, y2: float, z2: float) -> dict:
    """Measure distance between two 3D points."""
    return query_manager.measure_distance(x1, y1, z1, x2, y2, z2)

def measure_angle(x1: float, y1: float, z1: float, x2: float, y2: float, z2: float, x3: float, y3: float, z3: float) -> dict:
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


# === Recompute ===

def recompute() -> dict:
    """Recompute the active document and model."""
    return query_manager.recompute()


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
    # Modeling Mode
    mcp.tool()(get_modeling_mode)
    mcp.tool()(set_modeling_mode)
    # Select Set
    mcp.tool()(get_select_set)
    mcp.tool()(clear_select_set)
    mcp.tool()(select_add)
    # Recompute
    mcp.tool()(recompute)
