from typing import Optional

from solidedge_mcp.managers import assembly_manager

def assembly_add_component(file_path: str, x: float = 0, y: float = 0, z: float = 0) -> dict:
    """Add a component (part/assembly) to the active assembly at coordinates."""
    return assembly_manager.add_component(file_path, x, y, z)

def assembly_list_components() -> dict:
    """List all components in the active assembly."""
    return assembly_manager.list_components()

def assembly_create_mate(mate_type: str, component1_index: int, component2_index: int) -> dict:
    """Create a mate relationship. Types: 'Mate', 'PlanarAlign', 'AxialAlign'."""
    # Note: Backend implementation might require specific args logic
    return assembly_manager.create_mate(mate_type, component1_index, component2_index)

def assembly_add_align(component1_index: int, component2_index: int) -> dict:
    """Add search-based align constraint (requires UI determination of faces)."""
    return assembly_manager.add_align_constraint(component1_index, component2_index)

def assembly_add_planar_align(component1_index: int, component2_index: int) -> dict:
    """Add search-based planar align constraint."""
    return assembly_manager.add_planar_align_constraint(component1_index, component2_index)

def assembly_add_axial_align(component1_index: int, component2_index: int) -> dict:
    """Add search-based axial align constraint."""
    return assembly_manager.add_axial_align_constraint(component1_index, component2_index)

def assembly_add_angle(component1_index: int, component2_index: int, angle: float) -> dict:
    """Add angle constraint."""
    return assembly_manager.add_angle_constraint(component1_index, component2_index, angle)

def assembly_get_component_info(component_index: int) -> dict:
    """Get detailed info about a component."""
    return assembly_manager.get_component_info(component_index)

def assembly_update_component_position(component_index: int, x: float, y: float, z: float) -> dict:
    """Update a component's position."""
    return assembly_manager.update_component_position(component_index, x, y, z)

def assembly_pattern_component(component_index: int, count: int, spacing: float, direction: str = "X") -> dict:
    """Create a linear pattern of a component."""
    return assembly_manager.pattern_component(component_index, count, spacing, direction)

def assembly_suppress_component(component_index: int, suppress: bool = True) -> dict:
    """Suppress or unsuppress a component."""
    return assembly_manager.suppress_component(component_index, suppress)

def assembly_get_bounding_box(component_index: int) -> dict:
    """Get the bounding box of a component."""
    return assembly_manager.get_occurrence_bounding_box(component_index)

def assembly_get_bom() -> dict:
    """Get a Bill of Materials for the assembly."""
    return assembly_manager.get_bom()

def assembly_get_relations() -> dict:
    """Get all assembly relations/constraints."""
    return assembly_manager.get_assembly_relations()

def assembly_get_tree() -> dict:
    """Get the hierarchical component tree."""
    return assembly_manager.get_document_tree()

def assembly_set_component_visibility(component_index: int, visible: bool) -> dict:
    """Set the visibility of a component."""
    return assembly_manager.set_component_visibility(component_index, visible)

def assembly_delete_component(component_index: int) -> dict:
    """Delete a component from the assembly."""
    return assembly_manager.delete_component(component_index)

def assembly_ground_component(component_index: int, ground: bool = True) -> dict:
    """Ground (fix) or unground a component."""
    return assembly_manager.ground_component(component_index, ground)

def assembly_check_interference(component_index: Optional[int] = None) -> dict:
    """Run interference check."""
    return assembly_manager.check_interference(component_index)

def assembly_replace_component(component_index: int, new_file_path: str) -> dict:
    """Replace a component with another file."""
    return assembly_manager.replace_component(component_index, new_file_path)

def assembly_get_component_transform(component_index: int) -> dict:
    """Get the full transformation matrix of a component."""
    return assembly_manager.get_component_transform(component_index)

def assembly_get_structured_bom() -> dict:
    """Get a structured Bill of Materials (hierarchical)."""
    return assembly_manager.get_structured_bom()

def assembly_set_component_color(component_index: int, red: int, green: int, blue: int) -> dict:
    """Set the color of a component."""
    return assembly_manager.set_component_color(component_index, red, green, blue)

def assembly_occurrence_move(component_index: int, dx: float, dy: float, dz: float) -> dict:
    """Move a component by a relative delta."""
    return assembly_manager.occurrence_move(component_index, dx, dy, dz)

def assembly_occurrence_rotate(component_index: int, axis_x1: float, axis_y1: float, axis_z1: float, axis_x2: float, axis_y2: float, axis_z2: float, angle: float) -> dict:
    """Rotate a component around an axis (degrees)."""
    return assembly_manager.occurrence_rotate(component_index, axis_x1, axis_y1, axis_z1, axis_x2, axis_y2, axis_z2, angle)

def assembly_get_occurrence_count() -> dict:
    """Get the count of top-level occurrences."""
    return assembly_manager.get_occurrence_count()

def register(mcp):
    """Register assembly tools with the MCP server."""
    mcp.tool()(assembly_add_component)
    mcp.tool()(assembly_list_components)
    mcp.tool()(assembly_create_mate)
    mcp.tool()(assembly_add_align)
    mcp.tool()(assembly_add_planar_align)
    mcp.tool()(assembly_add_axial_align)
    mcp.tool()(assembly_add_angle)
    mcp.tool()(assembly_get_component_info)
    mcp.tool()(assembly_update_component_position)
    mcp.tool()(assembly_pattern_component)
    mcp.tool()(assembly_suppress_component)
    mcp.tool()(assembly_get_bounding_box)
    mcp.tool()(assembly_get_bom)
    mcp.tool()(assembly_get_relations)
    mcp.tool()(assembly_get_tree)
    mcp.tool()(assembly_set_component_visibility)
    mcp.tool()(assembly_delete_component)
    mcp.tool()(assembly_ground_component)
    mcp.tool()(assembly_check_interference)
    mcp.tool()(assembly_replace_component)
    mcp.tool()(assembly_get_component_transform)
    mcp.tool()(assembly_get_structured_bom)
    mcp.tool()(assembly_set_component_color)
    mcp.tool()(assembly_occurrence_move)
    mcp.tool()(assembly_occurrence_rotate)
    mcp.tool()(assembly_get_occurrence_count)
