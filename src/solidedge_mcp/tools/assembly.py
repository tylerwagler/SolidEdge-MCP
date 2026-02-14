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


def assembly_pattern_component(
    component_index: int, count: int, spacing: float, direction: str = "X"
) -> dict:
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


def assembly_check_interference(component_index: int | None = None) -> dict:
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


def assembly_occurrence_rotate(
    component_index: int,
    axis_x1: float,
    axis_y1: float,
    axis_z1: float,
    axis_x2: float,
    axis_y2: float,
    axis_z2: float,
    angle: float,
) -> dict:
    """Rotate a component around an axis (degrees)."""
    return assembly_manager.occurrence_rotate(
        component_index, axis_x1, axis_y1, axis_z1, axis_x2, axis_y2, axis_z2, angle
    )


def assembly_set_component_transform(
    component_index: int,
    origin_x: float,
    origin_y: float,
    origin_z: float,
    angle_x: float,
    angle_y: float,
    angle_z: float,
) -> dict:
    """Set a component's full transform (position + rotation). Angles in degrees."""
    return assembly_manager.set_component_transform(
        component_index, origin_x, origin_y, origin_z, angle_x, angle_y, angle_z
    )


def assembly_set_component_origin(component_index: int, x: float, y: float, z: float) -> dict:
    """Set a component's origin (position only)."""
    return assembly_manager.set_component_origin(component_index, x, y, z)


def assembly_mirror_component(component_index: int, plane_index: int) -> dict:
    """Mirror a component across a reference plane (1-based plane index)."""
    return assembly_manager.mirror_component(component_index, plane_index)


def assembly_get_occurrence_count() -> dict:
    """Get the count of top-level occurrences."""
    return assembly_manager.get_occurrence_count()


def assembly_is_subassembly(component_index: int) -> dict:
    """Check if a component is a subassembly (vs a part)."""
    return assembly_manager.is_subassembly(component_index)


def assembly_get_component_display_name(component_index: int) -> dict:
    """Get the display name of a component (user-visible label in assembly tree)."""
    return assembly_manager.get_component_display_name(component_index)


def assembly_get_occurrence_document(component_index: int) -> dict:
    """Get document info for a component's source file."""
    return assembly_manager.get_occurrence_document(component_index)


def assembly_get_sub_occurrences(component_index: int) -> dict:
    """Get sub-occurrences (children) of a component."""
    return assembly_manager.get_sub_occurrences(component_index)


def assembly_add_component_with_transform(
    file_path: str,
    origin_x: float = 0,
    origin_y: float = 0,
    origin_z: float = 0,
    angle_x: float = 0,
    angle_y: float = 0,
    angle_z: float = 0,
) -> dict:
    """Place a component with position and Euler rotation angles (degrees)."""
    return assembly_manager.add_component_with_transform(
        file_path, origin_x, origin_y, origin_z, angle_x, angle_y, angle_z
    )


def assembly_delete_relation(relation_index: int) -> dict:
    """Delete an assembly relation/constraint by 0-based index."""
    return assembly_manager.delete_relation(relation_index)


def assembly_get_relation_info(relation_index: int) -> dict:
    """Get detailed info about a specific assembly relation by 0-based index."""
    return assembly_manager.get_relation_info(relation_index)


def assembly_add_planar_relation(
    occurrence1_index: int,
    occurrence2_index: int,
    offset: float = 0.0,
    orientation: str = "Align",
) -> dict:
    """Add a planar relation between two components.

    orientation: 'Align', 'Antialign', 'NotSpecified'.
    """
    return assembly_manager.add_planar_relation(
        occurrence1_index, occurrence2_index, offset, orientation
    )


def assembly_add_axial_relation(
    occurrence1_index: int,
    occurrence2_index: int,
    orientation: str = "Align",
) -> dict:
    """Add an axial relation between two components.

    orientation: 'Align', 'Antialign', 'NotSpecified'.
    """
    return assembly_manager.add_axial_relation(occurrence1_index, occurrence2_index, orientation)


def assembly_add_angular_relation(
    occurrence1_index: int,
    occurrence2_index: int,
    angle: float = 0.0,
) -> dict:
    """Add an angular relation between two assembly components. Angle in degrees."""
    return assembly_manager.add_angular_relation(occurrence1_index, occurrence2_index, angle)


def assembly_add_point_relation(
    occurrence1_index: int,
    occurrence2_index: int,
) -> dict:
    """Add a point (connect) relation between two assembly components."""
    return assembly_manager.add_point_relation(occurrence1_index, occurrence2_index)


def assembly_add_tangent_relation(
    occurrence1_index: int,
    occurrence2_index: int,
) -> dict:
    """Add a tangent relation between two assembly components."""
    return assembly_manager.add_tangent_relation(occurrence1_index, occurrence2_index)


def assembly_add_gear_relation(
    occurrence1_index: int,
    occurrence2_index: int,
    ratio1: float = 1.0,
    ratio2: float = 1.0,
) -> dict:
    """Add a gear relation between two assembly components with gear ratios."""
    return assembly_manager.add_gear_relation(occurrence1_index, occurrence2_index, ratio1, ratio2)


def assembly_get_relation_offset(relation_index: int) -> dict:
    """Get the offset value from a planar relation (0-based index)."""
    return assembly_manager.get_relation_offset(relation_index)


def assembly_set_relation_offset(relation_index: int, offset: float) -> dict:
    """Set the offset value on a planar relation (0-based index). Offset in meters."""
    return assembly_manager.set_relation_offset(relation_index, offset)


def assembly_get_relation_angle(relation_index: int) -> dict:
    """Get the angle value from an angular relation (0-based index). Returns degrees."""
    return assembly_manager.get_relation_angle(relation_index)


def assembly_set_relation_angle(relation_index: int, angle: float) -> dict:
    """Set the angle on an angular relation (0-based index). Angle in degrees."""
    return assembly_manager.set_relation_angle(relation_index, angle)


def assembly_get_normals_aligned(relation_index: int) -> dict:
    """Get the NormalsAligned boolean from a relation (0-based index)."""
    return assembly_manager.get_normals_aligned(relation_index)


def assembly_set_normals_aligned(relation_index: int, aligned: bool) -> dict:
    """Set the NormalsAligned boolean on a relation (0-based index)."""
    return assembly_manager.set_normals_aligned(relation_index, aligned)


def assembly_suppress_relation(relation_index: int) -> dict:
    """Suppress an assembly relation (0-based index)."""
    return assembly_manager.suppress_relation(relation_index)


def assembly_unsuppress_relation(relation_index: int) -> dict:
    """Unsuppress an assembly relation (0-based index)."""
    return assembly_manager.unsuppress_relation(relation_index)


def assembly_get_relation_geometry(relation_index: int) -> dict:
    """Get geometry info (occurrence references) from a relation (0-based index)."""
    return assembly_manager.get_relation_geometry(relation_index)


def assembly_get_gear_ratio(relation_index: int) -> dict:
    """Get gear ratio values (RatioValue1, RatioValue2) from a gear relation (0-based index)."""
    return assembly_manager.get_gear_ratio(relation_index)


# ============================================================================
# BATCH 8: ASSEMBLY OCCURRENCES & PROPERTIES
# ============================================================================


def assembly_add_family_member(
    file_path: str, family_member_name: str, x: float = 0, y: float = 0, z: float = 0
) -> dict:
    """Add a Family of Parts member to the assembly."""
    return assembly_manager.add_family_member(file_path, family_member_name, x, y, z)


def assembly_add_family_with_transform(
    file_path: str,
    family_member_name: str,
    origin_x: float = 0,
    origin_y: float = 0,
    origin_z: float = 0,
    angle_x: float = 0,
    angle_y: float = 0,
    angle_z: float = 0,
) -> dict:
    """Add a Family of Parts member with position and rotation (degrees)."""
    return assembly_manager.add_family_with_transform(
        file_path, family_member_name, origin_x, origin_y, origin_z, angle_x, angle_y, angle_z
    )


def assembly_add_by_template(file_path: str, template_name: str) -> dict:
    """Add a component to the assembly using a template."""
    return assembly_manager.add_by_template(file_path, template_name)


def assembly_add_adjustable_part(file_path: str, x: float = 0, y: float = 0, z: float = 0) -> dict:
    """Add a part as an adjustable part to the assembly."""
    return assembly_manager.add_adjustable_part(file_path, x, y, z)


def assembly_reorder_occurrence(component_index: int, target_index: int) -> dict:
    """Reorder an occurrence in the assembly tree (0-based indices)."""
    return assembly_manager.reorder_occurrence(component_index, target_index)


def assembly_put_transform_euler(
    component_index: int,
    x: float,
    y: float,
    z: float,
    rx: float,
    ry: float,
    rz: float,
) -> dict:
    """Set a component's transform using Euler angles (degrees)."""
    return assembly_manager.put_transform_euler(component_index, x, y, z, rx, ry, rz)


def assembly_put_origin(component_index: int, x: float, y: float, z: float) -> dict:
    """Set a component's origin position (meters)."""
    return assembly_manager.put_origin(component_index, x, y, z)


def assembly_make_writable(component_index: int) -> dict:
    """Make a component writable (editable) in the assembly."""
    return assembly_manager.make_writable(component_index)


def assembly_swap_family_member(component_index: int, new_member_name: str) -> dict:
    """Swap a Family of Parts occurrence for a different family member."""
    return assembly_manager.swap_family_member(component_index, new_member_name)


def assembly_get_occurrence_bodies(component_index: int) -> dict:
    """Get body information from a specific component occurrence."""
    return assembly_manager.get_occurrence_bodies(component_index)


def assembly_get_occurrence_style(component_index: int) -> dict:
    """Get the style (appearance) of a component occurrence."""
    return assembly_manager.get_occurrence_style(component_index)


def assembly_is_tube(component_index: int) -> dict:
    """Check if a component occurrence is a tube."""
    return assembly_manager.is_tube(component_index)


def assembly_get_adjustable_part(component_index: int) -> dict:
    """Get adjustable part info from a component occurrence."""
    return assembly_manager.get_adjustable_part(component_index)


def assembly_get_face_style(component_index: int) -> dict:
    """Get the face style of a component occurrence."""
    return assembly_manager.get_face_style(component_index)


def assembly_add_family_with_matrix(
    family_file_path: str, member_name: str, matrix: list[float]
) -> dict:
    """Add a Family of Parts member with a 4x4 transformation matrix (16 floats)."""
    return assembly_manager.add_family_with_matrix(family_file_path, member_name, matrix)


def assembly_get_occurrence(internal_id: int) -> dict:
    """Get an occurrence by its internal ID (not by index)."""
    return assembly_manager.get_occurrence(internal_id)


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
    mcp.tool()(assembly_set_component_transform)
    mcp.tool()(assembly_set_component_origin)
    mcp.tool()(assembly_mirror_component)
    mcp.tool()(assembly_get_occurrence_count)
    mcp.tool()(assembly_is_subassembly)
    mcp.tool()(assembly_get_component_display_name)
    mcp.tool()(assembly_get_occurrence_document)
    mcp.tool()(assembly_get_sub_occurrences)
    mcp.tool()(assembly_add_component_with_transform)
    mcp.tool()(assembly_delete_relation)
    mcp.tool()(assembly_get_relation_info)
    mcp.tool()(assembly_add_planar_relation)
    mcp.tool()(assembly_add_axial_relation)
    mcp.tool()(assembly_add_angular_relation)
    mcp.tool()(assembly_add_point_relation)
    mcp.tool()(assembly_add_tangent_relation)
    mcp.tool()(assembly_add_gear_relation)
    mcp.tool()(assembly_get_relation_offset)
    mcp.tool()(assembly_set_relation_offset)
    mcp.tool()(assembly_get_relation_angle)
    mcp.tool()(assembly_set_relation_angle)
    mcp.tool()(assembly_get_normals_aligned)
    mcp.tool()(assembly_set_normals_aligned)
    mcp.tool()(assembly_suppress_relation)
    mcp.tool()(assembly_unsuppress_relation)
    mcp.tool()(assembly_get_relation_geometry)
    mcp.tool()(assembly_get_gear_ratio)
    # Batch 8: Assembly Occurrences & Properties
    mcp.tool()(assembly_add_family_member)
    mcp.tool()(assembly_add_family_with_transform)
    mcp.tool()(assembly_add_by_template)
    mcp.tool()(assembly_add_adjustable_part)
    mcp.tool()(assembly_reorder_occurrence)
    mcp.tool()(assembly_put_transform_euler)
    mcp.tool()(assembly_put_origin)
    mcp.tool()(assembly_make_writable)
    mcp.tool()(assembly_swap_family_member)
    mcp.tool()(assembly_get_occurrence_bodies)
    mcp.tool()(assembly_get_occurrence_style)
    mcp.tool()(assembly_is_tube)
    mcp.tool()(assembly_get_adjustable_part)
    mcp.tool()(assembly_get_face_style)
    mcp.tool()(assembly_add_family_with_matrix)
    mcp.tool()(assembly_get_occurrence)
