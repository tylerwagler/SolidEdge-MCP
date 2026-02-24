from solidedge_mcp.backends.validation import validate_numerics, validate_path
from solidedge_mcp.managers import assembly_manager

# ================================================================
# Group 72: add_assembly_component (7 -> 1)
# ================================================================


def add_assembly_component(
    method: str = "basic",
    file_path: str = "",
    x: float = 0,
    y: float = 0,
    z: float = 0,
    origin_x: float = 0,
    origin_y: float = 0,
    origin_z: float = 0,
    angle_x: float = 0,
    angle_y: float = 0,
    angle_z: float = 0,
    family_member_name: str = "",
    template_name: str = "",
    matrix: list[float] | None = None,
    segment_indices: list[int] | None = None,
) -> dict:
    """Add a component to the active assembly.

    method: 'basic' | 'with_transform' | 'family'
      | 'family_with_transform' | 'family_with_matrix'
      | 'by_template' | 'adjustable' | 'tube'

    Positions in meters. Angles in degrees. Matrix is 16-element 4x4.
    """
    if file_path:
        file_path, err = validate_path(file_path, must_exist=True)
        if err:
            return err
    err = validate_numerics(
        x=x, y=y, z=z, origin_x=origin_x, origin_y=origin_y,
        origin_z=origin_z, angle_x=angle_x, angle_y=angle_y, angle_z=angle_z,
    )
    if err:
        return err
    match method:
        case "basic":
            return assembly_manager.add_component(
                file_path, x, y, z
            )
        case "with_transform":
            return assembly_manager.add_component_with_transform(
                file_path,
                origin_x, origin_y, origin_z,
                angle_x, angle_y, angle_z,
            )
        case "family":
            return assembly_manager.add_family_member(
                file_path, family_member_name, x, y, z
            )
        case "family_with_transform":
            return assembly_manager.add_family_with_transform(
                file_path, family_member_name,
                origin_x, origin_y, origin_z,
                angle_x, angle_y, angle_z,
            )
        case "family_with_matrix":
            return assembly_manager.add_family_with_matrix(
                file_path, family_member_name,
                matrix or [],
            )
        case "by_template":
            return assembly_manager.add_by_template(
                file_path, template_name
            )
        case "adjustable":
            return assembly_manager.add_adjustable_part(
                file_path, x, y, z
            )
        case "tube":
            return assembly_manager.add_tube(
                segment_indices or [],
                file_path,
            )
        case _:
            return {
                "error": f"Unknown method: {method}"
            }


# ================================================================
# Group 73: manage_component (6 -> 1)
# ================================================================


def manage_component(
    action: str,
    component_index: int = 0,
    new_file_path: str = "",
    suppress: bool = True,
    target_index: int = 0,
    new_member_name: str = "",
    ground: bool = True,
    plane_index: int = 1,
    count: int = 1,
    spacing: float = 0.0,
    direction: str = "X",
) -> dict:
    """Manage an assembly component.

    action: 'delete' | 'replace' | 'suppress' | 'reorder'
      | 'make_writable' | 'swap_family' | 'ground'
      | 'pattern' | 'mirror'

    Spacing in meters. plane_index: 1=Top, 2=Front, 3=Right.
    """
    if action == "replace" and new_file_path:
        new_file_path, err = validate_path(new_file_path, must_exist=True)
        if err:
            return err
    err = validate_numerics(spacing=spacing)
    if err:
        return err
    match action:
        case "delete":
            return assembly_manager.delete_component(
                component_index
            )
        case "replace":
            return assembly_manager.replace_component(
                component_index, new_file_path
            )
        case "suppress":
            return assembly_manager.suppress_component(
                component_index, suppress
            )
        case "reorder":
            return assembly_manager.reorder_occurrence(
                component_index, target_index
            )
        case "make_writable":
            return assembly_manager.make_writable(
                component_index
            )
        case "swap_family":
            return assembly_manager.swap_family_member(
                component_index, new_member_name
            )
        case "ground":
            return assembly_manager.ground_component(
                component_index, ground
            )
        case "pattern":
            return assembly_manager.pattern_component(
                component_index, count, spacing, direction
            )
        case "mirror":
            return assembly_manager.mirror_component(
                component_index, plane_index
            )
        case _:
            return {
                "error": f"Unknown action: {action}"
            }


# ================================================================
# Group 74: query_component (18 -> 1)
# ================================================================


def query_component(
    property: str = "list",
    component_index: int = 0,
    internal_id: int = 0,
) -> dict:
    """Query assembly component information.

    property: 'list' | 'info' | 'bounding_box' | 'bom'
      | 'structured_bom' | 'tree' | 'transform' | 'count'
      | 'is_subassembly' | 'display_name' | 'document'
      | 'sub_occurrences' | 'bodies' | 'style'
      | 'is_tube' | 'adjustable_part' | 'face_style'
      | 'occurrence' | 'interference' | 'tube'

    0-based component_index. 'occurrence' uses internal_id instead.
    """
    match property:
        case "list":
            return assembly_manager.list_components()
        case "info":
            return assembly_manager.get_component_info(
                component_index
            )
        case "bounding_box":
            return assembly_manager.get_occurrence_bounding_box(
                component_index
            )
        case "bom":
            return assembly_manager.get_bom()
        case "structured_bom":
            return assembly_manager.get_structured_bom()
        case "tree":
            return assembly_manager.get_document_tree()
        case "transform":
            return assembly_manager.get_component_transform(
                component_index
            )
        case "count":
            return assembly_manager.get_occurrence_count()
        case "is_subassembly":
            return assembly_manager.is_subassembly(
                component_index
            )
        case "display_name":
            return assembly_manager.get_component_display_name(
                component_index
            )
        case "document":
            return assembly_manager.get_occurrence_document(
                component_index
            )
        case "sub_occurrences":
            return assembly_manager.get_sub_occurrences(
                component_index
            )
        case "bodies":
            return assembly_manager.get_occurrence_bodies(
                component_index
            )
        case "style":
            return assembly_manager.get_occurrence_style(
                component_index
            )
        case "is_tube":
            return assembly_manager.is_tube(component_index)
        case "adjustable_part":
            return assembly_manager.get_adjustable_part(
                component_index
            )
        case "face_style":
            return assembly_manager.get_face_style(
                component_index
            )
        case "occurrence":
            return assembly_manager.get_occurrence(
                internal_id
            )
        case "interference":
            return assembly_manager.check_interference(
                component_index if component_index else None
            )
        case "tube":
            return assembly_manager.get_tube(component_index)
        case _:
            return {
                "error": f"Unknown property: {property}"
            }


# ================================================================
# Group 75: set_component_appearance (2 -> 1)
# ================================================================


def set_component_appearance(
    property: str,
    component_index: int = 0,
    visible: bool = True,
    red: int = 0,
    green: int = 0,
    blue: int = 0,
) -> dict:
    """Set a component's visual appearance.

    property: 'visibility' | 'color'

    RGB values 0-255.
    """
    match property:
        case "visibility":
            return assembly_manager.set_component_visibility(
                component_index, visible
            )
        case "color":
            return assembly_manager.set_component_color(
                component_index, red, green, blue
            )
        case _:
            return {
                "error": f"Unknown property: {property}"
            }


# ================================================================
# Group 76a: transform_component (4 of 7)
# ================================================================


def transform_component(
    method: str,
    component_index: int = 0,
    x: float = 0,
    y: float = 0,
    z: float = 0,
    dx: float = 0,
    dy: float = 0,
    dz: float = 0,
) -> dict:
    """Move or reposition a component.

    method: 'update_position' | 'set_origin' | 'put_origin' | 'move'

    Coordinates in meters. 0-based component_index.
    """
    err = validate_numerics(
        x=x, y=y, z=z, dx=dx, dy=dy, dz=dz,
    )
    if err:
        return err
    match method:
        case "update_position":
            return assembly_manager.update_component_position(
                component_index, x, y, z
            )
        case "set_origin":
            return assembly_manager.set_component_origin(
                component_index, x, y, z
            )
        case "put_origin":
            return assembly_manager.put_origin(
                component_index, x, y, z
            )
        case "move":
            return assembly_manager.occurrence_move(
                component_index, dx, dy, dz
            )
        case _:
            return {
                "error": f"Unknown method: {method}"
            }


# ================================================================
# Group 76b: set_component_orientation (2 of 7)
# ================================================================


def set_component_orientation(
    method: str,
    component_index: int = 0,
    origin_x: float = 0,
    origin_y: float = 0,
    origin_z: float = 0,
    angle_x: float = 0,
    angle_y: float = 0,
    angle_z: float = 0,
    x: float = 0,
    y: float = 0,
    z: float = 0,
    rx: float = 0,
    ry: float = 0,
    rz: float = 0,
) -> dict:
    """Set a component's full transform (position + orientation).

    method: 'set_transform' | 'put_euler'

    - set_transform: position via origin_x/y/z (meters), rotation via angle_x/y/z (degrees)
    - put_euler: position via x/y/z (meters), rotation via rx/ry/rz (degrees)
    """
    err = validate_numerics(
        origin_x=origin_x, origin_y=origin_y, origin_z=origin_z,
        angle_x=angle_x, angle_y=angle_y, angle_z=angle_z,
        x=x, y=y, z=z, rx=rx, ry=ry, rz=rz,
    )
    if err:
        return err
    match method:
        case "set_transform":
            return assembly_manager.set_component_transform(
                component_index,
                origin_x, origin_y, origin_z,
                angle_x, angle_y, angle_z,
            )
        case "put_euler":
            return assembly_manager.put_transform_euler(
                component_index,
                x, y, z, rx, ry, rz,
            )
        case _:
            return {
                "error": f"Unknown method: {method}"
            }


# ================================================================
# Group 76c: rotate_component (1 of 7)
# ================================================================


def rotate_component(
    component_index: int = 0,
    axis_x1: float = 0,
    axis_y1: float = 0,
    axis_z1: float = 0,
    axis_x2: float = 0,
    axis_y2: float = 0,
    axis_z2: float = 0,
    angle: float = 0,
) -> dict:
    """Rotate a component around an axis.

    Axis defined by two points (meters). Angle in degrees. 0-based component_index.
    """
    err = validate_numerics(
        angle=angle,
        axis_x1=axis_x1, axis_y1=axis_y1, axis_z1=axis_z1,
        axis_x2=axis_x2, axis_y2=axis_y2, axis_z2=axis_z2,
    )
    if err:
        return err
    return assembly_manager.occurrence_rotate(
        component_index,
        axis_x1, axis_y1, axis_z1,
        axis_x2, axis_y2, axis_z2,
        angle,
    )


# ================================================================
# Group 77: add_assembly_constraint (5 -> 1)
# ================================================================


def add_assembly_constraint(
    type: str,
    component1_index: int = 0,
    component2_index: int = 0,
    mate_type: str = "Mate",
    angle: float = 0,
) -> dict:
    """Add a constraint between two assembly components.

    type: 'mate' | 'align' | 'planar_align' | 'axial_align' | 'angle'

    Angle in degrees. mate_type: 'Mate'/'PlanarAlign'/'AxialAlign'.
    """
    err = validate_numerics(angle=angle)
    if err:
        return err
    match type:
        case "mate":
            return assembly_manager.create_mate(
                mate_type, component1_index, component2_index
            )
        case "align":
            return assembly_manager.add_align_constraint(
                component1_index, component2_index
            )
        case "planar_align":
            return assembly_manager.add_planar_align_constraint(
                component1_index, component2_index
            )
        case "axial_align":
            return assembly_manager.add_axial_align_constraint(
                component1_index, component2_index
            )
        case "angle":
            return assembly_manager.add_angle_constraint(
                component1_index, component2_index, angle
            )
        case _:
            return {
                "error": f"Unknown type: {type}"
            }


# ================================================================
# Group 78: add_assembly_relation (6 -> 1)
# ================================================================


def add_assembly_relation(
    type: str,
    occurrence1_index: int = 0,
    occurrence2_index: int = 0,
    offset: float = 0.0,
    orientation: str = "Align",
    angle: float = 0.0,
    ratio1: float = 1.0,
    ratio2: float = 1.0,
) -> dict:
    """Add a relation between two assembly components.

    type: 'planar' | 'axial' | 'angular' | 'point' | 'tangent' | 'gear'

    Offset in meters. Angle in degrees.
    """
    err = validate_numerics(offset=offset, angle=angle, ratio1=ratio1, ratio2=ratio2)
    if err:
        return err
    match type:
        case "planar":
            return assembly_manager.add_planar_relation(
                occurrence1_index, occurrence2_index,
                offset, orientation,
            )
        case "axial":
            return assembly_manager.add_axial_relation(
                occurrence1_index, occurrence2_index,
                orientation,
            )
        case "angular":
            return assembly_manager.add_angular_relation(
                occurrence1_index, occurrence2_index,
                angle,
            )
        case "point":
            return assembly_manager.add_point_relation(
                occurrence1_index, occurrence2_index
            )
        case "tangent":
            return assembly_manager.add_tangent_relation(
                occurrence1_index, occurrence2_index
            )
        case "gear":
            return assembly_manager.add_gear_relation(
                occurrence1_index, occurrence2_index,
                ratio1, ratio2,
            )
        case _:
            return {
                "error": f"Unknown type: {type}"
            }


# ================================================================
# Group 79: manage_relation (13 -> 1)
# ================================================================


def manage_relation(
    action: str,
    relation_index: int = 0,
    offset: float = 0.0,
    angle: float = 0.0,
    aligned: bool = True,
) -> dict:
    """Manage assembly relations/constraints.

    action: 'list' | 'info' | 'delete' | 'get_offset'
      | 'set_offset' | 'get_angle' | 'set_angle'
      | 'get_normals' | 'set_normals' | 'suppress'
      | 'unsuppress' | 'get_geometry' | 'get_gear_ratio'

    Offset in meters. Angle in degrees.
    """
    err = validate_numerics(offset=offset, angle=angle)
    if err:
        return err
    match action:
        case "list":
            return assembly_manager.get_assembly_relations()
        case "info":
            return assembly_manager.get_relation_info(
                relation_index
            )
        case "delete":
            return assembly_manager.delete_relation(
                relation_index
            )
        case "get_offset":
            return assembly_manager.get_relation_offset(
                relation_index
            )
        case "set_offset":
            return assembly_manager.set_relation_offset(
                relation_index, offset
            )
        case "get_angle":
            return assembly_manager.get_relation_angle(
                relation_index
            )
        case "set_angle":
            return assembly_manager.set_relation_angle(
                relation_index, angle
            )
        case "get_normals":
            return assembly_manager.get_normals_aligned(
                relation_index
            )
        case "set_normals":
            return assembly_manager.set_normals_aligned(
                relation_index, aligned
            )
        case "suppress":
            return assembly_manager.suppress_relation(
                relation_index
            )
        case "unsuppress":
            return assembly_manager.unsuppress_relation(
                relation_index
            )
        case "get_geometry":
            return assembly_manager.get_relation_geometry(
                relation_index
            )
        case "get_gear_ratio":
            return assembly_manager.get_gear_ratio(
                relation_index
            )
        case _:
            return {
                "error": f"Unknown action: {action}"
            }


# ================================================================
# Group 80: assembly_feature (9 -> 1)
# ================================================================


def assembly_feature(
    type: str,
    scope_parts: list[int] | None = None,
    extent_type: str = "Finite",
    extent_side: str = "OneSide",
    profile_side: str = "Left",
    distance: float = 0.01,
    angle: float = 360.0,
    depth: float = 0.01,
    feature_indices: list[int] | None = None,
    plane_index: int = 1,
    mirror_type: int = 1,
    pattern_type: str = "Rectangular",
    num_trace_curves: int = 1,
    num_cross_sections: int = 1,
    options: int = 0,
) -> dict:
    """Create or manage assembly-level features.

    type: 'extruded_cutout' | 'revolved_cutout' | 'hole'
      | 'extruded_protrusion' | 'revolved_protrusion'
      | 'mirror' | 'pattern' | 'swept_protrusion'
      | 'recompute'

    Distance/depth in meters. Angle in degrees.
    """
    err = validate_numerics(distance=distance, angle=angle, depth=depth)
    if err:
        return err
    match type:
        case "extruded_cutout":
            return (
                assembly_manager.create_assembly_extruded_cutout(
                    scope_parts or [],
                    extent_type,
                    extent_side,
                    profile_side,
                    distance,
                )
            )
        case "revolved_cutout":
            return (
                assembly_manager.create_assembly_revolved_cutout(
                    scope_parts or [],
                    extent_type,
                    extent_side,
                    profile_side,
                    angle,
                )
            )
        case "hole":
            return assembly_manager.create_assembly_hole(
                scope_parts or [],
                extent_type,
                extent_side,
                depth,
            )
        case "extruded_protrusion":
            return (
                assembly_manager
                .create_assembly_extruded_protrusion(
                    extent_type,
                    extent_side,
                    profile_side,
                    distance,
                )
            )
        case "revolved_protrusion":
            return (
                assembly_manager
                .create_assembly_revolved_protrusion(
                    extent_type,
                    extent_side,
                    profile_side,
                    angle,
                )
            )
        case "mirror":
            return assembly_manager.create_assembly_mirror(
                feature_indices or [],
                plane_index,
                mirror_type,
            )
        case "pattern":
            return assembly_manager.create_assembly_pattern(
                feature_indices or [],
                pattern_type,
            )
        case "swept_protrusion":
            return (
                assembly_manager
                .create_assembly_swept_protrusion(
                    num_trace_curves,
                    num_cross_sections,
                )
            )
        case "recompute":
            return (
                assembly_manager.recompute_assembly_features(
                    options
                )
            )
        case _:
            return {
                "error": f"Unknown type: {type}"
            }


# ================================================================
# Group 81: virtual_component (3 -> 1)
# ================================================================


def virtual_component(
    method: str = "new",
    name: str = "",
    component_type: str = "Part",
    filename: str = "",
    doc_number: str = "",
    revision_id: str = "",
) -> dict:
    """Manage virtual components in the assembly.

    method: 'new' | 'predefined' | 'bidm'
    """
    if method == "predefined" and filename:
        filename, err = validate_path(filename, must_exist=True)
        if err:
            return err
    match method:
        case "new":
            return assembly_manager.add_virtual_component(
                name, component_type
            )
        case "predefined":
            return assembly_manager.add_virtual_component_predefined(
                filename
            )
        case "bidm":
            return assembly_manager.add_virtual_component_bidm(
                doc_number, revision_id, component_type
            )
        case _:
            return {
                "error": f"Unknown method: {method}"
            }


# ================================================================
# Group 82: structural_frame (2 -> 1)
# ================================================================


def structural_frame(
    method: str = "basic",
    part_filename: str = "",
    path_indices: list[int] | None = None,
    coord_system_name: str = "",
) -> dict:
    """Create a structural frame in the assembly.

    method: 'basic' | 'by_orientation'
    """
    if part_filename:
        part_filename, err = validate_path(part_filename, must_exist=True)
        if err:
            return err
    match method:
        case "basic":
            return assembly_manager.add_structural_frame(
                part_filename, path_indices or []
            )
        case "by_orientation":
            return (
                assembly_manager.add_structural_frame_by_orientation(
                    part_filename,
                    coord_system_name,
                    path_indices or [],
                )
            )
        case _:
            return {
                "error": f"Unknown method: {method}"
            }


# ================================================================
# Group 83: wiring (4 -> 1)
# ================================================================


def wiring(
    type: str = "wire",
    path_indices: list[int] | None = None,
    path_directions: list[bool] | None = None,
    conductor_indices: list[int] | None = None,
    wire_indices: list[int] | None = None,
    split_path_indices: list[int] | None = None,
    split_path_directions: list[bool] | None = None,
    description: str = "",
    x: float = 0,
    y: float = 0,
    z: float = 0,
) -> dict:
    """Create wiring harness elements in the assembly.

    type: 'wire' | 'cable' | 'bundle' | 'splice'

    Splice position in meters.
    """
    err = validate_numerics(x=x, y=y, z=z)
    if err:
        return err
    match type:
        case "wire":
            return assembly_manager.add_wire(
                path_indices or [],
                path_directions or [],
                description,
            )
        case "cable":
            return assembly_manager.add_cable(
                path_indices or [],
                path_directions or [],
                wire_indices or [],
                split_path_indices,
                split_path_directions,
                description,
            )
        case "bundle":
            return assembly_manager.add_bundle(
                path_indices or [],
                path_directions or [],
                conductor_indices or [],
                split_path_indices,
                split_path_directions,
                description,
            )
        case "splice":
            return assembly_manager.add_splice(
                x, y, z,
                conductor_indices or [],
                description,
            )
        case _:
            return {
                "error": f"Unknown type: {type}"
            }


# ================================================================
# Registration
# ================================================================


def register(mcp):
    """Register assembly tools with the MCP server."""
    mcp.tool()(add_assembly_component)
    mcp.tool()(manage_component)
    mcp.tool()(query_component)
    mcp.tool()(set_component_appearance)
    mcp.tool()(transform_component)
    mcp.tool()(set_component_orientation)
    mcp.tool()(rotate_component)
    mcp.tool()(add_assembly_constraint)
    mcp.tool()(add_assembly_relation)
    mcp.tool()(manage_relation)
    mcp.tool()(assembly_feature)
    mcp.tool()(virtual_component)
    mcp.tool()(structural_frame)
    mcp.tool()(wiring)
