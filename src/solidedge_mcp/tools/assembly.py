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

    - basic: place at (x, y, z) in meters
    - with_transform: place with position (meters) + Euler angles (degrees)
    - family: Family of Parts member at (x, y, z) in meters
    - family_with_transform: family member with Euler angles (degrees)
    - family_with_matrix: family member with 4x4 matrix
    - by_template: place using a template name
    - adjustable: place as adjustable part at (x, y, z) in meters
    - tube: add tube from segment_indices with file_path template

    Parameters:
        x, y, z: Position coordinates in meters (basic, family, adjustable).
        origin_x, origin_y, origin_z: Position in meters (with_transform,
            family_with_transform).
        angle_x, angle_y, angle_z: Euler angles in degrees (with_transform,
            family_with_transform).
        matrix: 16-element 4x4 transform matrix (family_with_matrix).
    """
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

    - delete: remove component at component_index
    - replace: replace with new_file_path
    - suppress: suppress/unsuppress (suppress=True/False)
    - reorder: move to target_index in tree
    - make_writable: make component editable
    - swap_family: swap Family of Parts member name
    - ground: fix/unfix component (ground=True/False)
    - pattern: linear pattern (count, spacing in meters, direction)
    - mirror: mirror across plane_index (1=Top,2=Front,3=Right)
    """
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

    - list: list all components
    - info: detailed info for component_index
    - bounding_box: bounding box of component_index
    - bom: flat Bill of Materials
    - structured_bom: hierarchical Bill of Materials
    - tree: hierarchical component tree
    - transform: full 4x4 matrix of component_index
    - count: number of top-level occurrences
    - is_subassembly: check if component is subassembly
    - display_name: user-visible label in assembly tree
    - document: source file info for component_index
    - sub_occurrences: children of component_index
    - bodies: body information from component_index
    - style: appearance style of component_index
    - is_tube: check if component is a tube
    - adjustable_part: adjustable part info
    - face_style: face style of component_index
    - occurrence: get occurrence by internal_id
    - tube: get tube properties from tube component
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

    - visibility: show/hide component (visible=True/False)
    - color: set RGB color (red, green, blue 0-255)
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
# Group 76: transform_component (7 -> 1)
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
    origin_x: float = 0,
    origin_y: float = 0,
    origin_z: float = 0,
    angle_x: float = 0,
    angle_y: float = 0,
    angle_z: float = 0,
    rx: float = 0,
    ry: float = 0,
    rz: float = 0,
    axis_x1: float = 0,
    axis_y1: float = 0,
    axis_z1: float = 0,
    axis_x2: float = 0,
    axis_y2: float = 0,
    axis_z2: float = 0,
    angle: float = 0,
) -> dict:
    """Transform (move/rotate/position) a component.

    method: 'update_position' | 'set_transform'
      | 'set_origin' | 'move' | 'rotate'
      | 'put_euler' | 'put_origin'

    All position/coordinate parameters are in meters.
    All angle parameters are in degrees.

    - update_position: set position to (x, y, z) meters
    - set_transform: set position (origin_x/y/z meters)
      + rotation (angle_x/y/z degrees)
    - set_origin: set origin to (x, y, z) meters
    - move: translate by (dx, dy, dz) meters
    - rotate: rotate around axis by angle (degrees);
      axis defined by points (axis_x1..z1) to (axis_x2..z2) meters
    - put_euler: set Euler angles (rx, ry, rz degrees)
      and position (x, y, z) meters
    - put_origin: set origin to (x, y, z) meters
    """
    match method:
        case "update_position":
            return assembly_manager.update_component_position(
                component_index, x, y, z
            )
        case "set_transform":
            return assembly_manager.set_component_transform(
                component_index,
                origin_x, origin_y, origin_z,
                angle_x, angle_y, angle_z,
            )
        case "set_origin":
            return assembly_manager.set_component_origin(
                component_index, x, y, z
            )
        case "move":
            return assembly_manager.occurrence_move(
                component_index, dx, dy, dz
            )
        case "rotate":
            return assembly_manager.occurrence_rotate(
                component_index,
                axis_x1, axis_y1, axis_z1,
                axis_x2, axis_y2, axis_z2,
                angle,
            )
        case "put_euler":
            return assembly_manager.put_transform_euler(
                component_index,
                x, y, z, rx, ry, rz,
            )
        case "put_origin":
            return assembly_manager.put_origin(
                component_index, x, y, z
            )
        case _:
            return {
                "error": f"Unknown method: {method}"
            }


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

    type: 'mate' | 'align' | 'planar_align'
      | 'axial_align' | 'angle'

    - mate: create mate (mate_type: 'Mate'/'PlanarAlign'/
      'AxialAlign')
    - align: search-based align constraint
    - planar_align: search-based planar align
    - axial_align: search-based axial align
    - angle: angle constraint (angle in degrees)
    """
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

    type: 'planar' | 'axial' | 'angular' | 'point'
      | 'tangent' | 'gear'

    - planar: with offset (meters), orientation
    - axial: with orientation
    - angular: with angle (degrees)
    - point: connect relation
    - tangent: tangent relation
    - gear: gear relation with ratio1, ratio2
    """
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

    - list: get all assembly relations
    - info: detailed info for relation_index
    - delete: delete relation at relation_index
    - get_offset: get offset from planar relation
    - set_offset: set offset (meters) on planar relation
    - get_angle: get angle from angular relation (deg)
    - set_angle: set angle on angular relation (deg)
    - get_normals: get NormalsAligned boolean
    - set_normals: set NormalsAligned (aligned=True/False)
    - suppress: suppress relation
    - unsuppress: unsuppress relation
    - get_geometry: get geometry info from relation
    - get_gear_ratio: get gear ratio values
    """
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

    - extruded_cutout: cut across scope_parts with
      extent_type, extent_side, profile_side, distance (meters)
    - revolved_cutout: revolve cut across scope_parts
      with angle (degrees)
    - hole: hole across scope_parts with depth (meters)
    - extruded_protrusion: extrude with distance (meters)
    - revolved_protrusion: revolve with angle (degrees)
    - mirror: mirror feature_indices across plane_index
    - pattern: pattern feature_indices (Rectangular/Circular)
    - swept_protrusion: sweep with trace/cross-section counts
    - recompute: recompute all assembly features
    """
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

    - new: create a virtual placeholder (name, component_type)
    - predefined: add from existing file (filename)
    - bidm: add by document number + revision
      (doc_number, revision_id, component_type)
    """
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

    - basic: add frame from part_filename along path_indices
    - by_orientation: add with coordinate system orientation
      (part_filename, coord_system_name, path_indices)
    """
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

    - wire: add wire along path (path_indices, path_directions)
    - cable: add cable with wires (+ wire_indices)
    - bundle: add bundle of conductors (+ conductor_indices)
    - splice: add splice at position (x, y, z in meters, conductor_indices)
    """
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
    mcp.tool()(add_assembly_constraint)
    mcp.tool()(add_assembly_relation)
    mcp.tool()(manage_relation)
    mcp.tool()(assembly_feature)
    mcp.tool()(virtual_component)
    mcp.tool()(structural_frame)
    mcp.tool()(wiring)
