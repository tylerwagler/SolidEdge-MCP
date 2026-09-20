"""Assembly tools for Solid Edge MCP."""

from typing import Any, Literal

from solidedge_mcp.backends.validation import validate_numerics, validate_path
from solidedge_mcp.managers import assembly_manager
from solidedge_mcp.tools._registry import register_tool

# ================================================================
# Group 72: add_assembly_component (7 -> 1)
# ================================================================


def add_assembly_component(
    method: Literal[
        "basic",
        "with_transform",
        "family",
        "family_with_transform",
        "family_with_matrix",
        "by_template",
        "adjustable",
        "tube",
    ] = "basic",
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
) -> dict[str, Any]:
    """Place a component (file_path) into the active assembly.

    basic/adjustable: x,y,z position. with_transform: origin_x/y/z + angle_x/y/z.
    family/family_with_transform/family_with_matrix: also family_member_name;
    family_with_matrix uses a 16-element row-major 4x4 matrix.
    by_template: template_name. tube: segment_indices (0-based path segments).
    Positions in meters, angles in degrees.
    """
    if file_path:
        file_path, err = validate_path(file_path, must_exist=True)
        if err:
            return err
    err = validate_numerics(
        x=x,
        y=y,
        z=z,
        origin_x=origin_x,
        origin_y=origin_y,
        origin_z=origin_z,
        angle_x=angle_x,
        angle_y=angle_y,
        angle_z=angle_z,
    )
    if err:
        return err
    match method:
        case "basic":
            return assembly_manager.add_component(file_path=file_path, x=x, y=y, z=z)
        case "with_transform":
            return assembly_manager.add_component_with_transform(
                file_path=file_path,
                origin_x=origin_x,
                origin_y=origin_y,
                origin_z=origin_z,
                angle_x=angle_x,
                angle_y=angle_y,
                angle_z=angle_z,
            )
        case "family":
            return assembly_manager.add_family_member(
                file_path=file_path, family_member_name=family_member_name, x=x, y=y, z=z
            )
        case "family_with_transform":
            return assembly_manager.add_family_with_transform(
                file_path=file_path,
                family_member_name=family_member_name,
                origin_x=origin_x,
                origin_y=origin_y,
                origin_z=origin_z,
                angle_x=angle_x,
                angle_y=angle_y,
                angle_z=angle_z,
            )
        case "family_with_matrix":
            return assembly_manager.add_family_with_matrix(
                family_file_path=file_path, member_name=family_member_name, matrix=matrix or []
            )
        case "by_template":
            return assembly_manager.add_by_template(
                file_path=file_path, template_name=template_name
            )
        case "adjustable":
            return assembly_manager.add_adjustable_part(file_path=file_path, x=x, y=y, z=z)
        case "tube":
            return assembly_manager.add_tube(
                segment_indices=segment_indices or [], part_filename=file_path
            )
        case _:
            return {"error": f"Unknown method: {method}"}


# ================================================================
# Group 73: manage_component (6 -> 1)
# ================================================================


def manage_component(
    action: Literal[
        "delete",
        "replace",
        "suppress",
        "reorder",
        "make_writable",
        "swap_family",
        "ground",
        "pattern",
        "mirror",
    ],
    component_index: int = 0,
    new_file_path: str = "",
    replace_all: bool = False,
    suppress: bool = True,
    target_index: int = 0,
    new_member_name: str = "",
    ground: bool = True,
    plane_index: int = 1,
    count: int = 1,
    spacing: float = 0.0,
    direction: Literal["X", "Y", "Z"] = "X",
) -> dict[str, Any]:
    """Modify one assembly component (0-based component_index).

    delete removes it. replace: new_file_path, plus replace_all to swap every
    occurrence of that file rather than this one. suppress: suppress flag
    (suppress=False is unsupported - COM offers no way back to the
    SuppressComponent object; unsuppress in the Solid Edge UI).
    reorder: target_index. swap_family: new_member_name. ground: ground flag.
    pattern: linear copies via count, spacing (meters), direction X/Y/Z.
    mirror: plane_index is 1-based (1=Top/XY, 2=Right/YZ, 3=Front/XZ).
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
            return assembly_manager.delete_component(component_index=component_index)
        case "replace":
            return assembly_manager.replace_component(
                component_index=component_index,
                new_file_path=new_file_path,
                replace_all=replace_all,
            )
        case "suppress":
            return assembly_manager.suppress_component(
                component_index=component_index, suppress=suppress
            )
        case "reorder":
            return assembly_manager.reorder_occurrence(
                component_index=component_index, target_index=target_index
            )
        case "make_writable":
            return assembly_manager.make_writable(component_index=component_index)
        case "swap_family":
            return assembly_manager.swap_family_member(
                component_index=component_index, new_member_name=new_member_name
            )
        case "ground":
            return assembly_manager.ground_component(component_index=component_index, ground=ground)
        case "pattern":
            return assembly_manager.pattern_component(
                component_index=component_index, count=count, spacing=spacing, direction=direction
            )
        case "mirror":
            return assembly_manager.mirror_component(
                component_index=component_index, plane_index=plane_index
            )
        case _:
            return {"error": f"Unknown action: {action}"}


# ================================================================
# Group 74: query_component (18 -> 1)
# ================================================================


def query_component(
    property: Literal[
        "list",
        "info",
        "bounding_box",
        "bom",
        "structured_bom",
        "tree",
        "transform",
        "count",
        "is_subassembly",
        "display_name",
        "document",
        "sub_occurrences",
        "bodies",
        "style",
        "is_tube",
        "adjustable_part",
        "face_style",
        "occurrence",
        "interference",
        "tube",
    ] = "list",
    component_index: int | None = None,
    internal_id: int = 0,
) -> dict[str, Any]:
    """Read assembly component data (read-only).

    list/bom/structured_bom/tree/count ignore component_index.
    Per-component properties use 0-based component_index (default 0).
    'occurrence' looks up by internal_id instead.
    'interference' checks component_index against all others, or every
    pair when component_index is omitted.
    """
    idx = component_index if component_index is not None else 0
    match property:
        case "list":
            return assembly_manager.list_components()
        case "info":
            return assembly_manager.get_component_info(component_index=idx)
        case "bounding_box":
            return assembly_manager.get_occurrence_bounding_box(component_index=idx)
        case "bom":
            return assembly_manager.get_bom()
        case "structured_bom":
            return assembly_manager.get_structured_bom()
        case "tree":
            return assembly_manager.get_document_tree()
        case "transform":
            return assembly_manager.get_component_transform(component_index=idx)
        case "count":
            return assembly_manager.get_occurrence_count()
        case "is_subassembly":
            return assembly_manager.is_subassembly(component_index=idx)
        case "display_name":
            return assembly_manager.get_component_display_name(component_index=idx)
        case "document":
            return assembly_manager.get_occurrence_document(component_index=idx)
        case "sub_occurrences":
            return assembly_manager.get_sub_occurrences(component_index=idx)
        case "bodies":
            return assembly_manager.get_occurrence_bodies(component_index=idx)
        case "style":
            return assembly_manager.get_occurrence_style(component_index=idx)
        case "is_tube":
            return assembly_manager.is_tube(component_index=idx)
        case "adjustable_part":
            return assembly_manager.get_adjustable_part(component_index=idx)
        case "face_style":
            return assembly_manager.get_face_style(component_index=idx)
        case "occurrence":
            return assembly_manager.get_occurrence(internal_id=internal_id)
        case "interference":
            # None means "check all pairs"; index 0 is a real component.
            return assembly_manager.check_interference(
                component_index=component_index if component_index is not None else None
            )
        case "tube":
            return assembly_manager.get_tube(component_index=idx)
        case _:
            return {"error": f"Unknown property: {property}"}


# ================================================================
# Group 75: set_component_appearance (2 -> 1)
# ================================================================


def set_component_appearance(
    property: Literal["visibility", "color"],
    component_index: int = 0,
    visible: bool = True,
    red: int = 0,
    green: int = 0,
    blue: int = 0,
) -> dict[str, Any]:
    """Set visibility or RGB color (0-255) of a component (0-based index)."""
    match property:
        case "visibility":
            return assembly_manager.set_component_visibility(
                component_index=component_index, visible=visible
            )
        case "color":
            return assembly_manager.set_component_color(
                component_index=component_index, red=red, green=green, blue=blue
            )
        case _:
            return {"error": f"Unknown property: {property}"}


# ================================================================
# Group 76a: transform_component (4 of 7)
# ================================================================


def transform_component(
    method: Literal["update_position", "set_origin", "put_origin", "move"],
    component_index: int = 0,
    x: float = 0,
    y: float = 0,
    z: float = 0,
    dx: float = 0,
    dy: float = 0,
    dz: float = 0,
) -> dict[str, Any]:
    """Reposition a component (0-based index). Meters.

    update_position/set_origin/put_origin: absolute x,y,z.
    move: relative dx,dy,dz.
    """
    err = validate_numerics(
        x=x,
        y=y,
        z=z,
        dx=dx,
        dy=dy,
        dz=dz,
    )
    if err:
        return err
    match method:
        case "update_position":
            return assembly_manager.update_component_position(
                component_index=component_index, x=x, y=y, z=z
            )
        case "set_origin":
            return assembly_manager.set_component_origin(
                component_index=component_index, x=x, y=y, z=z
            )
        case "put_origin":
            return assembly_manager.put_origin(component_index=component_index, x=x, y=y, z=z)
        case "move":
            return assembly_manager.occurrence_move(
                component_index=component_index, dx=dx, dy=dy, dz=dz
            )
        case _:
            return {"error": f"Unknown method: {method}"}


# ================================================================
# Group 76b: set_component_orientation (2 of 7)
# ================================================================


def set_component_orientation(
    method: Literal["set_transform", "put_euler"],
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
) -> dict[str, Any]:
    """Set full position + rotation of a component (0-based index).

    Position in meters and rotation in degrees, named either way: origin_x/y/z
    with angle_x/y/z, or x/y/z with rx/ry/rz. Whichever set you fill is used,
    so naming the rotation the other way no longer rotates by zero and reports
    success.
    """
    err = validate_numerics(
        origin_x=origin_x,
        origin_y=origin_y,
        origin_z=origin_z,
        angle_x=angle_x,
        angle_y=angle_y,
        angle_z=angle_z,
        x=x,
        y=y,
        z=z,
        rx=rx,
        ry=ry,
        rz=rz,
    )
    if err:
        return err
    # Each branch used to read only its own vocabulary, so naming the rotation
    # the other way left it at zero: the component did not move and the result
    # still said "updated". Each branch now prefers its own names and falls
    # back to the other set.
    match method:
        case "set_transform":
            return assembly_manager.set_component_transform(
                component_index=component_index,
                origin_x=origin_x or x,
                origin_y=origin_y or y,
                origin_z=origin_z or z,
                angle_x=angle_x or rx,
                angle_y=angle_y or ry,
                angle_z=angle_z or rz,
            )
        case "put_euler":
            return assembly_manager.put_transform_euler(
                component_index=component_index,
                x=x or origin_x,
                y=y or origin_y,
                z=z or origin_z,
                rx=rx or angle_x,
                ry=ry or angle_y,
                rz=rz or angle_z,
            )
        case _:
            return {"error": f"Unknown method: {method}"}


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
) -> dict[str, Any]:
    """Rotate a component (0-based index) about the axis from point 1 to point 2.

    Axis points in meters, angle in degrees.
    """
    err = validate_numerics(
        angle=angle,
        axis_x1=axis_x1,
        axis_y1=axis_y1,
        axis_z1=axis_z1,
        axis_x2=axis_x2,
        axis_y2=axis_y2,
        axis_z2=axis_z2,
    )
    if err:
        return err
    return assembly_manager.occurrence_rotate(
        component_index=component_index,
        axis_x1=axis_x1,
        axis_y1=axis_y1,
        axis_z1=axis_z1,
        axis_x2=axis_x2,
        axis_y2=axis_y2,
        axis_z2=axis_z2,
        angle=angle,
    )


# ================================================================
# Group 77: add_assembly_constraint (5 -> 1)
# ================================================================


def add_assembly_constraint(
    type: Literal["mate", "align", "planar_align", "axial_align", "angle"],
    component1_index: int = 0,
    component2_index: int = 0,
    mate_type: str = "Mate",
    angle: float = 0,
    face1_index: int = 0,
    face2_index: int = 0,
) -> dict[str, Any]:
    """Constrain two components (0-based indices) through their faces.

    mate / align / planar_align relate planar faces face1_index and
    face2_index (0-based faces of each component's body): mate makes them
    touch, align aligns them. axial_align makes two cylindrical faces
    coaxial. angle takes angle in degrees. All go through
    AssemblyDocument.CreateReference, the documented route.
    """
    err = validate_numerics(angle=angle)
    if err:
        return err
    match type:
        case "mate":
            return assembly_manager.create_mate(
                mate_type=mate_type,
                component1_index=component1_index,
                component2_index=component2_index,
                face1_index=face1_index,
                face2_index=face2_index,
            )
        case "align":
            return assembly_manager.add_align_constraint(
                component1_index=component1_index,
                component2_index=component2_index,
                face1_index=face1_index,
                face2_index=face2_index,
            )
        case "planar_align":
            return assembly_manager.add_planar_align_constraint(
                component1_index=component1_index,
                component2_index=component2_index,
                face1_index=face1_index,
                face2_index=face2_index,
            )
        case "axial_align":
            return assembly_manager.add_axial_align_constraint(
                component1_index=component1_index,
                component2_index=component2_index,
                face1_index=face1_index,
                face2_index=face2_index,
            )
        case "angle":
            return assembly_manager.add_angle_constraint(
                component1_index=component1_index, component2_index=component2_index, angle=angle
            )
        case _:
            return {"error": f"Unknown type: {type}"}


# ================================================================
# Group 78: add_assembly_relation (6 -> 1)
# ================================================================


def add_assembly_relation(
    type: Literal["planar", "axial", "angular", "point", "tangent", "gear"],
    occurrence1_index: int = 0,
    occurrence2_index: int = 0,
    offset: float = 0.0,
    orientation: Literal["Align", "Antialign", "NotSpecified"] = "Align",
    angle: float = 0.0,
    ratio1: float = 1.0,
    ratio2: float = 1.0,
    face1_index: int = 0,
    face2_index: int = 0,
) -> dict[str, Any]:
    """Add a 3D relation between two occurrences (0-based indices).

    planar and axial take face1_index/face2_index, the 0-based faces of each
    occurrence's body (list them with the part's face query): planar wants
    planar faces, axial cylindrical ones. orientation 'Antialign' mates
    planar faces (they touch); 'Align' aligns them. offset applies to nothing
    here and must be 0. angular: angle in degrees. gear: ratio1/ratio2.
    """
    err = validate_numerics(offset=offset, angle=angle, ratio1=ratio1, ratio2=ratio2)
    if err:
        return err
    match type:
        case "planar":
            return assembly_manager.add_planar_relation(
                occurrence1_index=occurrence1_index,
                occurrence2_index=occurrence2_index,
                offset=offset,
                orientation=orientation,
                face1_index=face1_index,
                face2_index=face2_index,
            )
        case "axial":
            return assembly_manager.add_axial_relation(
                occurrence1_index=occurrence1_index,
                occurrence2_index=occurrence2_index,
                orientation=orientation,
                face1_index=face1_index,
                face2_index=face2_index,
            )
        case "angular":
            return assembly_manager.add_angular_relation(
                occurrence1_index=occurrence1_index,
                occurrence2_index=occurrence2_index,
                angle=angle,
            )
        case "point":
            return assembly_manager.add_point_relation(
                occurrence1_index=occurrence1_index, occurrence2_index=occurrence2_index
            )
        case "tangent":
            return assembly_manager.add_tangent_relation(
                occurrence1_index=occurrence1_index, occurrence2_index=occurrence2_index
            )
        case "gear":
            return assembly_manager.add_gear_relation(
                occurrence1_index=occurrence1_index,
                occurrence2_index=occurrence2_index,
                ratio1=ratio1,
                ratio2=ratio2,
            )
        case _:
            return {"error": f"Unknown type: {type}"}


# ================================================================
# Group 79: manage_relation (13 -> 1)
# ================================================================


def manage_relation(
    action: Literal[
        "list",
        "info",
        "delete",
        "get_offset",
        "set_offset",
        "get_angle",
        "set_angle",
        "get_normals",
        "set_normals",
        "suppress",
        "unsuppress",
        "get_geometry",
        "get_gear_ratio",
    ],
    relation_index: int = 0,
    offset: float = 0.0,
    angle: float = 0.0,
    aligned: bool = True,
) -> dict[str, Any]:
    """Inspect or edit assembly relations (0-based relation_index; list ignores it).

    set_offset: offset (meters). set_angle: angle (degrees). set_normals: aligned.
    delete removes the relation.
    """
    err = validate_numerics(offset=offset, angle=angle)
    if err:
        return err
    match action:
        case "list":
            return assembly_manager.get_assembly_relations()
        case "info":
            return assembly_manager.get_relation_info(relation_index=relation_index)
        case "delete":
            return assembly_manager.delete_relation(relation_index=relation_index)
        case "get_offset":
            return assembly_manager.get_relation_offset(relation_index=relation_index)
        case "set_offset":
            return assembly_manager.set_relation_offset(
                relation_index=relation_index, offset=offset
            )
        case "get_angle":
            return assembly_manager.get_relation_angle(relation_index=relation_index)
        case "set_angle":
            return assembly_manager.set_relation_angle(relation_index=relation_index, angle=angle)
        case "get_normals":
            return assembly_manager.get_normals_aligned(relation_index=relation_index)
        case "set_normals":
            return assembly_manager.set_normals_aligned(
                relation_index=relation_index, aligned=aligned
            )
        case "suppress":
            return assembly_manager.suppress_relation(relation_index=relation_index)
        case "unsuppress":
            return assembly_manager.unsuppress_relation(relation_index=relation_index)
        case "get_geometry":
            return assembly_manager.get_relation_geometry(relation_index=relation_index)
        case "get_gear_ratio":
            return assembly_manager.get_gear_ratio(relation_index=relation_index)
        case _:
            return {"error": f"Unknown action: {action}"}


# ================================================================
# Group 80: assembly_feature (9 -> 1)
# ================================================================


def assembly_feature(
    type: Literal[
        "extruded_cutout",
        "revolved_cutout",
        "hole",
        "extruded_protrusion",
        "revolved_protrusion",
        "mirror",
        "pattern",
        "swept_protrusion",
        "recompute",
    ],
    scope_parts: list[int] | None = None,
    extent_type: Literal["Finite", "ThroughAll"] = "Finite",
    extent_side: Literal["OneSide", "BothSides"] = "OneSide",
    profile_side: Literal["Left", "Right", "Symmetric"] = "Left",
    distance: float = 0.01,
    angle: float = 360.0,
    depth: float = 0.01,
    diameter: float = 0.006,
    feature_indices: list[int] | None = None,
    plane_index: int = 1,
    mirror_type: int = 1,
    pattern_type: Literal["Rectangular", "Circular"] = "Rectangular",
    num_trace_curves: int = 1,
    num_cross_sections: int = 1,
    options: int = 0,
) -> dict[str, Any]:
    """Create assembly-level features from the active closed sketch.

    Cutouts/hole: scope_parts = 0-based occurrence indices to cut.
    extruded_*: extent_type/extent_side/profile_side + distance (meters).
    revolved_*: angle (degrees). hole: depth and diameter (meters).
    swept_protrusion: num_trace_curves/num_cross_sections.
    recompute: options (raw COM flags, 0=default).
    'mirror'/'pattern' unsupported: AssemblyFeaturesMirrors.Add and
    AssemblyFeaturesPatterns.Add return E_ACCESSDENIED on SE 2025/2026.
    Mirror or pattern in the part document, or use manage_component.
    """
    err = validate_numerics(distance=distance, angle=angle, depth=depth)
    if err:
        return err
    match type:
        case "extruded_cutout":
            return assembly_manager.create_assembly_extruded_cutout(
                scope_parts=scope_parts or [],
                extent_type=extent_type,
                extent_side=extent_side,
                profile_side=profile_side,
                distance=distance,
            )
        case "revolved_cutout":
            return assembly_manager.create_assembly_revolved_cutout(
                scope_parts=scope_parts or [],
                extent_type=extent_type,
                extent_side=extent_side,
                profile_side=profile_side,
                angle=angle,
            )
        case "hole":
            return assembly_manager.create_assembly_hole(
                scope_parts=scope_parts or [],
                extent_type=extent_type,
                extent_side=extent_side,
                depth=depth,
                diameter=diameter,
            )
        case "extruded_protrusion":
            return assembly_manager.create_assembly_extruded_protrusion(
                extent_type=extent_type,
                extent_side=extent_side,
                profile_side=profile_side,
                distance=distance,
            )
        case "revolved_protrusion":
            return assembly_manager.create_assembly_revolved_protrusion(
                extent_type=extent_type,
                extent_side=extent_side,
                profile_side=profile_side,
                angle=angle,
            )
        case "mirror":
            return assembly_manager.create_assembly_mirror(
                feature_indices=feature_indices or [],
                plane_index=plane_index,
                mirror_type=mirror_type,
            )
        case "pattern":
            return assembly_manager.create_assembly_pattern(
                feature_indices=feature_indices or [], pattern_type=pattern_type
            )
        case "swept_protrusion":
            return assembly_manager.create_assembly_swept_protrusion(
                num_trace_curves=num_trace_curves, num_cross_sections=num_cross_sections
            )
        case "recompute":
            return assembly_manager.recompute_assembly_features(options=options)
        case _:
            return {"error": f"Unknown type: {type}"}


# ================================================================
# Group 81: virtual_component (3 -> 1)
# ================================================================


def virtual_component(
    method: Literal["new", "predefined", "bidm"] = "new",
    name: str = "",
    component_type: Literal["Part", "Assembly", "Sheetmetal", "Unknown"] = "Part",
    filename: str = "",
    doc_number: str = "",
    revision_id: str = "",
) -> dict[str, Any]:
    """Add a virtual (placeholder) component to the assembly.

    new: name + component_type. predefined: filename of an existing file.
    bidm: doc_number + revision_id + component_type (managed documents).
    """
    if method == "predefined" and filename:
        filename, err = validate_path(filename, must_exist=True)
        if err:
            return err
    match method:
        case "new":
            return assembly_manager.add_virtual_component(name=name, component_type=component_type)
        case "predefined":
            return assembly_manager.add_virtual_component_predefined(filename=filename)
        case "bidm":
            return assembly_manager.add_virtual_component_bidm(
                doc_number=doc_number, revision_id=revision_id, component_type=component_type
            )
        case _:
            return {"error": f"Unknown method: {method}"}


# ================================================================
# Group 82: structural_frame (2 -> 1)
# ================================================================


def structural_frame(
    method: Literal["basic", "by_orientation"] = "basic",
    part_filename: str = "",
    path_indices: list[int] | None = None,
    coord_system_name: str = "",
) -> dict[str, Any]:
    """Create a structural frame from part_filename along path_indices (0-based).

    by_orientation additionally takes coord_system_name.
    """
    if part_filename:
        part_filename, err = validate_path(part_filename, must_exist=True)
        if err:
            return err
    match method:
        case "basic":
            return assembly_manager.add_structural_frame(
                part_filename=part_filename, path_indices=path_indices or []
            )
        case "by_orientation":
            return assembly_manager.add_structural_frame_by_orientation(
                part_filename=part_filename,
                coord_system_name=coord_system_name,
                path_indices=path_indices or [],
            )
        case _:
            return {"error": f"Unknown method: {method}"}


# ================================================================
# Group 83: wiring (4 -> 1)
# ================================================================


def wiring(
    type: Literal["wire", "cable", "bundle", "splice"] = "wire",
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
) -> dict[str, Any]:
    """Create wire-harness elements in the assembly.

    wire: path_indices + path_directions. cable: + wire_indices.
    bundle: + conductor_indices. Both may take split_path_indices/directions.
    splice: x,y,z (meters) + conductor_indices. All indices 0-based.
    """
    err = validate_numerics(x=x, y=y, z=z)
    if err:
        return err
    match type:
        case "wire":
            return assembly_manager.add_wire(
                path_indices=path_indices or [],
                path_directions=path_directions or [],
                description=description,
            )
        case "cable":
            return assembly_manager.add_cable(
                path_indices=path_indices or [],
                path_directions=path_directions or [],
                wire_indices=wire_indices or [],
                split_path_indices=split_path_indices,
                split_path_directions=split_path_directions,
                description=description,
            )
        case "bundle":
            return assembly_manager.add_bundle(
                path_indices=path_indices or [],
                path_directions=path_directions or [],
                conductor_indices=conductor_indices or [],
                split_path_indices=split_path_indices,
                split_path_directions=split_path_directions,
                description=description,
            )
        case "splice":
            return assembly_manager.add_splice(
                x=x, y=y, z=z, conductor_indices=conductor_indices or [], description=description
            )
        case _:
            return {"error": f"Unknown type: {type}"}


# ================================================================
# Registration
# ================================================================


def register(mcp: Any) -> None:
    """Register assembly tools with the MCP server."""
    tags = {"assembly"}
    register_tool(mcp, add_assembly_component, tags=tags)
    register_tool(mcp, manage_component, tags=tags, destructive=True)
    register_tool(mcp, query_component, tags=tags, read_only=True)
    register_tool(mcp, set_component_appearance, tags=tags, idempotent=True)
    register_tool(mcp, transform_component, tags=tags)
    register_tool(mcp, set_component_orientation, tags=tags, idempotent=True)
    register_tool(mcp, rotate_component, tags=tags)
    register_tool(mcp, add_assembly_constraint, tags=tags)
    register_tool(mcp, add_assembly_relation, tags=tags)
    register_tool(mcp, manage_relation, tags=tags, destructive=True)
    register_tool(mcp, assembly_feature, tags=tags)
    register_tool(mcp, virtual_component, tags=tags)
    register_tool(mcp, structural_frame, tags=tags)
    register_tool(mcp, wiring, tags=tags)
