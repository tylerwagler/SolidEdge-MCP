"""Miscellaneous feature tools (thicken, pattern, mirror, thin wall, etc.)."""

from solidedge_mcp.managers import feature_manager


def thicken(
    method: str = "basic",
    thickness: float = 0.0,
    direction: str = "Both",
) -> dict:
    """Thicken a surface to create a solid.

    method: 'basic' | 'sync'

    Parameters:
        thickness: Thicken distance.
        direction: 'Both'|'Normal'|'Reverse'.
    """
    match method:
        case "basic":
            return feature_manager.thicken_surface(thickness, direction)
        case "sync":
            return feature_manager.create_thicken_sync(thickness, direction)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_pattern(
    method: str = "rectangular_ex",
    feature_index: int = 0,
    feature_name: str = "",
    x_count: int = 1,
    y_count: int = 1,
    x_gap: float = 0.0,
    y_gap: float = 0.0,
    x_spacing: float = 0.0,
    y_spacing: float = 0.0,
    count: int = 1,
    angle: float = 0.0,
    radius: float = 0.0,
    axis_face_index: int = 0,
    fill_region_face_index: int = 0,
    x_offsets: list[float] | None = None,
    y_offsets: list[float] | None = None,
    curve_edge_index: int = 0,
    spacing: float = 0.0,
) -> dict:
    """Create a pattern of a feature.

    method: 'rectangular_ex' | 'rectangular' | 'circular'
        | 'circular_ex' | 'duplicate' | 'by_fill'
        | 'by_table' | 'by_table_sync' | 'by_fill_ex'
        | 'by_curve_ex' | 'user_defined'

    Parameters (used per method):
        feature_index: 0-based feature index (rectangular,
            circular).
        feature_name: Feature name string (rectangular_ex,
            circular_ex, duplicate, by_fill, by_table,
            by_table_sync, by_fill_ex, by_curve_ex,
            user_defined).
        x_count, y_count: Pattern counts (rectangular,
            rectangular_ex).
        x_gap, y_gap: Gap between instances (rectangular).
        x_spacing, y_spacing: Spacing (rectangular_ex,
            by_fill, by_fill_ex).
        count: Instance count (circular, circular_ex,
            by_curve_ex).
        angle: Pattern angle (circular, circular_ex).
        radius: Pattern radius (circular).
        axis_face_index: Axis face (circular_ex).
        fill_region_face_index: Fill region face (by_fill,
            by_fill_ex).
        x_offsets, y_offsets: Offset lists (by_table,
            by_table_sync).
        curve_edge_index: Curve edge (by_curve_ex).
        spacing: Spacing along curve (by_curve_ex).
    """
    match method:
        case "rectangular_ex":
            return feature_manager.create_pattern_rectangular_ex(
                feature_name,
                x_count,
                y_count,
                x_spacing,
                y_spacing,
            )
        case "rectangular":
            return feature_manager.create_pattern_rectangular(
                feature_index,
                x_count,
                y_count,
                x_gap,
                y_gap,
            )
        case "circular":
            return feature_manager.create_pattern_circular(feature_index, count, angle, radius)
        case "circular_ex":
            return feature_manager.create_pattern_circular_ex(
                feature_name,
                count,
                angle,
                axis_face_index,
            )
        case "duplicate":
            return feature_manager.create_pattern_duplicate(feature_name)
        case "by_fill":
            return feature_manager.create_pattern_by_fill(
                feature_name,
                fill_region_face_index,
                x_spacing,
                y_spacing,
            )
        case "by_table":
            return feature_manager.create_pattern_by_table(feature_name, x_offsets, y_offsets)
        case "by_table_sync":
            return feature_manager.create_pattern_by_table_sync(feature_name, x_offsets, y_offsets)
        case "by_fill_ex":
            return feature_manager.create_pattern_by_fill_ex(
                feature_name,
                fill_region_face_index,
                x_spacing,
                y_spacing,
            )
        case "by_curve_ex":
            return feature_manager.create_pattern_by_curve_ex(
                feature_name,
                curve_edge_index,
                count,
                spacing,
            )
        case "user_defined":
            return feature_manager.create_user_defined_pattern(feature_name)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_mirror(
    method: str = "basic",
    feature_name: str = "",
    mirror_plane_index: int = 0,
    new_file_name: str = "",
    link_to_original: bool = True,
) -> dict:
    """Mirror a feature across a reference plane, or save as mirror part.

    method: 'basic' | 'sync_ex' | 'save_as_part'

    Parameters (used per method):
        feature_name: Name of the feature to mirror (basic,
            sync_ex).
        mirror_plane_index: 1-based ref plane index (basic,
            sync_ex, save_as_part).
        new_file_name: Full path for mirrored part .par
            (save_as_part).
        link_to_original: Maintain link to original
            (save_as_part).
    """
    match method:
        case "basic":
            return feature_manager.create_mirror(feature_name, mirror_plane_index)
        case "sync_ex":
            return feature_manager.create_mirror_sync_ex(feature_name, mirror_plane_index)
        case "save_as_part":
            return feature_manager.save_as_mirror_part(
                new_file_name, mirror_plane_index or 3, link_to_original
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def create_thin_wall(
    method: str = "basic",
    thickness: float = 0.0,
    open_face_indices: list[int] | None = None,
) -> dict:
    """Convert a solid body to a thin wall (shell).

    method: 'basic' | 'with_open_faces'

    Parameters:
        thickness: Shell wall thickness.
        open_face_indices: Faces to leave open
            (with_open_faces).
    """
    match method:
        case "basic":
            return feature_manager.create_shell(thickness)
        case "with_open_faces":
            return feature_manager.create_shell(thickness, open_face_indices)
        case _:
            return {"error": f"Unknown method: {method}"}


def face_operation(
    type: str = "rotate_by_points",
    face_index: int = 0,
    vertex1_index: int = 0,
    vertex2_index: int = 0,
    edge_index: int = 0,
    angle: float = 0.0,
) -> dict:
    """Perform a face operation (rotate).

    type: 'rotate_by_points' | 'rotate_by_edge'

    Parameters (used per type):
        face_index: Face to rotate.
        vertex1_index, vertex2_index: Axis vertices
            (rotate_by_points).
        edge_index: Axis edge (rotate_by_edge).
        angle: Rotation angle in degrees.
    """
    match type:
        case "rotate_by_points":
            return feature_manager.create_face_rotate_by_points(
                face_index,
                vertex1_index,
                vertex2_index,
                angle,
            )
        case "rotate_by_edge":
            return feature_manager.create_face_rotate_by_edge(face_index, edge_index, angle)
        case _:
            return {"error": f"Unknown type: {type}"}


def add_body(
    method: str = "basic",
    body_type: str = "Solid",
    tag: str = "",
) -> dict:
    """Add a body to the part.

    method: 'basic' | 'by_mesh' | 'feature' | 'construction'
        | 'by_tag'

    Parameters (used per method):
        body_type: Body type string (basic).
        tag: Tag reference string (by_tag).
    """
    match method:
        case "basic":
            return feature_manager.add_body(body_type)
        case "by_mesh":
            return feature_manager.add_body_by_mesh()
        case "feature":
            return feature_manager.add_body_feature()
        case "construction":
            return feature_manager.add_by_construction()
        case "by_tag":
            return feature_manager.add_body_by_tag(tag)
        case _:
            return {"error": f"Unknown method: {method}"}


def simplify(
    method: str = "auto",
) -> dict:
    """Simplify the model.

    method: 'auto' | 'enclosure' | 'duplicate'
        | 'local_enclosure'
    """
    match method:
        case "auto":
            return feature_manager.auto_simplify()
        case "enclosure":
            return feature_manager.simplify_enclosure()
        case "duplicate":
            return feature_manager.simplify_duplicate()
        case "local_enclosure":
            return feature_manager.local_simplify_enclosure()
        case _:
            return {"error": f"Unknown method: {method}"}


def manage_feature(
    action: str = "delete",
    index: int = 0,
    target_index: int = 0,
    after: bool = True,
    new_name: str = "",
    feature_name: str = "",
    target_type: str = "",
) -> dict:
    """Manage features in the feature tree.

    action: 'delete' | 'suppress' | 'unsuppress' | 'reorder'
        | 'rename' | 'convert'

    Parameters (used per action):
        index: Feature index (delete, suppress, unsuppress,
            reorder, rename).
        target_index: Target position index (reorder).
        after: Insert after target (reorder).
        new_name: New feature name (rename).
        feature_name: Feature name (convert).
        target_type: Target type string (convert).
    """
    match action:
        case "delete":
            return feature_manager.delete_feature(index)
        case "suppress":
            return feature_manager.feature_suppress(index)
        case "unsuppress":
            return feature_manager.feature_unsuppress(index)
        case "reorder":
            return feature_manager.feature_reorder(index, target_index, after)
        case "rename":
            return feature_manager.feature_rename(index, new_name)
        case "convert":
            return feature_manager.convert_feature_type(feature_name, target_type)
        case _:
            return {"error": f"Unknown action: {action}"}


def create_draft_angle(
    face_index: int,
    angle: float,
    plane_index: int = 1,
) -> dict:
    """Add a draft angle to a face."""
    return feature_manager.create_draft_angle(face_index, angle, plane_index)
