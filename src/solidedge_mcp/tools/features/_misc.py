"""Miscellaneous feature tools (thicken, pattern, mirror, face ops, etc.)."""

from typing import Any, Literal

from solidedge_mcp.backends.validation import validate_numerics
from solidedge_mcp.managers import feature_manager


def thicken(
    method: Literal["basic", "sync"] = "basic",
    thickness: float = 0.0,
    direction: str = "Both",
) -> dict[str, Any]:
    """Thicken an existing surface body into a solid.

    thickness in meters. direction applies to 'sync' only and accepts
    'Both' | 'Normal' | 'Reverse'; the 'basic' backend records it but the COM
    call ignores it (thickens both sides).
    """
    err = validate_numerics(thickness=thickness)
    if err:
        return err
    match method:
        case "basic":
            return feature_manager.thicken_surface(thickness, direction)
        case "sync":
            return feature_manager.create_thicken_sync(thickness, direction)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_pattern(
    method: Literal[
        "rectangular_ex",
        "circular_ex",
        "duplicate",
        "by_fill",
        "by_table",
        "by_table_sync",
        "by_fill_ex",
        "by_curve_ex",
        "user_defined",
    ] = "rectangular_ex",
    feature_name: str = "",
    x_count: int = 1,
    y_count: int = 1,
    x_spacing: float = 0.0,
    y_spacing: float = 0.0,
    count: int = 1,
    angle: float = 0.0,
    axis_face_index: int = 0,
    fill_region_face_index: int = 0,
    x_offsets: list[float] | None = None,
    y_offsets: list[float] | None = None,
    curve_edge_index: int = 0,
    spacing: float = 0.0,
) -> dict[str, Any]:
    """Pattern an existing feature, selected by name (see list_features).

    Spacings/offsets in meters, angle in degrees. Face/edge indices 0-based.
    x_count/y_count/x_spacing/y_spacing: rectangular_ex. x_spacing/y_spacing
    also size the grid for by_fill/by_fill_ex (with fill_region_face_index).
    count/angle/axis_face_index: circular_ex. x_offsets/y_offsets: by_table*.
    curve_edge_index/count/spacing: by_curve_ex. duplicate and user_defined
    take only feature_name.
    """
    err = validate_numerics(
        x_spacing=x_spacing,
        y_spacing=y_spacing,
        angle=angle,
        spacing=spacing,
    )
    if err:
        return err
    match method:
        case "rectangular_ex":
            return feature_manager.create_pattern_rectangular_ex(
                feature_name,
                x_count,
                y_count,
                x_spacing,
                y_spacing,
            )
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
            return feature_manager.create_pattern_by_table(
                feature_name, x_offsets or [], y_offsets or []
            )
        case "by_table_sync":
            return feature_manager.create_pattern_by_table_sync(
                feature_name, x_offsets or [], y_offsets or []
            )
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
    method: Literal["basic", "sync_ex", "save_as_part"] = "basic",
    feature_name: str = "",
    mirror_plane_index: int = 3,
    new_file_name: str = "",
    link_to_original: bool = True,
) -> dict[str, Any]:
    """Mirror a feature across a reference plane, or save a mirrored part.

    mirror_plane_index is 1-based (1=Top/XY, 2=Right/YZ, 3=Front/XZ, 4+ = user
    planes). feature_name (see list_features): basic/sync_ex. new_file_name
    (absolute .par path) and link_to_original: save_as_part.
    """
    if mirror_plane_index < 1:
        return {
            "error": (
                f"mirror_plane_index must be >= 1 (got {mirror_plane_index}); "
                "plane indices are 1-based (1=Top/XY, 2=Right/YZ, 3=Front/XZ)."
            )
        }
    match method:
        case "basic":
            return feature_manager.create_mirror(feature_name, mirror_plane_index)
        case "sync_ex":
            return feature_manager.create_mirror_sync_ex(feature_name, mirror_plane_index)
        case "save_as_part":
            return feature_manager.save_as_mirror_part(
                new_file_name, mirror_plane_index, link_to_original
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def face_operation(
    type: Literal["rotate_by_points", "rotate_by_edge"] = "rotate_by_points",
    face_index: int = 0,
    vertex1_index: int = 0,
    vertex2_index: int = 0,
    edge_index: int = 0,
    angle: float = 0.0,
) -> dict[str, Any]:
    """Rotate a face of the solid body about an axis.

    angle in degrees. All indices are 0-based. rotate_by_points takes the axis
    from vertex1_index/vertex2_index; rotate_by_edge takes it from edge_index.
    """
    err = validate_numerics(angle=angle)
    if err:
        return err
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
    method: Literal["basic", "by_mesh", "feature", "construction", "by_tag"] = "basic",
    body_type: str = "Solid",
    tag: str = "",
) -> dict[str, Any]:
    """Add a new body to the part document.

    body_type ('Solid' | 'Surface' | 'Construction') is echoed back but the COM
    AddBody call takes no type argument, so it does not change the result.
    tag: by_tag only.
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
    method: Literal["auto", "enclosure", "duplicate", "local_enclosure"] = "auto",
) -> dict[str, Any]:
    """Simplify the model for downstream use (lighter assembly representation).

    auto: automatic simplification. enclosure / local_enclosure: replace the
    body (or a local region) with its bounding enclosure. duplicate: simplify
    by removing duplicate geometry. Takes no other parameters.
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
    action: Literal["delete", "suppress", "unsuppress", "reorder", "rename", "convert"] = "delete",
    index: int = 0,
    target_index: int = 0,
    after: bool = True,
    new_name: str = "",
    feature_name: str = "",
    target_type: str = "",
) -> dict[str, Any]:
    """Manage entries in the feature tree (delete, suppress, reorder, rename).

    index and target_index are 0-based positions in the feature tree (see
    list_features). reorder moves 'index' to 'target_index' ('after' places it
    behind the target). rename uses new_name. convert selects the feature by
    feature_name and needs target_type: 'cutout' or 'protrusion'.
    Deleting a feature is destructive and removes its geometry.
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
) -> dict[str, Any]:
    """Add a draft (taper) to a face, pulled from a reference plane.

    angle in degrees. face_index is 0-based. plane_index is the 1-based draft
    (parting) plane: 1=Top/XY, 2=Right/YZ, 3=Front/XZ.
    """
    err = validate_numerics(angle=angle)
    if err:
        return err
    return feature_manager.create_draft_angle(face_index, angle, plane_index)
