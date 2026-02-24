"""Reference plane creation tools."""

from solidedge_mcp.backends.validation import validate_numerics
from solidedge_mcp.managers import feature_manager


def create_ref_plane(
    method: str = "offset",
    parent_plane_index: int = 1,
    distance: float = 0.0,
    normal_side: str = "Normal",
    angle: float = 0.0,
    x1: float = 0.0,
    y1: float = 0.0,
    z1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    z2: float = 0.0,
    x3: float = 0.0,
    y3: float = 0.0,
    z3: float = 0.0,
    plane1_index: int = 0,
    plane2_index: int = 0,
) -> dict:
    """Create a reference plane from basic geometric definitions.

    method: 'offset' | 'angle' | 'three_points' | 'midplane'

    distance/coordinates in meters. angle in degrees.
    parent_plane_index and plane1/2_index are 1-based.
    normal_side: 'Normal' or 'Reverse'.
    """
    err = validate_numerics(
        distance=distance, angle=angle,
        x1=x1, y1=y1, z1=z1, x2=x2, y2=y2, z2=z2,
        x3=x3, y3=y3, z3=z3,
    )
    if err:
        return err
    match method:
        case "offset":
            return feature_manager.create_ref_plane_by_offset(
                parent_plane_index, distance, normal_side
            )
        case "angle":
            return feature_manager.create_ref_plane_by_angle(
                parent_plane_index, angle, normal_side
            )
        case "three_points":
            return feature_manager.create_ref_plane_by_3_points(
                x1, y1, z1, x2, y2, z2, x3, y3, z3
            )
        case "midplane":
            return feature_manager.create_ref_plane_midplane(
                plane1_index, plane2_index
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def create_ref_plane_on_curve(
    method: str = "normal_to_curve",
    curve_end: str = "End",
    pivot_plane_index: int = 2,
    distance: float = 0.0,
    ratio: float = 0.0,
    distance_along: float = 0.0,
    keypoint_type: str = "End",
    curve_edge_index: int = 0,
    orientation_plane_index: int = 0,
    normal_side_int: int = 2,
) -> dict:
    """Create a reference plane normal to a curve or edge.

    method: 'normal_to_curve' | 'normal_at_distance' | 'normal_at_arc_ratio'
        | 'normal_at_distance_along' | 'normal_at_keypoint'
        | 'normal_at_distance_v2' | 'normal_at_arc_ratio_v2'
        | 'normal_at_distance_along_v2'

    distance/distance_along in meters. ratio is 0-1 arc-length fraction.
    curve_end/keypoint_type: 'Start' or 'End'.
    pivot_plane_index and orientation_plane_index are 1-based.
    v2 variants use curve_edge_index (0-based) and orientation_plane_index
    instead of active-sketch curve_end/pivot_plane_index.
    """
    err = validate_numerics(
        distance=distance, ratio=ratio, distance_along=distance_along,
    )
    if err:
        return err
    match method:
        case "normal_to_curve":
            return feature_manager.create_ref_plane_normal_to_curve(
                curve_end, pivot_plane_index
            )
        case "normal_at_distance":
            return feature_manager.create_ref_plane_normal_at_distance(
                distance, curve_end, pivot_plane_index
            )
        case "normal_at_arc_ratio":
            return feature_manager.create_ref_plane_normal_at_arc_ratio(
                ratio, curve_end, pivot_plane_index
            )
        case "normal_at_distance_along":
            return feature_manager.create_ref_plane_normal_at_distance_along(
                distance_along, curve_end, pivot_plane_index,
            )
        case "normal_at_keypoint":
            return feature_manager.create_ref_plane_normal_at_keypoint(
                keypoint_type, pivot_plane_index
            )
        case "normal_at_distance_v2":
            return feature_manager.create_ref_plane_normal_at_distance_v2(
                curve_edge_index, orientation_plane_index,
                distance, normal_side_int,
            )
        case "normal_at_arc_ratio_v2":
            return feature_manager.create_ref_plane_normal_at_arc_ratio_v2(
                curve_edge_index, orientation_plane_index,
                ratio, normal_side_int,
            )
        case "normal_at_distance_along_v2":
            return feature_manager.create_ref_plane_normal_at_distance_along_v2(
                curve_edge_index, orientation_plane_index,
                distance_along, normal_side_int,
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def create_ref_plane_tangent(
    method: str = "parallel_by_tangent",
    parent_plane_index: int = 1,
    face_index: int = 0,
    normal_side: str = "Normal",
    angle: float = 0.0,
    keypoint_type: str = "End",
    normal_side_int: int = 2,
) -> dict:
    """Create a reference plane tangent to a face or cylinder.

    method: 'parallel_by_tangent' | 'tangent_cylinder_angle'
        | 'tangent_cylinder_keypoint' | 'tangent_surface_keypoint'
        | 'tangent_parallel'

    parent_plane_index is 1-based. face_index is 0-based.
    angle in degrees. normal_side: 'Normal' or 'Reverse'.
    keypoint_type: 'Start' or 'End'.
    """
    err = validate_numerics(angle=angle)
    if err:
        return err
    match method:
        case "parallel_by_tangent":
            return feature_manager.create_ref_plane_parallel_by_tangent(
                parent_plane_index, face_index, normal_side,
            )
        case "tangent_cylinder_angle":
            return feature_manager.create_ref_plane_tangent_cylinder_angle(
                face_index, angle, parent_plane_index
            )
        case "tangent_cylinder_keypoint":
            return feature_manager.create_ref_plane_tangent_cylinder_keypoint(
                face_index, keypoint_type, parent_plane_index,
            )
        case "tangent_surface_keypoint":
            return feature_manager.create_ref_plane_tangent_surface_keypoint(
                face_index, keypoint_type, parent_plane_index,
            )
        case "tangent_parallel":
            return feature_manager.create_ref_plane_tangent_parallel(
                parent_plane_index, face_index, normal_side_int,
            )
        case _:
            return {"error": f"Unknown method: {method}"}
