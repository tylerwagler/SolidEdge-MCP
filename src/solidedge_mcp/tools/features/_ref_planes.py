"""Reference plane creation tools."""

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
    curve_end: str = "End",
    pivot_plane_index: int = 2,
    ratio: float = 0.0,
    distance_along: float = 0.0,
    face_index: int = 0,
    keypoint_type: str = "End",
    curve_edge_index: int = 0,
    orientation_plane_index: int = 0,
    normal_side_int: int = 2,
) -> dict:
    """Create a reference plane.

    method: 'offset' | 'angle' | 'three_points' | 'midplane'
        | 'normal_to_curve' | 'normal_at_distance'
        | 'normal_at_arc_ratio' | 'normal_at_distance_along'
        | 'parallel_by_tangent' | 'normal_at_keypoint'
        | 'tangent_cylinder_angle'
        | 'tangent_cylinder_keypoint'
        | 'tangent_surface_keypoint'
        | 'normal_at_distance_v2'
        | 'normal_at_arc_ratio_v2'
        | 'normal_at_distance_along_v2'
        | 'tangent_parallel'

    Parameters (used per method):
        parent_plane_index: 1-based parent plane (offset, angle,
            parallel_by_tangent, tangent_cylinder_angle,
            tangent_cylinder_keypoint,
            tangent_surface_keypoint, tangent_parallel).
        distance: Offset distance (offset, normal_at_distance).
        normal_side: 'Normal' or 'Reverse' (offset, angle,
            parallel_by_tangent).
        angle: Angle in degrees (angle,
            tangent_cylinder_angle).
        x1..z3: 9 coordinates (three_points).
        plane1_index, plane2_index: Planes (midplane).
        curve_end: 'Start' or 'End' (normal_to_curve,
            normal_at_distance, normal_at_arc_ratio,
            normal_at_distance_along).
        pivot_plane_index: Pivot plane (normal_to_curve,
            normal_at_distance, normal_at_arc_ratio,
            normal_at_distance_along).
        ratio: Arc-length ratio 0-1 (normal_at_arc_ratio).
        distance_along: Distance along curve
            (normal_at_distance_along).
        face_index: Face index (parallel_by_tangent,
            tangent_cylinder_angle,
            tangent_cylinder_keypoint,
            tangent_surface_keypoint, tangent_parallel).
        keypoint_type: 'Start' or 'End' (normal_at_keypoint,
            tangent_cylinder_keypoint,
            tangent_surface_keypoint).
        curve_edge_index: Edge index (v2 variants).
        orientation_plane_index: Orientation plane (v2 variants).
        normal_side_int: Integer normal side (v2 variants,
            tangent_parallel).
    """
    match method:
        case "offset":
            return feature_manager.create_ref_plane_by_offset(
                parent_plane_index, distance, normal_side
            )
        case "angle":
            return feature_manager.create_ref_plane_by_angle(parent_plane_index, angle, normal_side)
        case "three_points":
            return feature_manager.create_ref_plane_by_3_points(x1, y1, z1, x2, y2, z2, x3, y3, z3)
        case "midplane":
            return feature_manager.create_ref_plane_midplane(plane1_index, plane2_index)
        case "normal_to_curve":
            return feature_manager.create_ref_plane_normal_to_curve(curve_end, pivot_plane_index)
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
                distance_along,
                curve_end,
                pivot_plane_index,
            )
        case "parallel_by_tangent":
            return feature_manager.create_ref_plane_parallel_by_tangent(
                parent_plane_index,
                face_index,
                normal_side,
            )
        case "normal_at_keypoint":
            return feature_manager.create_ref_plane_normal_at_keypoint(
                keypoint_type, pivot_plane_index
            )
        case "tangent_cylinder_angle":
            return feature_manager.create_ref_plane_tangent_cylinder_angle(
                face_index, angle, parent_plane_index
            )
        case "tangent_cylinder_keypoint":
            return feature_manager.create_ref_plane_tangent_cylinder_keypoint(
                face_index,
                keypoint_type,
                parent_plane_index,
            )
        case "tangent_surface_keypoint":
            return feature_manager.create_ref_plane_tangent_surface_keypoint(
                face_index,
                keypoint_type,
                parent_plane_index,
            )
        case "normal_at_distance_v2":
            return feature_manager.create_ref_plane_normal_at_distance_v2(
                curve_edge_index,
                orientation_plane_index,
                distance,
                normal_side_int,
            )
        case "normal_at_arc_ratio_v2":
            return feature_manager.create_ref_plane_normal_at_arc_ratio_v2(
                curve_edge_index,
                orientation_plane_index,
                ratio,
                normal_side_int,
            )
        case "normal_at_distance_along_v2":
            return feature_manager.create_ref_plane_normal_at_distance_along_v2(
                curve_edge_index,
                orientation_plane_index,
                distance_along,
                normal_side_int,
            )
        case "tangent_parallel":
            return feature_manager.create_ref_plane_tangent_parallel(
                parent_plane_index,
                face_index,
                normal_side_int,
            )
        case _:
            return {"error": f"Unknown method: {method}"}
