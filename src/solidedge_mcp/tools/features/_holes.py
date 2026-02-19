"""Hole creation tools."""

from solidedge_mcp.managers import feature_manager


def create_hole(
    method: str = "finite",
    x: float = 0.0,
    y: float = 0.0,
    diameter: float = 0.0,
    depth: float = 0.0,
    direction: str = "Normal",
    plane_index: int = 1,
    from_plane_index: int = 0,
    to_plane_index: int = 0,
) -> dict:
    """Create a hole at coordinates (meters).

    method: 'finite' | 'through_all' | 'from_to'
        | 'through_next' | 'sync' | 'finite_ex'
        | 'from_to_ex' | 'through_next_ex'
        | 'through_all_ex' | 'sync_ex' | 'multi_body'
        | 'sync_multi_body'

    Parameters (used per method):
        x, y: Hole center coordinates.
        diameter: Hole diameter.
        depth: Hole depth (finite, sync, finite_ex, sync_ex,
            multi_body, sync_multi_body).
        direction: 'Normal'|'Reverse' (finite, through_next,
            finite_ex, through_next_ex, multi_body).
        plane_index: 1-based ref plane index.
        from_plane_index: 1-based ref plane (from_to,
            from_to_ex).
        to_plane_index: 1-based ref plane (from_to,
            from_to_ex).
    """
    match method:
        case "finite":
            return feature_manager.create_hole(x, y, diameter, depth, direction)
        case "through_all":
            return feature_manager.create_hole_through_all(x, y, diameter, plane_index, direction)
        case "from_to":
            return feature_manager.create_hole_from_to(
                x,
                y,
                diameter,
                from_plane_index,
                to_plane_index,
            )
        case "through_next":
            return feature_manager.create_hole_through_next(x, y, diameter, direction, plane_index)
        case "sync":
            return feature_manager.create_hole_sync(x, y, diameter, depth, plane_index)
        case "finite_ex":
            return feature_manager.create_hole_finite_ex(
                x, y, diameter, depth, direction, plane_index
            )
        case "from_to_ex":
            return feature_manager.create_hole_from_to_ex(
                x,
                y,
                diameter,
                from_plane_index,
                to_plane_index,
            )
        case "through_next_ex":
            return feature_manager.create_hole_through_next_ex(
                x, y, diameter, direction, plane_index
            )
        case "through_all_ex":
            return feature_manager.create_hole_through_all_ex(x, y, diameter, plane_index)
        case "sync_ex":
            return feature_manager.create_hole_sync_ex(x, y, diameter, depth, plane_index)
        case "multi_body":
            return feature_manager.create_hole_multi_body(
                x, y, diameter, depth, plane_index, direction
            )
        case "sync_multi_body":
            return feature_manager.create_hole_sync_multi_body(x, y, diameter, depth, plane_index)
        case _:
            return {"error": f"Unknown method: {method}"}
