"""Extrude protrusion tools."""

from solidedge_mcp.managers import feature_manager


def create_extrude(
    method: str = "finite",
    distance: float = 0.0,
    direction: str = "Normal",
    wall_thickness: float = 0.0,
    from_plane_index: int = 0,
    to_plane_index: int = 0,
) -> dict:
    """Create an extruded protrusion from the active sketch profile.

    method: 'finite' | 'infinite' | 'through_next' | 'from_to'
        | 'thin_wall' | 'symmetric' | 'through_next_v2'
        | 'from_to_v2' | 'by_keypoint' | 'from_to_single'
        | 'through_next_single'

    Parameters (used per method):
        distance: Extrusion distance in meters (finite, thin_wall, symmetric).
        direction: 'Normal'|'Reverse'|'Both' (finite, infinite,
            through_next, thin_wall, through_next_v2, by_keypoint,
            through_next_single).
        wall_thickness: Wall thickness in meters (thin_wall).
        from_plane_index: 1-based ref plane index (from_to, from_to_v2,
            from_to_single).
        to_plane_index: 1-based ref plane index (from_to, from_to_v2,
            from_to_single).
    """
    match method:
        case "finite":
            return feature_manager.create_extrude(distance, direction)
        case "infinite":
            return feature_manager.create_extrude_infinite(direction)
        case "through_next":
            return feature_manager.create_extrude_through_next(direction)
        case "from_to":
            return feature_manager.create_extrude_from_to(from_plane_index, to_plane_index)
        case "thin_wall":
            return feature_manager.create_extrude_thin_wall(distance, wall_thickness, direction)
        case "symmetric":
            return feature_manager.create_extrude_symmetric(distance)
        case "through_next_v2":
            return feature_manager.create_extrude_through_next_v2(direction)
        case "from_to_v2":
            return feature_manager.create_extrude_from_to_v2(from_plane_index, to_plane_index)
        case "by_keypoint":
            return feature_manager.create_extrude_by_keypoint(direction)
        case "from_to_single":
            return feature_manager.create_extrude_from_to_single(from_plane_index, to_plane_index)
        case "through_next_single":
            return feature_manager.create_extrude_through_next_single(direction)
        case _:
            return {"error": f"Unknown method: {method}"}
