"""Round, chamfer, blend, and topology deletion tools."""

from solidedge_mcp.managers import feature_manager


def create_round(
    method: str = "all_edges",
    radius: float = 0.0,
    face_index: int | None = None,
    radii: list | None = None,
    face_index1: int = 0,
    face_index2: int = 0,
) -> dict:
    """Round (fillet) edges of the active body.

    method: 'all_edges' | 'on_face' | 'variable' | 'blend'
        | 'surface_blend'

    Parameters (used per method):
        radius: Fillet radius (all_edges, on_face, blend,
            surface_blend).
        face_index: Face index (on_face, variable).
        radii: List of radii (variable).
        face_index1: First face (blend, surface_blend).
        face_index2: Second face (blend, surface_blend).
    """
    match method:
        case "all_edges":
            return feature_manager.create_round(radius)
        case "on_face":
            return feature_manager.create_round_on_face(radius, face_index)
        case "variable":
            return feature_manager.create_variable_round(radii, face_index)
        case "blend":
            return feature_manager.create_round_blend(face_index1, face_index2, radius)
        case "surface_blend":
            return feature_manager.create_round_surface_blend(face_index1, face_index2, radius)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_chamfer(
    method: str = "equal",
    distance: float = 0.0,
    face_index: int = 0,
    distance1: float = 0.0,
    distance2: float = 0.0,
    angle: float = 0.0,
) -> dict:
    """Chamfer edges of the active body.

    method: 'equal' | 'on_face' | 'unequal'
        | 'unequal_on_face' | 'angle'

    Parameters (used per method):
        distance: Chamfer distance (equal, on_face, angle).
        face_index: Face index (on_face, unequal,
            unequal_on_face, angle).
        distance1: First distance (unequal, unequal_on_face).
        distance2: Second distance (unequal, unequal_on_face).
        angle: Chamfer angle in degrees (angle).
    """
    match method:
        case "equal":
            return feature_manager.create_chamfer(distance)
        case "on_face":
            return feature_manager.create_chamfer_on_face(distance, face_index)
        case "unequal":
            return feature_manager.create_chamfer_unequal(distance1, distance2, face_index)
        case "unequal_on_face":
            return feature_manager.create_chamfer_unequal_on_face(distance1, distance2, face_index)
        case "angle":
            return feature_manager.create_chamfer_angle(distance, angle, face_index)
        case _:
            return {"error": f"Unknown method: {method}"}


def create_blend(
    method: str = "basic",
    radius: float = 0.0,
    face_index: int | None = None,
    radius1: float = 0.0,
    radius2: float = 0.0,
    face_index1: int = 0,
    face_index2: int = 0,
) -> dict:
    """Create a blend (face-to-face fillet).

    method: 'basic' | 'variable' | 'surface'

    Parameters (used per method):
        radius: Blend radius (basic).
        face_index: Face index (basic, variable).
        radius1: Start radius (variable).
        radius2: End radius (variable).
        face_index1: First face (surface).
        face_index2: Second face (surface).
    """
    match method:
        case "basic":
            return feature_manager.create_blend(radius, face_index)
        case "variable":
            return feature_manager.create_blend_variable(radius1, radius2, face_index)
        case "surface":
            return feature_manager.create_blend_surface(face_index1, face_index2)
        case _:
            return {"error": f"Unknown method: {method}"}


def delete_topology(
    type: str = "hole",
    max_diameter: float = 1.0,
    hole_type: str = "All",
    face_index: int = 0,
    face_indices: list[int] | None = None,
) -> dict:
    """Delete topology features (holes, blends, faces).

    type: 'hole' | 'hole_by_face' | 'blend' | 'faces'

    Parameters (used per type):
        max_diameter: Maximum hole diameter (hole).
        hole_type: Hole type filter (hole).
        face_index: Face index (hole_by_face, blend).
        face_indices: List of face indices (faces).
    """
    match type:
        case "hole":
            return feature_manager.create_delete_hole(max_diameter, hole_type)
        case "hole_by_face":
            return feature_manager.delete_hole_by_face(face_index)
        case "blend":
            return feature_manager.create_delete_blend(face_index)
        case "faces":
            return feature_manager.delete_faces(face_indices)
        case _:
            return {"error": f"Unknown type: {type}"}
