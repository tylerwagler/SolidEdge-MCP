"""Round, chamfer, blend, and topology deletion tools."""

from typing import Any

from solidedge_mcp.backends.validation import validate_numerics
from solidedge_mcp.managers import feature_manager


def create_round(
    method: str = "all_edges",
    radius: float = 0.0,
    face_index: int | None = None,
    radii: list[float] | None = None,
    face_index1: int = 0,
    face_index2: int = 0,
) -> dict[str, Any]:
    """Round (fillet) edges of the active body.

    method: 'all_edges' | 'on_face' | 'variable' | 'blend'
        | 'surface_blend'

    radius/radii in meters.
    """
    err = validate_numerics(radius=radius)
    if err:
        return err
    match method:
        case "all_edges":
            return feature_manager.create_round(radius)
        case "on_face":
            return feature_manager.create_round_on_face(radius, face_index or 0)
        case "variable":
            return feature_manager.create_variable_round(radii or [], face_index)
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
) -> dict[str, Any]:
    """Chamfer edges of the active body.

    method: 'equal' | 'on_face' | 'unequal'
        | 'unequal_on_face' | 'angle'

    Distances in meters. angle in degrees.
    """
    err = validate_numerics(
        distance=distance, distance1=distance1,
        distance2=distance2, angle=angle,
    )
    if err:
        return err
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
) -> dict[str, Any]:
    """Create a blend (face-to-face fillet).

    method: 'basic' | 'variable' | 'surface'

    Radii in meters.
    """
    err = validate_numerics(radius=radius, radius1=radius1, radius2=radius2)
    if err:
        return err
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
) -> dict[str, Any]:
    """Delete topology features (holes, blends, faces).

    type: 'hole' | 'hole_by_face' | 'blend' | 'faces'

    max_diameter in meters.
    """
    err = validate_numerics(max_diameter=max_diameter)
    if err:
        return err
    match type:
        case "hole":
            return feature_manager.create_delete_hole(max_diameter, hole_type)
        case "hole_by_face":
            return feature_manager.delete_hole_by_face(face_index)
        case "blend":
            return feature_manager.create_delete_blend(face_index)
        case "faces":
            return feature_manager.delete_faces(face_indices or [])
        case _:
            return {"error": f"Unknown type: {type}"}
