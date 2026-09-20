"""Round, chamfer, blend, and topology deletion tools."""

from typing import Any, Literal

from solidedge_mcp.backends.validation import validate_numerics
from solidedge_mcp.managers import feature_manager


def create_round(
    method: Literal["all_edges", "on_face", "variable", "blend", "surface_blend"] = "all_edges",
    radius: float = 0.0,
    face_index: int | None = None,
    radii: list[float] | None = None,
    face_index1: int = 0,
    face_index2: int = 0,
) -> dict[str, Any]:
    """Round (fillet) edges of the active solid body.

    radius and radii in meters; all face indices are 0-based.
    all_edges: radius on every edge. on_face: radius on the edges of
    face_index. blend / surface_blend: radius between face_index1 and
    face_index2. 'variable' (unsupported: Rounds.AddVariable needs a
    VertexArray naming the vertices each radius applies to, which cannot be
    selected here); use a constant-radius method.
    """
    err = validate_numerics(radius=radius)
    if err:
        return err
    match method:
        case "all_edges":
            return feature_manager.create_round(radius=radius)
        case "on_face":
            return feature_manager.create_round_on_face(radius=radius, face_index=face_index or 0)
        case "variable":
            return feature_manager.create_variable_round(radii=radii or [], face_index=face_index)
        case "blend":
            return feature_manager.create_round_blend(
                face_index1=face_index1, face_index2=face_index2, radius=radius
            )
        case "surface_blend":
            return feature_manager.create_round_surface_blend(
                face_index1=face_index1, face_index2=face_index2, radius=radius
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def create_chamfer(
    method: Literal["equal", "on_face", "unequal", "unequal_on_face", "angle"] = "equal",
    distance: float = 0.0,
    face_index: int = 0,
    distance1: float = 0.0,
    distance2: float = 0.0,
    angle: float = 0.0,
) -> dict[str, Any]:
    """Chamfer edges of the active solid body.

    Distances in meters; angle in degrees. face_index is 0-based and selects
    the face whose edges are chamfered (ignored by 'equal', which chamfers all
    edges). distance: equal / on_face / angle. distance1 and distance2:
    unequal / unequal_on_face. angle: 'angle' method only.
    """
    err = validate_numerics(
        distance=distance,
        distance1=distance1,
        distance2=distance2,
        angle=angle,
    )
    if err:
        return err
    match method:
        case "equal":
            return feature_manager.create_chamfer(distance=distance)
        case "on_face":
            return feature_manager.create_chamfer_on_face(distance=distance, face_index=face_index)
        case "unequal":
            return feature_manager.create_chamfer_unequal(
                distance1=distance1, distance2=distance2, face_index=face_index
            )
        case "unequal_on_face":
            return feature_manager.create_chamfer_unequal_on_face(
                distance1=distance1, distance2=distance2, face_index=face_index
            )
        case "angle":
            return feature_manager.create_chamfer_angle(
                distance=distance, angle=angle, face_index=face_index
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def create_blend(
    method: Literal["basic", "variable", "surface"] = "basic",
    radius: float = 0.0,
    face_index: int | None = None,
    radius1: float = 0.0,
    radius2: float = 0.0,
    face_index1: int = 0,
    face_index2: int = 0,
) -> dict[str, Any]:
    """Create a blend (face-to-face fillet).

    Radii in meters; all face indices are 0-based.
    basic: radius on face_index. surface: radius (required, must be > 0)
    between face_index1 and face_index2. 'variable' (unsupported:
    Blends.AddVariable needs a VertexArray naming the vertices each radius
    applies to, which cannot be selected here); use 'basic' or 'surface'.
    """
    err = validate_numerics(radius=radius, radius1=radius1, radius2=radius2)
    if err:
        return err
    match method:
        case "basic":
            return feature_manager.create_blend(radius=radius, face_index=face_index)
        case "variable":
            return feature_manager.create_blend_variable(
                radius1=radius1, radius2=radius2, face_index=face_index
            )
        case "surface":
            return feature_manager.create_blend_surface(
                face_index1=face_index1, face_index2=face_index2, radius=radius
            )
        case _:
            return {"error": f"Unknown method: {method}"}


def delete_topology(
    type: Literal["hole", "hole_by_face", "blend", "faces"] = "hole",
    max_diameter: float = 1.0,
    hole_type: Literal["All", "Round", "NonRound"] = "All",
    face_index: int = 0,
    face_indices: list[int] | None = None,
) -> dict[str, Any]:
    """Remove topology from the body: fill holes, remove blends or faces.

    max_diameter in meters; face indices are 0-based. This deletes geometry.
    hole: fills every hole of hole_type up to max_diameter.
    hole_by_face: fills the hole owning face_index.
    blend: removes the blend on face_index.
    'faces' (unsupported: DeleteFaces.Add takes one FaceSet object and the
    Part API exposes no way to build a FaceSet from face indices); delete the
    faces in the Solid Edge UI.
    """
    err = validate_numerics(max_diameter=max_diameter)
    if err:
        return err
    match type:
        case "hole":
            return feature_manager.create_delete_hole(
                max_diameter=max_diameter, hole_type=hole_type
            )
        case "hole_by_face":
            return feature_manager.delete_hole_by_face(face_index=face_index)
        case "blend":
            return feature_manager.create_delete_blend(face_index=face_index)
        case "faces":
            return feature_manager.delete_faces(face_indices=face_indices or [])
        case _:
            return {"error": f"Unknown type: {type}"}
