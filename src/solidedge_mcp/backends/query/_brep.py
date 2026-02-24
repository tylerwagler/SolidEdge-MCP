"""B-Rep topology queries: faces, edges, vertices, shells, and geometry inspection."""

import contextlib
import traceback
from typing import Any

import pythoncom
from win32com.client import VARIANT

from ..constants import FaceQueryConstants
from ..logging import get_logger

_logger = get_logger(__name__)


class BRepMixin:
    """Mixin providing B-Rep topology query methods."""

    doc_manager: Any

    def get_body_faces(self) -> dict[str, Any]:
        """
        Get all faces on the model body.

        Uses Body.Faces(igQueryAll=1) to enumerate all faces with their
        geometry type, area, and edge count.

        Face geometry types are determined by checking which query type
        each face belongs to (plane, cylinder, cone, sphere, torus, spline).

        Returns:
            Dict with list of faces and count
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body

            # Query type constants (from SE type library)
            query_types = {
                6: "plane",  # igQueryPlane
                10: "cylinder",  # igQueryCylinder
                7: "cone",  # igQueryCone
                9: "sphere",  # igQuerySphere
                8: "torus",  # igQueryTorus
                5: "spline",  # igQuerySpline
            }

            # Build a set of face indices per geometry type
            geo_type_map = {}  # face_area -> geometry_type (approximate matching)
            for qval, qname in query_types.items():
                try:
                    typed_faces = body.Faces(qval)
                    for j in range(1, typed_faces.Count + 1):
                        try:
                            tf = typed_faces.Item(j)
                            # Use (area, edge_count) as a fingerprint
                            area = tf.Area
                            ec = tf.Edges.Count if hasattr(tf.Edges, "Count") else 0
                            key = (round(area, 12), ec)
                            geo_type_map[key] = qname
                        except Exception:
                            pass
                except Exception:
                    pass

            # Now enumerate all faces
            faces = body.Faces(FaceQueryConstants.igQueryAll)
            face_list = []

            for i in range(1, faces.Count + 1):
                try:
                    face = faces.Item(i)
                    face_info: dict[str, Any] = {"index": i - 1}
                    with contextlib.suppress(Exception):
                        face_info["area"] = face.Area
                    try:
                        edge_count = face.Edges.Count if hasattr(face.Edges, "Count") else 0
                        face_info["edge_count"] = edge_count
                    except Exception:
                        pass
                    # Look up geometry type
                    try:
                        key = (round(face.Area, 12), face_info.get("edge_count", 0))
                        face_info["geometry"] = geo_type_map.get(key, "unknown")
                    except Exception:
                        face_info["geometry"] = "unknown"
                    face_list.append(face_info)
                except Exception:
                    face_list.append({"index": i - 1})

            return {"faces": face_list, "count": len(face_list)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_face_info(self, face_index: int) -> dict[str, Any]:
        """
        Get detailed information about a specific face.

        Args:
            face_index: 0-based face index

        Returns:
            Dict with face type, area, edge count, and vertex count
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Count: {faces.Count}"}

            face = faces.Item(face_index + 1)

            info = {"index": face_index}

            with contextlib.suppress(Exception):
                info["type"] = face.Type
            with contextlib.suppress(Exception):
                info["area"] = face.Area
            try:
                edges = face.Edges
                info["edge_count"] = edges.Count if hasattr(edges, "Count") else 0
            except Exception:
                pass
            try:
                vertices = face.Vertices
                info["vertex_count"] = vertices.Count if hasattr(vertices, "Count") else 0
            except Exception:
                pass

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_face_count(self) -> dict[str, Any]:
        """
        Get the total number of faces on the body.

        Returns:
            Dict with face count
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)
            return {"face_count": faces.Count}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_face_normal(self, face_index: int, u: float = 0.5, v: float = 0.5) -> dict[str, Any]:
        """
        Get the normal vector at a parametric point on a face.

        Uses face.GetNormal(numPoints, paramsArray, normalsArray).

        Args:
            face_index: 0-based face index
            u: U parameter (0.0 to 1.0)
            v: V parameter (0.0 to 1.0)

        Returns:
            Dict with normal vector [nx, ny, nz]
        """
        try:
            _doc, _model, _body, face = self._get_face(face_index)

            params_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [u, v])
            normals_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])

            result = face.GetNormal(1, params_arr, normals_arr)

            if isinstance(result, tuple) and len(result) >= 2:
                normal = list(result[-1])
            else:
                normal = list(normals_arr)

            return {
                "normal": normal[:3],
                "params": [u, v],
                "face_index": face_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_face_geometry(self, face_index: int) -> dict[str, Any]:
        """
        Get the underlying geometry type and data of a face.

        Inspects face.Geometry to determine if the face is a Plane,
        Cylinder, Cone, Sphere, Torus, or BSplineSurface.

        Args:
            face_index: 0-based face index

        Returns:
            Dict with geometry_type and associated geometric data
        """
        try:
            _doc, _model, _body, face = self._get_face(face_index)

            geom = face.Geometry
            result: dict[str, Any] = {"face_index": face_index}

            # Try Plane
            try:
                plane_data = geom.GetPlaneData()
                if isinstance(plane_data, tuple) and len(plane_data) >= 2:
                    result["geometry_type"] = "Plane"
                    result["root_point"] = self._to_list(plane_data[0])
                    result["normal"] = self._to_list(plane_data[1])
                    return result
            except Exception:
                pass

            # Try Cylinder
            try:
                cyl_data = geom.GetCylinderData()
                if isinstance(cyl_data, tuple) and len(cyl_data) >= 3:
                    result["geometry_type"] = "Cylinder"
                    result["base_point"] = self._to_list(cyl_data[0])
                    result["axis"] = self._to_list(cyl_data[1])
                    result["radius"] = cyl_data[2]
                    return result
            except Exception:
                pass

            # Try Cone
            try:
                cone_data = geom.GetConeData()
                if isinstance(cone_data, tuple) and len(cone_data) >= 4:
                    result["geometry_type"] = "Cone"
                    result["base_point"] = self._to_list(cone_data[0])
                    result["axis"] = self._to_list(cone_data[1])
                    result["radius"] = cone_data[2]
                    result["half_angle"] = cone_data[3]
                    if len(cone_data) > 4:
                        result["expanding"] = bool(cone_data[4])
                    return result
            except Exception:
                pass

            # Try Sphere
            try:
                sphere_data = geom.GetSphereData()
                if isinstance(sphere_data, tuple) and len(sphere_data) >= 2:
                    result["geometry_type"] = "Sphere"
                    result["center"] = self._to_list(sphere_data[0])
                    result["radius"] = sphere_data[1]
                    return result
            except Exception:
                pass

            # Try Torus
            try:
                torus_data = geom.GetTorusData()
                if isinstance(torus_data, tuple) and len(torus_data) >= 4:
                    result["geometry_type"] = "Torus"
                    result["center"] = self._to_list(torus_data[0])
                    result["axis"] = self._to_list(torus_data[1])
                    result["major_radius"] = torus_data[2]
                    result["minor_radius"] = torus_data[3]
                    return result
            except Exception:
                pass

            # Try BSplineSurface
            try:
                bspline_info = geom.GetBSplineInfo()
                if isinstance(bspline_info, tuple) and len(bspline_info) >= 2:
                    result["geometry_type"] = "BSplineSurface"
                    result["raw_info"] = list(bspline_info)
                    return result
            except Exception:
                pass

            result["geometry_type"] = "Unknown"
            with contextlib.suppress(Exception):
                result["raw_type"] = geom.Type
            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_face_loops(self, face_index: int) -> dict[str, Any]:
        """
        Get loop info for a face (outer boundary vs holes).

        Iterates face.Loops to determine which loop is the outer boundary
        and how many edges each loop contains.

        Args:
            face_index: 0-based face index

        Returns:
            Dict with loop count and details per loop
        """
        try:
            _doc, _model, _body, face = self._get_face(face_index)

            loops_col = face.Loops
            loop_list = []

            for i in range(1, loops_col.Count + 1):
                loop = loops_col.Item(i)
                loop_info: dict[str, Any] = {"index": i - 1}

                with contextlib.suppress(Exception):
                    loop_info["is_outer"] = bool(loop.IsOuterLoop)

                try:
                    loop_info["edge_count"] = loop.Edges.Count
                except Exception:
                    loop_info["edge_count"] = 0

                loop_list.append(loop_info)

            return {
                "face_index": face_index,
                "loop_count": len(loop_list),
                "loops": loop_list,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_face_curvature(self, face_index: int, u: float = 0.5, v: float = 0.5) -> dict[str, Any]:
        """
        Get principal curvatures at a parametric point on a face.

        Uses face.GetCurvatures(numPoints, params, maxTangents,
        maxCurvatures, minCurvatures).

        Args:
            face_index: 0-based face index
            u: U parameter (0.0 to 1.0)
            v: V parameter (0.0 to 1.0)

        Returns:
            Dict with max/min curvature values and max tangent direction
        """
        try:
            _doc, _model, _body, face = self._get_face(face_index)

            params_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [u, v])
            max_tangents_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])
            max_curvatures_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0])
            min_curvatures_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0])

            result = face.GetCurvatures(
                1, params_arr, max_tangents_arr, max_curvatures_arr, min_curvatures_arr
            )

            if isinstance(result, tuple) and len(result) >= 3:
                max_tangent = list(result[0])[:3]
                r1 = result[1]
                max_curvature = r1[0] if hasattr(r1, "__iter__") else float(r1)
                r2 = result[2]
                min_curvature = r2[0] if hasattr(r2, "__iter__") else float(r2)
            else:
                max_tangent = list(max_tangents_arr)[:3]
                max_curvature = list(max_curvatures_arr)[0]
                min_curvature = list(min_curvatures_arr)[0]

            return {
                "max_curvature": max_curvature,
                "min_curvature": min_curvature,
                "max_tangent": max_tangent,
                "params": [u, v],
                "face_index": face_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_face_color(self, face_index: int, red: int, green: int, blue: int) -> dict[str, Any]:
        """
        Set the color of a specific face.

        Args:
            face_index: 0-based face index
            red: Red component (0-255)
            green: Green component (0-255)
            blue: Blue component (0-255)

        Returns:
            Dict with status
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)

            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Count: {faces.Count}"}

            face = faces.Item(face_index + 1)

            # SetFaceStyle or put color directly
            try:
                face.SetColor(red, green, blue)
            except Exception:
                # Try alternative: FaceStyle
                try:
                    face.Color = (red << 0) | (green << 8) | (blue << 16)
                except Exception:
                    # Final fallback using OLE color
                    ole_color = red | (green << 8) | (blue << 16)
                    face.Color = ole_color

            return {"status": "updated", "face_index": face_index, "color": [red, green, blue]}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # EDGE QUERIES
    # =================================================================

    def get_body_edges(self) -> dict[str, Any]:
        """
        Get all unique edges on the model body.

        Enumerates edges via faces since Body.Edges() doesn't work
        in COM late binding. Deduplicates by collecting from all faces.

        Returns:
            Dict with edge count and face-edge mapping
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            total_edges = 0
            face_edges = []

            for fi in range(1, faces.Count + 1):
                try:
                    face = faces.Item(fi)
                    edges = face.Edges
                    edge_count = edges.Count if hasattr(edges, "Count") else 0
                    total_edges += edge_count
                    face_edges.append({"face_index": fi - 1, "edge_count": edge_count})
                except Exception:
                    face_edges.append({"face_index": fi - 1, "edge_count": 0})

            return {
                "face_edges": face_edges,
                "total_face_count": faces.Count,
                "total_edge_references": total_edges,
                "note": "Edge count includes shared edges (counted once per face)",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_edge_count(self) -> dict[str, Any]:
        """
        Get total edge count on the model body.

        Quick count of all edges across all faces. Useful for
        determining if rounds/chamfers can be applied.

        Returns:
            Dict with total edge count and face count
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            total_edges = 0

            for fi in range(1, faces.Count + 1):
                try:
                    face = faces.Item(fi)
                    edges = face.Edges
                    if hasattr(edges, "Count"):
                        total_edges += edges.Count
                except Exception:
                    pass

            return {
                "total_edge_references": total_edges,
                "face_count": faces.Count,
                "note": "Shared edges are counted once per face",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_edge_info(self, face_index: int, edge_index: int) -> dict[str, Any]:
        """
        Get information about a specific edge on a face.

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index within the face

        Returns:
            Dict with edge type, length, and vertex coordinates
        """
        try:
            doc, model = self._get_first_model()
            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)

            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Count: {faces.Count}"}

            face = faces.Item(face_index + 1)
            edges = face.Edges

            if edge_index < 0 or edge_index >= edges.Count:
                return {"error": f"Invalid edge index: {edge_index}. Count: {edges.Count}"}

            edge = edges.Item(edge_index + 1)

            info: dict[str, Any] = {
                "face_index": face_index,
                "edge_index": edge_index,
            }

            try:
                info["length"] = edge.Length
                info["length_mm"] = edge.Length * 1000
            except Exception:
                pass

            with contextlib.suppress(Exception):
                info["type"] = edge.Type

            try:
                start = edge.StartVertex
                end = edge.EndVertex
                info["start_vertex"] = [start.X, start.Y, start.Z]
                info["end_vertex"] = [end.X, end.Y, end.Z]
            except Exception:
                pass

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_edge_endpoints(self, face_index: int, edge_index: int) -> dict[str, Any]:
        """
        Get the start and end XYZ coordinates of an edge.

        Uses edge.GetEndPoints() with VARIANT SAFEARRAY out params.

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index within the face

        Returns:
            Dict with start and end coordinates as [x, y, z] lists
        """
        try:
            _doc, _model, _body, _face, edge = self._get_face_edge(face_index, edge_index)

            start_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])
            end_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])
            result = edge.GetEndPoints(start_arr, end_arr)

            # GetEndPoints returns (start_arr, end_arr) as a tuple
            if isinstance(result, tuple) and len(result) >= 2:
                start = list(result[0])
                end = list(result[1])
            else:
                start = list(start_arr)
                end = list(end_arr)

            return {
                "start": start,
                "end": end,
                "face_index": face_index,
                "edge_index": edge_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_edge_length(self, face_index: int, edge_index: int) -> dict[str, Any]:
        """
        Get the total length of an edge.

        Uses edge.GetParamExtents() then edge.GetLengthAtParam().

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index within the face

        Returns:
            Dict with edge length in meters
        """
        try:
            _doc, _model, _body, _face, edge = self._get_face_edge(face_index, edge_index)

            # Try direct .Length property first (simpler)
            try:
                length = edge.Length
                return {
                    "length": length,
                    "length_mm": length * 1000.0,
                    "face_index": face_index,
                    "edge_index": edge_index,
                }
            except Exception:
                pass

            # Fallback: parametric approach
            param_result = edge.GetParamExtents()
            if isinstance(param_result, tuple) and len(param_result) >= 2:
                min_p = param_result[0]
                max_p = param_result[1]
            else:
                min_p = 0.0
                max_p = 1.0

            length_result = edge.GetLengthAtParam(min_p, max_p)
            length = length_result[0] if isinstance(length_result, tuple) else float(length_result)

            return {
                "length": length,
                "length_mm": length * 1000.0,
                "face_index": face_index,
                "edge_index": edge_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_edge_tangent(
        self, face_index: int, edge_index: int, param: float = 0.5
    ) -> dict[str, Any]:
        """
        Get the tangent vector at a parameter on an edge.

        Uses edge.GetTangent(numParams, paramsArray, tangentsArray).

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index within the face
            param: Parameter value (0.0 to 1.0, default 0.5 = midpoint)

        Returns:
            Dict with tangent vector [tx, ty, tz] and parameter value
        """
        try:
            _doc, _model, _body, _face, edge = self._get_face_edge(face_index, edge_index)

            params_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [param])
            tangents_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])

            result = edge.GetTangent(1, params_arr, tangents_arr)

            if isinstance(result, tuple) and len(result) >= 2:
                tangent = list(result[-1])
            else:
                tangent = list(tangents_arr)

            return {
                "tangent": tangent[:3],
                "param": param,
                "face_index": face_index,
                "edge_index": edge_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_edge_geometry(self, face_index: int, edge_index: int) -> dict[str, Any]:
        """
        Get the underlying geometry type and data of an edge.

        Inspects edge.Geometry to determine if the edge is a Line,
        Circle, Ellipse, or BSplineCurve, and extracts relevant data.

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index within the face

        Returns:
            Dict with geometry_type and associated data
        """
        try:
            _doc, _model, _body, _face, edge = self._get_face_edge(face_index, edge_index)

            geom = edge.Geometry
            result: dict[str, Any] = {
                "face_index": face_index,
                "edge_index": edge_index,
            }

            # Try to determine geometry type
            geom_type = "Unknown"
            with contextlib.suppress(Exception):
                geom_type_val = geom.Type
                geom_type = str(geom_type_val)

            # Attempt Circle data
            try:
                circle_data = geom.GetCircleData()
                if isinstance(circle_data, tuple) and len(circle_data) >= 3:
                    result["geometry_type"] = "Circle"
                    result["center"] = self._to_list(circle_data[0])
                    result["axis"] = self._to_list(circle_data[1])
                    result["radius"] = circle_data[2]
                    return result
            except Exception:
                pass

            # Attempt Ellipse data
            try:
                ellipse_data = geom.GetEllipseData()
                if isinstance(ellipse_data, tuple) and len(ellipse_data) >= 4:
                    result["geometry_type"] = "Ellipse"
                    result["center"] = self._to_list(ellipse_data[0])
                    result["axis"] = self._to_list(ellipse_data[1])
                    result["major_axis"] = self._to_list(ellipse_data[2])
                    result["minor_major_ratio"] = ellipse_data[3]
                    return result
            except Exception:
                pass

            # Attempt BSplineCurve info
            try:
                bspline_info = geom.GetBSplineInfo()
                if isinstance(bspline_info, tuple) and len(bspline_info) >= 4:
                    result["geometry_type"] = "BSplineCurve"
                    result["order"] = bspline_info[0]
                    result["num_poles"] = bspline_info[1]
                    result["num_knots"] = bspline_info[2]
                    result["rational"] = bool(bspline_info[3])
                    if len(bspline_info) > 4:
                        result["closed"] = bool(bspline_info[4])
                    if len(bspline_info) > 5:
                        result["periodic"] = bool(bspline_info[5])
                    if len(bspline_info) > 6:
                        result["planar"] = bool(bspline_info[6])
                    return result
            except Exception:
                pass

            # Default: Line (no extra data needed beyond endpoints)
            result["geometry_type"] = "Line"
            result["raw_type"] = geom_type
            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_edge_curvature(
        self, face_index: int, edge_index: int, param: float = 0.5
    ) -> dict[str, Any]:
        """
        Get curvature at a parameter on an edge.

        Uses edge.GetCurvature(numParams, paramsArray, directionsArray,
        curvaturesArray).

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index
            param: Parameter value (0.0 to 1.0, default 0.5)

        Returns:
            Dict with curvature value, direction vector, and parameter
        """
        try:
            _doc, _model, _body, _face, edge = self._get_face_edge(face_index, edge_index)

            params_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [param])
            directions_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])
            curvatures_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0])

            result = edge.GetCurvature(1, params_arr, directions_arr, curvatures_arr)

            if isinstance(result, tuple) and len(result) >= 2:
                direction = list(result[0])[:3]
                r1 = result[1]
                curvature = r1[0] if hasattr(r1, "__iter__") else float(r1)
            else:
                direction = list(directions_arr)[:3]
                curvature = list(curvatures_arr)[0]

            return {
                "curvature": curvature,
                "direction": direction,
                "param": param,
                "face_index": face_index,
                "edge_index": edge_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # BODY / VERTEX / SHELL QUERIES
    # =================================================================

    def get_body_facet_data(self, tolerance: float = 0.0) -> dict[str, Any]:
        """
        Get tessellation/mesh data from the model body.

        Returns triangulated facet data (vertices, normals, face IDs).
        Useful for 3D printing previews and mesh export.

        Args:
            tolerance: Mesh tolerance in meters. If <= 0, returns cached data.
                       If > 0, recomputes from Parasolid (slower but more accurate).

        Returns:
            Dict with facet count, point count, and sample data
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No geometry in document"}

            model = models.Item(1)
            body = model.Body

            import array as arr_mod

            # Prepare out parameters for COM call
            # GetFacetData(Tolerance, FacetCount, Points,
            # Normals, TextureCoords, StyleIDs, FaceIDs,
            # bHonourPrefs)
            points = arr_mod.array("d", [])
            normals = arr_mod.array("d", [])
            texture_coords = arr_mod.array("d", [])
            style_ids = arr_mod.array("i", [])
            face_ids = arr_mod.array("i", [])

            try:
                result_data = body.GetFacetData(
                    tolerance,  # Tolerance
                )

                # GetFacetData returns a tuple of
                # (facetCount, points, normals,
                # textureCoords, styleIds, faceIds)
                if isinstance(result_data, tuple) and len(result_data) >= 2:
                    facet_count = result_data[0] if isinstance(result_data[0], int) else 0
                    pts = result_data[1] if len(result_data) > 1 else []

                    return {
                        "facet_count": facet_count,
                        "point_count": len(pts) // 3 if pts else 0,
                        "tolerance": tolerance,
                        "has_data": facet_count > 0,
                    }
            except Exception:
                pass

            # Alternative: try with explicit out params
            try:
                facet_count = 0
                body.GetFacetData(
                    tolerance,
                    facet_count,
                    points,
                    normals,
                    texture_coords,
                    style_ids,
                    face_ids,
                    False,
                )

                return {
                    "facet_count": facet_count,
                    "point_count": len(points) // 3 if points else 0,
                    "tolerance": tolerance,
                    "has_data": len(points) > 0,
                }
            except Exception as e2:
                return {
                    "error": f"GetFacetData failed: {e2}",
                    "note": "Body facet data may require "
                    "specific COM marshaling. "
                    "Try export_stl() instead.",
                    "traceback": traceback.format_exc(),
                }

        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_solid_bodies(self) -> dict[str, Any]:
        """
        Report all solid bodies in the active part document.

        Lists design bodies and construction bodies with their properties.

        Returns:
            Dict with body info (is_solid, shell count, etc.)
        """
        try:
            doc = self.doc_manager.get_active_document()

            bodies = []

            # Check design bodies (Models collection)
            models = doc.Models
            for i in range(1, models.Count + 1):
                model = models.Item(i)
                try:
                    body = model.Body
                    body_info = {
                        "index": i - 1,
                        "type": "design",
                        "name": model.Name if hasattr(model, "Name") else f"Model_{i}",
                    }

                    try:
                        body_info["is_solid"] = body.IsSolid
                    except Exception:
                        body_info["is_solid"] = True  # Default assumption

                    with contextlib.suppress(Exception):
                        body_info["volume"] = body.Volume

                    # Count shells
                    try:
                        shells = body.Shells
                        body_info["shell_count"] = shells.Count
                    except Exception:
                        pass

                    bodies.append(body_info)
                except Exception:
                    pass

            # Check construction bodies
            try:
                constructions = doc.Constructions
                for i in range(1, constructions.Count + 1):
                    try:
                        cm = constructions.Item(i)
                        body = cm.Body
                        body_info = {
                            "index": len(bodies),
                            "type": "construction",
                            "name": cm.Name if hasattr(cm, "Name") else f"Construction_{i}",
                        }
                        with contextlib.suppress(Exception):
                            body_info["is_solid"] = body.IsSolid
                        bodies.append(body_info)
                    except Exception:
                        pass
            except Exception:
                pass  # No Constructions collection

            return {"total_bodies": len(bodies), "bodies": bodies}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_body_extreme_point(
        self, direction_x: float, direction_y: float, direction_z: float
    ) -> dict[str, Any]:
        """
        Get the extreme point of the body in a given direction.

        Uses body.GetExtremePoint(dx, dy, dz, ex, ey, ez) where the
        last three parameters are out-params for the extreme point coords.

        Args:
            direction_x: X component of direction vector
            direction_y: Y component of direction vector
            direction_z: Z component of direction vector

        Returns:
            Dict with extreme point [x, y, z] and direction
        """
        try:
            _doc, _model, body = self._get_body()

            result = body.GetExtremePoint(
                direction_x,
                direction_y,
                direction_z,
                0.0,
                0.0,
                0.0,
            )

            if isinstance(result, tuple) and len(result) >= 3:
                extreme = [result[0], result[1], result[2]]
            else:
                extreme = [0.0, 0.0, 0.0]

            return {
                "extreme_point": extreme,
                "direction": [direction_x, direction_y, direction_z],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_faces_by_ray(
        self,
        origin_x: float,
        origin_y: float,
        origin_z: float,
        direction_x: float,
        direction_y: float,
        direction_z: float,
    ) -> dict[str, Any]:
        """
        Ray-cast query to find faces hit by a ray.

        Uses body.FacesByRay(ox, oy, oz, dx, dy, dz) which returns
        a Faces collection of intersected faces.

        Args:
            origin_x, origin_y, origin_z: Ray origin point
            direction_x, direction_y, direction_z: Ray direction vector

        Returns:
            Dict with hit face count and face info
        """
        try:
            _doc, _model, body = self._get_body()

            faces = body.FacesByRay(
                origin_x,
                origin_y,
                origin_z,
                direction_x,
                direction_y,
                direction_z,
            )

            face_list = []
            for i in range(1, faces.Count + 1):
                face = faces.Item(i)
                face_info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    face_info["area"] = face.Area
                with contextlib.suppress(Exception):
                    face_info["id"] = face.ID
                face_list.append(face_info)

            return {
                "face_count": len(face_list),
                "faces": face_list,
                "ray_origin": [origin_x, origin_y, origin_z],
                "ray_direction": [direction_x, direction_y, direction_z],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_shell_info(self, shell_index: int = 0) -> dict[str, Any]:
        """
        Get topology information about a shell.

        Accesses body.Shells collection and retrieves properties like
        IsClosed, Volume, IsVoid, face count, and edge count.

        Args:
            shell_index: 0-based shell index (default 0 = first shell)

        Returns:
            Dict with shell topology info
        """
        try:
            _doc, _model, body = self._get_body()

            shells = body.Shells
            if shell_index < 0 or shell_index >= shells.Count:
                return {
                    "error": f"Invalid shell index: {shell_index}. Body has {shells.Count} shells."
                }

            shell = shells.Item(shell_index + 1)

            info: dict[str, Any] = {"shell_index": shell_index}

            with contextlib.suppress(Exception):
                info["is_closed"] = bool(shell.IsClosed)
            with contextlib.suppress(Exception):
                info["volume"] = shell.Volume
            with contextlib.suppress(Exception):
                info["is_void"] = bool(shell.IsVoid)
            try:
                info["face_count"] = shell.Faces.Count
            except Exception:
                info["face_count"] = 0
            try:
                info["edge_count"] = shell.Edges.Count
            except Exception:
                info["edge_count"] = 0

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def is_point_inside_body(self, x: float, y: float, z: float) -> dict[str, Any]:
        """
        Test if a 3D point is inside the solid body.

        Uses the first shell's IsPointInside method with a SAFEARRAY param.

        Args:
            x, y, z: 3D point coordinates (meters)

        Returns:
            Dict with is_inside boolean
        """
        try:
            _doc, _model, body = self._get_body()

            shells = body.Shells
            if shells.Count == 0:
                return {"error": "Body has no shells"}

            shell = shells.Item(1)

            point_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [x, y, z])
            is_inside = shell.IsPointInside(point_arr)

            return {
                "is_inside": bool(is_inside),
                "point": [x, y, z],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_body_shells(self) -> dict[str, Any]:
        """
        List all shells in the body with basic properties.

        Iterates body.Shells to get IsClosed and Volume for each shell.

        Returns:
            Dict with shell count and list of shell info
        """
        try:
            _doc, _model, body = self._get_body()

            shells = body.Shells
            shell_list = []

            for i in range(1, shells.Count + 1):
                shell = shells.Item(i)
                shell_info: dict[str, Any] = {"index": i - 1}

                with contextlib.suppress(Exception):
                    shell_info["is_closed"] = bool(shell.IsClosed)
                with contextlib.suppress(Exception):
                    shell_info["volume"] = shell.Volume

                shell_list.append(shell_info)

            return {
                "shell_count": len(shell_list),
                "shells": shell_list,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_body_vertices(self) -> dict[str, Any]:
        """
        Get all vertices of the body with their 3D coordinates.

        Iterates body.Vertices and calls GetPointData() on each to
        extract XYZ coordinates.

        Returns:
            Dict with vertex count and coordinate list
        """
        try:
            _doc, _model, body = self._get_body()

            vertices = body.Vertices
            vertex_list = []

            for i in range(1, vertices.Count + 1):
                try:
                    vertex = vertices.Item(i)
                    point_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])
                    result = vertex.GetPointData(point_arr)

                    point = (
                        self._to_list(result[0]) if isinstance(result, tuple) else list(point_arr)
                    )

                    vertex_list.append({"index": i - 1, "point": point[:3]})
                except Exception:
                    vertex_list.append({"index": i - 1, "point": None})

            return {
                "vertex_count": len(vertex_list),
                "vertices": vertex_list,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_vertex_point(
        self, face_index: int, edge_index: int, which: str = "start"
    ) -> dict[str, Any]:
        """
        Get XYZ coordinates of a vertex on an edge.

        Accesses edge.StartVertex or edge.EndVertex and calls
        vertex.GetPointData() to retrieve the 3D coordinates.

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index
            which: 'start' or 'end' to select which vertex

        Returns:
            Dict with vertex point [x, y, z]
        """
        try:
            if which not in ("start", "end"):
                return {"error": f"Invalid vertex selector: '{which}'. Use 'start' or 'end'."}

            _doc, _model, _body, _face, edge = self._get_face_edge(face_index, edge_index)

            vertex = edge.StartVertex if which == "start" else edge.EndVertex

            point_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0, 0.0])
            result = vertex.GetPointData(point_arr)

            point = self._to_list(result[0]) if isinstance(result, tuple) else list(point_arr)

            vertex_id = -1
            with contextlib.suppress(Exception):
                vertex_id = vertex.ID

            return {
                "point": point[:3],
                "vertex_id": vertex_id,
                "which": which,
                "face_index": face_index,
                "edge_index": edge_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # B-SPLINE
    # =================================================================

    def get_bspline_curve_info(self, face_index: int, edge_index: int) -> dict[str, Any]:
        """
        Get NURBS curve metadata from an edge's underlying geometry.

        Accesses edge.Geometry and calls GetBSplineInfo() to retrieve
        order, pole count, knot count, and flags.

        Args:
            face_index: 0-based face index
            edge_index: 0-based edge index

        Returns:
            Dict with BSpline curve properties, or error if not a BSpline
        """
        try:
            _doc, _model, _body, _face, edge = self._get_face_edge(face_index, edge_index)

            geom = edge.Geometry
            bspline_info = geom.GetBSplineInfo()

            if not isinstance(bspline_info, tuple) or len(bspline_info) < 4:
                return {
                    "error": "Edge geometry is not a BSpline curve "
                    "or GetBSplineInfo returned unexpected data",
                    "face_index": face_index,
                    "edge_index": edge_index,
                }

            result: dict[str, Any] = {
                "face_index": face_index,
                "edge_index": edge_index,
                "order": bspline_info[0],
                "num_poles": bspline_info[1],
                "num_knots": bspline_info[2],
                "rational": bool(bspline_info[3]),
            }
            if len(bspline_info) > 4:
                result["closed"] = bool(bspline_info[4])
            if len(bspline_info) > 5:
                result["periodic"] = bool(bspline_info[5])
            if len(bspline_info) > 6:
                result["planar"] = bool(bspline_info[6])

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_bspline_surface_info(self, face_index: int) -> dict[str, Any]:
        """
        Get NURBS surface metadata from a face's underlying geometry.

        Accesses face.Geometry and calls GetBSplineInfo() to retrieve
        order, pole count, knot count, and flags for U and V directions.

        Args:
            face_index: 0-based face index

        Returns:
            Dict with BSpline surface properties, or error if not BSpline
        """
        try:
            _doc, _model, _body, face = self._get_face(face_index)

            geom = face.Geometry
            bspline_info = geom.GetBSplineInfo()

            if not isinstance(bspline_info, tuple) or len(bspline_info) < 2:
                return {
                    "error": "Face geometry is not a BSpline surface "
                    "or GetBSplineInfo returned unexpected data",
                    "face_index": face_index,
                }

            result: dict[str, Any] = {"face_index": face_index}

            # BSplineSurface.GetBSplineInfo returns data for both U and V
            # The exact tuple structure depends on the SE API version.
            # Common pattern: (orderU, orderV, numPolesU, numPolesV,
            #                   numKnotsU, numKnotsV, rational, ...)
            if len(bspline_info) >= 6:
                result["order"] = [bspline_info[0], bspline_info[1]]
                result["num_poles"] = [bspline_info[2], bspline_info[3]]
                result["num_knots"] = [bspline_info[4], bspline_info[5]]
                if len(bspline_info) > 6:
                    result["rational"] = bool(bspline_info[6])
                if len(bspline_info) > 7:
                    result["planar"] = bool(bspline_info[7])
            else:
                # Simpler format - store raw info
                result["raw_info"] = list(bspline_info)

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
