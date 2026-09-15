"""B-Rep topology queries: faces, edges, vertices, shells, and geometry inspection."""

import contextlib
from typing import Any

from solidedge_mcp.backends.errors import error_result

from ..comutil import OWNED_STYLE_PREFIX, com_get, face_style_named
from ..logging import get_logger
from ._base import (
    DEFAULT_PAGE_LIMIT,
    all_faces,
    body_of,
    bool_array,
    i4_array,
    page_bounds,
    page_result,
    r8_array,
)

_logger = get_logger(__name__)

# geometry.tlb > GNTTypePropertyConstants. Face.GeometryForm / Edge.GeometryForm
# return one of these, which is a single COM property read per entity — far
# cheaper than re-querying Body.Faces() once per geometry type.
_GEOMETRY_FORM_NAMES: dict[int, str] = {
    -1909484335: "plane",  # igPlane
    -114972029: "cylinder",  # igCylinder
    -114972031: "cone",  # igCone
    -114972027: "sphere",  # igSphere
    -114972025: "torus",  # igTorus
    1465959633: "bspline_surface",  # igBSplineSurface
    -2071771273: "mesh",  # igMesh
    167551103: "bspline_curve",  # igBSplineCurve
    167551105: "circle",  # igCircle
    167551107: "ellipse",  # igEllipse
    167551109: "line",  # igLine
    -1811952078: "param_bspline_curve",  # igParamBSplineCurve
}


class BRepMixin:
    """Mixin providing B-Rep topology query methods."""

    doc_manager: Any

    def get_body_faces(self, offset: int = 0, limit: int = DEFAULT_PAGE_LIMIT) -> dict[str, Any]:
        """
        Get a page of faces on the model body.

        Uses Body.Faces(igQueryAll=1) for the total and reads area, edge count
        and geometry form for the requested window only, so the cost is bounded
        by ``limit`` rather than by the size of the model.

        Args:
            offset: 0-based index of the first face to return
            limit: maximum number of faces to return (clamped to MAX_PAGE_LIMIT)

        Returns:
            Paging envelope: total, offset, limit, items, truncated
        """
        try:
            doc, model = self._get_first_model()
            body = body_of(model)

            faces = all_faces(body, model)
            total = faces.Count
            start, stop, limit = page_bounds(total, offset, limit)

            face_list = []
            for i in range(start + 1, stop + 1):
                try:
                    face = faces.Item(i)
                    face_info: dict[str, Any] = {"index": i - 1}
                    with contextlib.suppress(Exception):
                        face_info["area"] = face.Area
                    with contextlib.suppress(Exception):
                        face_info["edge_count"] = face.Edges.Count
                    face_info["geometry"] = self._geometry_form_name(face)
                    face_list.append(face_info)
                except Exception:
                    face_list.append({"index": i - 1})

            return page_result(face_list, total, start, limit)
        except Exception as e:
            return error_result(e)

    @staticmethod
    def _geometry_form_name(entity: Any) -> str:
        """Map a Face/Edge onto a readable geometry name.

        ``Geometry.Type`` returns a GNTTypePropertyConstants value and is the
        documented route. ``GeometryForm`` is declared as a bare VT_I4 with no
        enum behind it and returns unrelated small integers on Solid Edge 2026
        (9 for a plane), so it is only a fallback.
        """
        try:
            return _GEOMETRY_FORM_NAMES.get(int(entity.Geometry.Type), "unknown")
        except Exception:
            pass
        try:
            return _GEOMETRY_FORM_NAMES.get(int(entity.GeometryForm), "unknown")
        except Exception:
            return "unknown"

    @staticmethod
    def _bspline_surface_info(geom: Any) -> Any:
        """Call BSplineSurface.GetBSplineInfo with the buffers it requires.

        geometry.tlb BSplineSurface.GetBSplineInfo(
            Order SAFEARRAY(VT_I4)* [in,out],
            NumPoles SAFEARRAY(VT_I4)* [in,out],
            NumKnots SAFEARRAY(VT_I4)* [in,out],
            Rational VT_BOOL* [out],
            Closed SAFEARRAY(VT_BOOL)* [in,out],
            Periodic SAFEARRAY(VT_BOOL)* [in,out],
            Planar VT_BOOL* [out])

        ``Rational`` sits between the [in,out] arrays, so the buffers go in by
        keyword and pywin32 fills the [out] slots itself. Every array holds two
        entries, one for U and one for V.
        """
        return geom.GetBSplineInfo(
            Order=i4_array(2),
            NumPoles=i4_array(2),
            NumKnots=i4_array(2),
            Closed=bool_array(2),
            Periodic=bool_array(2),
        )

    @staticmethod
    def _bspline_curve_info(geom: Any) -> Any:
        """Call BSplineCurve.GetBSplineInfo with the buffer it requires.

        geometry.tlb BSplineCurve.GetBSplineInfo(
            Order VT_I4* [out], NumPoles VT_I4* [out], NumKnots VT_I4* [out],
            Rational VT_BOOL* [out], Closed VT_BOOL* [out],
            Periodic VT_BOOL* [out], Planar VT_BOOL* [out],
            PlaneVector SAFEARRAY(VT_R8)* [in,out])

        Only the trailing plane vector is an [in,out] buffer; it is passed by
        keyword so the seven [out] slots ahead of it stay untouched.
        """
        return geom.GetBSplineInfo(PlaneVector=r8_array(3))

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
            body = body_of(model)

            faces = all_faces(body, model)
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
                info["edge_count"] = com_get(edges, "Count", 0)
            except Exception:
                pass
            try:
                vertices = face.Vertices
                info["vertex_count"] = com_get(vertices, "Count", 0)
            except Exception:
                pass

            return info
        except Exception as e:
            return error_result(e)

    def get_face_count(self) -> dict[str, Any]:
        """
        Get the total number of faces on the body.

        Returns:
            Dict with face count
        """
        try:
            doc, model = self._get_first_model()
            body = body_of(model)
            faces = all_faces(body, model)
            return {"face_count": faces.Count}
        except Exception as e:
            return error_result(e)

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

            params_arr = [u, v]
            normals_arr = [0.0, 0.0, 0.0]

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
            return error_result(e)

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
            # geometry.tlb Plane.GetPlaneData(
            #   RootPoint SAFEARRAY(VT_R8)* [in,out],
            #   NormalVector SAFEARRAY(VT_R8)* [in,out])
            try:
                plane_data = geom.GetPlaneData(r8_array(3), r8_array(3))
                if isinstance(plane_data, tuple) and len(plane_data) >= 2:
                    result["geometry_type"] = "Plane"
                    result["root_point"] = self._to_list(plane_data[0])
                    result["normal"] = self._to_list(plane_data[1])
                    return result
            except Exception:
                pass

            # Try Cylinder
            # geometry.tlb Cylinder.GetCylinderData(
            #   BasePoint SAFEARRAY(VT_R8)* [in,out],
            #   AxisVector SAFEARRAY(VT_R8)* [in,out], Radius VT_R8* [out])
            try:
                cyl_data = geom.GetCylinderData(r8_array(3), r8_array(3))
                if isinstance(cyl_data, tuple) and len(cyl_data) >= 3:
                    result["geometry_type"] = "Cylinder"
                    result["base_point"] = self._to_list(cyl_data[0])
                    result["axis"] = self._to_list(cyl_data[1])
                    result["radius"] = cyl_data[2]
                    return result
            except Exception:
                pass

            # Try Cone
            # geometry.tlb Cone.GetConeData(
            #   BasePoint SAFEARRAY(VT_R8)* [in,out],
            #   AxisVector SAFEARRAY(VT_R8)* [in,out], Radius VT_R8* [out],
            #   HalfAngle VT_R8* [out], Expanding VT_BOOL* [out])
            try:
                cone_data = geom.GetConeData(r8_array(3), r8_array(3))
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
            # geometry.tlb Sphere.GetSphereData(
            #   CenterPoint SAFEARRAY(VT_R8)* [in,out], Radius VT_R8* [out])
            try:
                sphere_data = geom.GetSphereData(r8_array(3))
                if isinstance(sphere_data, tuple) and len(sphere_data) >= 2:
                    result["geometry_type"] = "Sphere"
                    result["center"] = self._to_list(sphere_data[0])
                    result["radius"] = sphere_data[1]
                    return result
            except Exception:
                pass

            # Try Torus
            # geometry.tlb Torus.GetTorusData(
            #   CenterPoint SAFEARRAY(VT_R8)* [in,out],
            #   AxisVector SAFEARRAY(VT_R8)* [in,out],
            #   MajorRadius VT_R8* [out], MinorRadius VT_R8* [out])
            try:
                torus_data = geom.GetTorusData(r8_array(3), r8_array(3))
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
                bspline_info = self._bspline_surface_info(geom)
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
            return error_result(e)

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
            return error_result(e)

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

            params_arr = [u, v]
            max_tangents_arr = [0.0, 0.0, 0.0]
            max_curvatures_arr = [0.0]
            min_curvatures_arr = [0.0]

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
            return error_result(e)

    def set_face_color(self, face_index: int, red: int, green: int, blue: int) -> dict[str, Any]:
        """Set the colour of one face.

        ``Face.SetColor`` and ``Face.Color`` are on no Solid Edge interface.
        This tried both, then a third spelling of the same missing property,
        and Solid Edge 2026 answered "Property 'Item.Color' can not be set."
        every time -- the three nested try/excepts hid which of them failed.

        ``Face.Style`` is a get/put ``FaceStyle``, exactly like ``Body.Style``,
        so a face is coloured the same way a body is: by assigning it a style
        whose diffuse colour is the one you want. Faces sharing a colour share
        one style, which is why it is named after the colour.

        Args:
            face_index: 0-based face index.
            red: Red channel, 0-255.
            green: Green channel, 0-255.
            blue: Blue channel, 0-255.

        Returns:
            Dict with status, the colour, and the style carrying it.
        """
        try:
            doc, model = self._get_first_model()
            body = body_of(model)
            faces = all_faces(body, model)

            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Count: {faces.Count}"}

            red = max(0, min(255, red))
            green = max(0, min(255, green))
            blue = max(0, min(255, blue))

            name = f"{OWNED_STYLE_PREFIX}Face {red:02X}{green:02X}{blue:02X}"
            style, err = face_style_named(doc, name)
            if err:
                return err

            # SetDiffuse takes 0.0-1.0 per channel, not 0-255.
            style.SetDiffuse(red / 255.0, green / 255.0, blue / 255.0)
            faces.Item(face_index + 1).Style = style

            return {
                "status": "updated",
                "face_index": face_index,
                "color": [red, green, blue],
                "hex": f"#{red:02x}{green:02x}{blue:02x}",
                "style": name,
            }
        except Exception as e:
            return error_result(e)

    # =================================================================
    # EDGE QUERIES
    # =================================================================

    def get_body_edges(self, offset: int = 0, limit: int = DEFAULT_PAGE_LIMIT) -> dict[str, Any]:
        """
        Get a page of the body's face-to-edge mapping.

        Enumerates edges via faces since Body.Edges() doesn't work in COM late
        binding, so one item is emitted per face. ``total`` is the face count;
        page with ``offset``/``limit`` until ``truncated`` is False.

        Args:
            offset: 0-based index of the first face to report
            limit: maximum number of faces to report (clamped to MAX_PAGE_LIMIT)

        Returns:
            Paging envelope: total, offset, limit, items, truncated, plus
            page_edge_references (edges counted across this page only)
        """
        try:
            doc, model = self._get_first_model()
            body = body_of(model)

            faces = all_faces(body, model)
            total = faces.Count
            start, stop, limit = page_bounds(total, offset, limit)

            page_edges = 0
            face_edges = []
            for fi in range(start + 1, stop + 1):
                try:
                    face = faces.Item(fi)
                    edge_count = face.Edges.Count
                    page_edges += edge_count
                    face_edges.append({"face_index": fi - 1, "edge_count": edge_count})
                except Exception:
                    face_edges.append({"face_index": fi - 1, "edge_count": 0})

            return page_result(
                face_edges,
                total,
                start,
                limit,
                page_edge_references=page_edges,
                note="Edge counts include shared edges (counted once per face)",
            )
        except Exception as e:
            return error_result(e)

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
            body = body_of(model)

            faces = all_faces(body, model)
            total_edges = 0

            for fi in range(1, faces.Count + 1):
                try:
                    face = faces.Item(fi)
                    edges = face.Edges
                    total_edges += com_get(edges, "Count", 0)
                except Exception:
                    pass

            return {
                "total_edge_references": total_edges,
                "face_count": faces.Count,
                "note": "Shared edges are counted once per face",
            }
        except Exception as e:
            return error_result(e)

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
            body = body_of(model)
            faces = all_faces(body, model)

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
            return error_result(e)

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

            start_arr = [0.0, 0.0, 0.0]
            end_arr = [0.0, 0.0, 0.0]
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
            return error_result(e)

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
            return error_result(e)

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

            params_arr = [param]
            tangents_arr = [0.0, 0.0, 0.0]

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
            return error_result(e)

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
            # geometry.tlb Circle.GetCircleData(
            #   CenterPoint SAFEARRAY(VT_R8)* [in,out],
            #   AxisVector SAFEARRAY(VT_R8)* [in,out], Radius VT_R8* [out])
            try:
                circle_data = geom.GetCircleData(r8_array(3), r8_array(3))
                if isinstance(circle_data, tuple) and len(circle_data) >= 3:
                    result["geometry_type"] = "Circle"
                    result["center"] = self._to_list(circle_data[0])
                    result["axis"] = self._to_list(circle_data[1])
                    result["radius"] = circle_data[2]
                    return result
            except Exception:
                pass

            # Attempt Ellipse data
            # geometry.tlb Ellipse.GetEllipseData(
            #   CenterPoint SAFEARRAY(VT_R8)* [in,out],
            #   AxisVector SAFEARRAY(VT_R8)* [in,out],
            #   MajorAxis SAFEARRAY(VT_R8)* [in,out],
            #   MinorMajorRatio VT_R8* [out])
            try:
                ellipse_data = geom.GetEllipseData(r8_array(3), r8_array(3), r8_array(3))
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
                bspline_info = self._bspline_curve_info(geom)
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
            return error_result(e)

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

            params_arr = [param]
            directions_arr = [0.0, 0.0, 0.0]
            curvatures_arr = [0.0]

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
            return error_result(e)

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
            body = body_of(model)

            # geometry.tlb Body.GetFacetData(
            #   Tolerance VT_R8 [in], FacetCount VT_I4* [out],
            #   Points SAFEARRAY(VT_R8)* [in,out],
            #   Normals/TextureCoords/StyleIDs/FaceIDs VT_VARIANT* [out,optional],
            #   bHonourPrefs VT_VARIANT [in,optional])
            # FacetCount sits between the two arguments we supply, so Points is
            # passed by keyword. Solid Edge resizes the buffer it is handed.
            try:
                result_data = body.GetFacetData(Tolerance=tolerance, Points=r8_array(1))
            except Exception as e2:
                return error_result(
                    e2,
                    note="Body facet data may require specific COM marshaling. "
                    "Try export_stl() instead.",
                )

            # Returns (FacetCount, Points, ...optional out params).
            facet_count = 0
            points: Any = []
            if isinstance(result_data, tuple) and len(result_data) >= 2:
                if isinstance(result_data[0], int):
                    facet_count = result_data[0]
                points = result_data[1] or []

            return {
                "facet_count": facet_count,
                "point_count": len(points) // 3 if points else 0,
                "tolerance": tolerance,
                "has_data": facet_count > 0 or bool(points),
            }
        except Exception as e:
            return error_result(e)

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
                    body = body_of(model)
                    body_info = {
                        "index": i - 1,
                        "type": "design",
                        "name": com_get(model, "Name", f"Model_{i}"),
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
                            "name": com_get(cm, "Name", f"Construction_{i}"),
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
            return error_result(e)

    def get_body_extreme_point(
        self, direction_x: float, direction_y: float, direction_z: float
    ) -> dict[str, Any]:
        """
        Get the extreme point of the body in a given direction.

        geometry.tlb Body.GetExtremePoint(
            DirectionX VT_R8 [in], DirectionY VT_R8 [in], DirectionZ VT_R8 [in],
            ExtremeX VT_R8* [out], ExtremeY VT_R8* [out], ExtremeZ VT_R8* [out])

        Only the three direction components are passed; pywin32 returns the
        [out] coordinates as the result tuple.

        Args:
            direction_x: X component of direction vector
            direction_y: Y component of direction vector
            direction_z: Z component of direction vector

        Returns:
            Dict with extreme point [x, y, z] and direction
        """
        try:
            _doc, _model, body = self._get_body()

            result = body.GetExtremePoint(direction_x, direction_y, direction_z)

            if isinstance(result, tuple) and len(result) >= 3:
                extreme = [result[0], result[1], result[2]]
            else:
                extreme = [0.0, 0.0, 0.0]

            return {
                "extreme_point": extreme,
                "direction": [direction_x, direction_y, direction_z],
            }
        except Exception as e:
            return error_result(e)

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
            return error_result(e)

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
            return error_result(e)

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

            point_arr = [x, y, z]
            is_inside = shell.IsPointInside(point_arr)

            return {
                "is_inside": bool(is_inside),
                "point": [x, y, z],
            }
        except Exception as e:
            return error_result(e)

    def get_body_shells(self, offset: int = 0, limit: int = DEFAULT_PAGE_LIMIT) -> dict[str, Any]:
        """
        List a page of shells in the body with basic properties.

        Iterates body.Shells to get IsClosed and Volume for each shell in the
        requested window; page with ``offset``/``limit`` until ``truncated``
        is False.

        Args:
            offset: 0-based index of the first shell to return
            limit: maximum number of shells to return (clamped to MAX_PAGE_LIMIT)

        Returns:
            Paging envelope: total, offset, limit, items, truncated
        """
        try:
            _doc, _model, body = self._get_body()

            shells = body.Shells
            total = shells.Count
            start, stop, limit = page_bounds(total, offset, limit)

            shell_list = []
            for i in range(start + 1, stop + 1):
                shell = shells.Item(i)
                shell_info: dict[str, Any] = {"index": i - 1}

                with contextlib.suppress(Exception):
                    shell_info["is_closed"] = bool(shell.IsClosed)
                with contextlib.suppress(Exception):
                    shell_info["volume"] = shell.Volume

                shell_list.append(shell_info)

            return page_result(shell_list, total, start, limit)
        except Exception as e:
            return error_result(e)

    def get_body_vertices(self, offset: int = 0, limit: int = DEFAULT_PAGE_LIMIT) -> dict[str, Any]:
        """
        Get a page of body vertices with their 3D coordinates.

        Iterates body.Vertices and calls GetPointData() on each vertex in the
        requested window; page with ``offset``/``limit`` until ``truncated``
        is False.

        Args:
            offset: 0-based index of the first vertex to return
            limit: maximum number of vertices to return (clamped to MAX_PAGE_LIMIT)

        Returns:
            Paging envelope: total, offset, limit, items, truncated
        """
        try:
            _doc, _model, body = self._get_body()

            vertices = body.Vertices
            total = vertices.Count
            start, stop, limit = page_bounds(total, offset, limit)

            vertex_list = []
            for i in range(start + 1, stop + 1):
                try:
                    vertex = vertices.Item(i)
                    point_arr = [0.0, 0.0, 0.0]
                    result = vertex.GetPointData(point_arr)

                    point = (
                        self._to_list(result[0]) if isinstance(result, tuple) else list(point_arr)
                    )

                    vertex_list.append({"index": i - 1, "point": point[:3]})
                except Exception:
                    vertex_list.append({"index": i - 1, "point": None})

            return page_result(vertex_list, total, start, limit)
        except Exception as e:
            return error_result(e)

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

            point_arr = [0.0, 0.0, 0.0]
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
            return error_result(e)

    # =================================================================
    # B-SPLINE
    # =================================================================

    def _not_a_bspline(self, entity: Any, kind: str, **where: Any) -> dict[str, Any] | None:
        """Refuse a GetBSplineInfo call on geometry that is not a spline.

        Only BSplineCurve and BSplineSurface carry GetBSplineInfo, so a planar
        face or a straight edge answers with a bare attribute error naming a
        member the reader then cannot find. Say what the geometry is instead.
        """
        form = self._geometry_form_name(entity)
        if form in ("bspline_curve", "bspline_surface", "unknown"):
            return None
        return {
            "error": (
                f"This {kind} is a {form.replace('_', ' ')}, not a B-spline, so it "
                f"has no NURBS data. Read solidedge://geometry/face/N to see what "
                f"each face is."
            ),
            "geometry": form,
            **where,
        }

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

            err = self._not_a_bspline(edge, "edge", face_index=face_index, edge_index=edge_index)
            if err:
                return err

            geom = edge.Geometry
            bspline_info = self._bspline_curve_info(geom)

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
            return error_result(e)

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

            err = self._not_a_bspline(face, "face", face_index=face_index)
            if err:
                return err

            geom = face.Geometry
            bspline_info = self._bspline_surface_info(geom)

            if not isinstance(bspline_info, tuple) or len(bspline_info) < 2:
                return {
                    "error": "Face geometry is not a BSpline surface "
                    "or GetBSplineInfo returned unexpected data",
                    "face_index": face_index,
                }

            result: dict[str, Any] = {"face_index": face_index}

            # Returns (Order, NumPoles, NumKnots, Rational, Closed, Periodic,
            # Planar); Order/NumPoles/NumKnots/Closed/Periodic are two-element
            # arrays holding the U value then the V value.
            if len(bspline_info) >= 7:
                result["order"] = self._to_list(bspline_info[0])
                result["num_poles"] = self._to_list(bspline_info[1])
                result["num_knots"] = self._to_list(bspline_info[2])
                result["rational"] = bool(bspline_info[3])
                result["closed"] = [bool(v) for v in self._to_list(bspline_info[4])]
                result["periodic"] = [bool(v) for v in self._to_list(bspline_info[5])]
                result["planar"] = bool(bspline_info[6])
            else:
                # Unexpected shape - store raw info rather than mislabel it.
                result["raw_info"] = list(bspline_info)

            return result
        except Exception as e:
            return error_result(e)
