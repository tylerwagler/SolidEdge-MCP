"""
Solid Edge Query and Inspection Operations

Handles querying model data, measurements, and properties.
"""

from typing import Dict, Any
import math
import traceback


class QueryManager:
    """Manages query and inspection operations"""

    def __init__(self, document_manager):
        self.doc_manager = document_manager

    def _get_first_model(self):
        """Get the first model from the active document."""
        doc = self.doc_manager.get_active_document()
        if not hasattr(doc, 'Models'):
            raise Exception("Document does not have a Models collection")
        models = doc.Models
        if models.Count == 0:
            raise Exception("No features in document")
        return doc, models.Item(1)

    def get_mass_properties(self, density: float = 7850) -> Dict[str, Any]:
        """
        Get mass properties of the part.

        Uses Model.ComputePhysicalProperties(status, density, accuracy) which
        returns a tuple: (volume, area, mass, cog_tuple, cov_tuple, moi_tuple, ...)

        Args:
            density: Material density in kg/m³ (default: 7850 for steel)

        Returns:
            Dict with volume, mass, surface area, center of gravity, moments of inertia
        """
        try:
            doc, model = self._get_first_model()

            # ComputePhysicalPropertiesWithSpecifiedDensity(Density, Accuracy)
            # Returns tuple: (volume, area, mass, cog, cov, moi, principal_moi,
            #                  principal_axes, radii_of_gyration, ?, ?)
            result = model.ComputePhysicalPropertiesWithSpecifiedDensity(
                density, 0.99
            )

            volume = result[0] if len(result) > 0 else 0
            surface_area = result[1] if len(result) > 1 else 0
            mass_val = result[2] if len(result) > 2 else 0
            cog = result[3] if len(result) > 3 else (0, 0, 0)
            cov = result[4] if len(result) > 4 else (0, 0, 0)
            moi = result[5] if len(result) > 5 else (0, 0, 0, 0, 0, 0)
            principal_moi = result[6] if len(result) > 6 else (0, 0, 0)

            return {
                "status": "computed",
                "density": density,
                "volume": volume,
                "surface_area": surface_area,
                "mass": mass_val,
                "center_of_gravity": list(cog) if cog else [0, 0, 0],
                "center_of_volume": list(cov) if cov else [0, 0, 0],
                "moments_of_inertia": {
                    "Ixx": moi[0] if len(moi) > 0 else 0,
                    "Iyy": moi[1] if len(moi) > 1 else 0,
                    "Izz": moi[2] if len(moi) > 2 else 0,
                    "Ixy": moi[3] if len(moi) > 3 else 0,
                    "Ixz": moi[4] if len(moi) > 4 else 0,
                    "Iyz": moi[5] if len(moi) > 5 else 0,
                },
                "principal_moments": list(principal_moi) if principal_moi else [0, 0, 0],
                "units": {
                    "volume": "m³",
                    "surface_area": "m²",
                    "mass": "kg",
                    "density": "kg/m³",
                    "moments_of_inertia": "kg·m²",
                    "coordinates": "meters"
                }
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_bounding_box(self) -> Dict[str, Any]:
        """
        Get the bounding box of the model.

        Uses Body.GetRange() which returns ((min_x, min_y, min_z), (max_x, max_y, max_z)).

        Returns:
            Dict with min/max coordinates and dimensions
        """
        try:
            doc, model = self._get_first_model()

            body = model.Body
            range_data = body.GetRange()

            min_pt = range_data[0]
            max_pt = range_data[1]

            return {
                "status": "computed",
                "min": list(min_pt),
                "max": list(max_pt),
                "dimensions": {
                    "x": max_pt[0] - min_pt[0],
                    "y": max_pt[1] - min_pt[1],
                    "z": max_pt[2] - min_pt[2]
                },
                "units": "meters"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def list_features(self) -> Dict[str, Any]:
        """
        List all features in the active document.

        Uses Model.Features collection and DesignEdgebarFeatures for the feature tree.

        Returns:
            Dict with list of features
        """
        try:
            doc, model = self._get_first_model()

            features = []

            # Use DesignEdgebarFeatures for the full feature tree
            if hasattr(doc, 'DesignEdgebarFeatures'):
                debf = doc.DesignEdgebarFeatures
                for i in range(1, debf.Count + 1):
                    try:
                        feat = debf.Item(i)
                        feat_info = {
                            "index": i - 1,
                            "name": feat.Name if hasattr(feat, 'Name') else f"Feature_{i}",
                        }
                        features.append(feat_info)
                    except Exception:
                        features.append({"index": i - 1, "name": f"Feature_{i}"})
            else:
                # Fallback to Model.Features
                model_features = model.Features
                for i in range(1, model_features.Count + 1):
                    try:
                        feat = model_features.Item(i)
                        feat_info = {
                            "index": i - 1,
                            "name": feat.Name if hasattr(feat, 'Name') else f"Feature_{i}",
                        }
                        features.append(feat_info)
                    except Exception:
                        features.append({"index": i - 1, "name": f"Feature_{i}"})

            return {
                "features": features,
                "count": len(features)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def measure_distance(self, x1: float, y1: float, z1: float,
                        x2: float, y2: float, z2: float) -> Dict[str, Any]:
        """
        Measure distance between two points.

        Args:
            x1, y1, z1: First point coordinates
            x2, y2, z2: Second point coordinates

        Returns:
            Dict with distance and components
        """
        try:
            dx = x2 - x1
            dy = y2 - y1
            dz = z2 - z1

            distance = math.sqrt(dx**2 + dy**2 + dz**2)

            return {
                "distance": distance,
                "delta": {
                    "x": dx,
                    "y": dy,
                    "z": dz
                },
                "point1": [x1, y1, z1],
                "point2": [x2, y2, z2],
                "units": "meters"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_document_properties(self) -> Dict[str, Any]:
        """Get document properties and metadata"""
        try:
            doc = self.doc_manager.get_active_document()

            properties = {
                "name": doc.Name if hasattr(doc, 'Name') else "Unknown",
                "path": doc.FullName if hasattr(doc, 'FullName') else "Unsaved",
                "modified": not doc.Saved if hasattr(doc, 'Saved') else False,
                "read_only": doc.ReadOnly if hasattr(doc, 'ReadOnly') else False
            }

            # Try to get summary info
            try:
                if hasattr(doc, 'SummaryInfo'):
                    summary = doc.SummaryInfo
                    if hasattr(summary, 'Title'):
                        properties["title"] = summary.Title
                    if hasattr(summary, 'Author'):
                        properties["author"] = summary.Author
                    if hasattr(summary, 'Subject'):
                        properties["subject"] = summary.Subject
                    if hasattr(summary, 'Comments'):
                        properties["comments"] = summary.Comments
            except:
                pass

            # Add body topology info
            try:
                models = doc.Models
                if models.Count > 0:
                    body = models.Item(1).Body
                    properties["volume_m3"] = body.Volume
            except:
                pass

            return properties
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_feature_count(self) -> Dict[str, Any]:
        """Get count of features in the document"""
        try:
            doc = self.doc_manager.get_active_document()

            counts = {}

            if hasattr(doc, 'DesignEdgebarFeatures'):
                counts["features"] = doc.DesignEdgebarFeatures.Count

            if hasattr(doc, 'Models'):
                counts["models"] = doc.Models.Count

            if hasattr(doc, 'ProfileSets'):
                counts["sketches"] = doc.ProfileSets.Count

            if hasattr(doc, 'RefPlanes'):
                counts["ref_planes"] = doc.RefPlanes.Count

            if hasattr(doc, 'Variables'):
                counts["variables"] = doc.Variables.Count

            return counts
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }
