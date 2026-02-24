"""
Diagnostic tools for Solid Edge API exploration
"""

from typing import Any


def get_available_methods(obj: Any, filter_prefix: str | None = None) -> dict[str, Any]:
    """
    Get all available methods and properties on a COM object

    Args:
        obj: COM object to inspect
        filter_prefix: Optional prefix to filter methods (e.g., "Add")

    Returns:
        Dictionary with methods and properties
    """
    methods: list[str] = []
    properties: list[str] = []

    try:
        # Get all attributes
        for attr_name in dir(obj):
            if filter_prefix and not attr_name.startswith(filter_prefix):
                continue

            try:
                attr = getattr(obj, attr_name)
                if callable(attr):
                    methods.append(attr_name)
                else:
                    properties.append(attr_name)
            except Exception:
                pass
    except Exception:
        pass

    return {
        "methods": sorted(methods),
        "properties": sorted(properties),
        "total_methods": len(methods),
        "total_properties": len(properties),
    }


def diagnose_document(doc: Any) -> dict[str, Any]:
    """
    Diagnose available features and collections in a document

    Args:
        doc: Solid Edge document object

    Returns:
        Dictionary with available collections and methods
    """
    info: dict[str, Any] = {
        "document_type": type(doc).__name__,
        "available_collections": [],
        "models_methods": [],
        "cutout_related_methods": [],
    }

    # Check for common collections
    collection_names = [
        "Models",
        "ExtrudedCutouts",
        "Cutouts",
        "Features",
        "ProfileSets",
        "Profiles",
        "ExtrudedProtrusions",
        "Holes",
        "Rounds",
        "Chamfers",
        "Patterns",
        "RibWebs",
        "Threads",
        "Constructions",
        "RefPlanes",
        "UserDefinedPatterns",
        "Assemblies",
        "Occurrences",
        "SolidEdgePart",
        "Sketches",
    ]

    for name in collection_names:
        if hasattr(doc, name):
            info["available_collections"].append(name)
            try:
                collection = getattr(doc, name)
                # Get Add methods for this collection
                add_methods = get_available_methods(collection, "Add")
                info[f"{name}_add_methods"] = add_methods["methods"]
            except Exception:
                pass

    # Get all methods on Models collection
    if hasattr(doc, "Models"):
        models = doc.Models
        all_methods = get_available_methods(models)
        info["models_methods"] = all_methods["methods"]

        # Filter cutout-related
        cutout_methods = [
            m
            for m in all_methods["methods"]
            if "cutout" in m.lower() or "cut" in m.lower()
        ]
        info["cutout_related_methods"] = cutout_methods

    return info


def diagnose_feature(model: Any) -> dict[str, Any]:
    """
    Diagnose properties and methods available on a feature/model object.

    Args:
        model: Solid Edge Model/Feature object

    Returns:
        Dictionary with available properties, methods, and their values
    """
    info: dict[str, Any] = {
        "model_type": type(model).__name__,
        "properties": {},
        "all_attributes": [],
        "operation_related": [],
    }

    # Properties to check
    property_names = [
        "Name",
        "Type",
        "Visible",
        "Suppressed",
        "FeatureType",
        "Operation",
        "OperationType",
        "SideStep",
        "ExtrusionType",
        "ProfileSide",
        "ProfilePlaneSide",
        "KeypointType",
        "FeatureOperationType",
    ]

    for prop_name in property_names:
        try:
            value = getattr(model, prop_name, None)
            if value is not None:
                info["properties"][prop_name] = str(value)
        except Exception as e:
            info["properties"][prop_name] = f"Error: {str(e)}"

    # Get all attributes
    all_attrs = get_available_methods(model)
    info["all_attributes"] = all_attrs["methods"] + all_attrs["properties"]

    # Find operation-related attributes
    keywords = ["operation", "side", "type", "cut", "add", "subtract", "feature"]
    operation_attrs = [
        attr
        for attr in info["all_attributes"]
        if any(keyword in attr.lower() for keyword in keywords)
    ]
    info["operation_related"] = operation_attrs

    return info
