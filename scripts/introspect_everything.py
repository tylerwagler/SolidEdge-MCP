#!/usr/bin/env python
"""
Comprehensive introspection of ALL Solid Edge COM API methods
Discovers signatures, parameters, and available methods
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))

import inspect

from solidedge_mcp.backends.connection import SolidEdgeConnection
from solidedge_mcp.backends.documents import DocumentManager


def print_section(title):
    print("\n" + "=" * 80)
    print(f"  {title}")
    print("=" * 80)


def introspect_method(obj, method_name, show_params=True):
    """Introspect a single method and print details"""
    try:
        method = getattr(obj, method_name)
        print(f"\n{method_name}")
        print("-" * 80)

        # Get signature
        try:
            sig = inspect.signature(method)
            params = list(sig.parameters.items())
            print(f"Parameters: {len(params)}")
            print(f"Signature: {sig}")

            if show_params and params:
                print("\nParameter Details:")
                for param_name, param in params:
                    kind = param.kind.name
                    default = type(param.default).__name__
                    print(f"  - {param_name}: kind={kind}, default={default}")
        except Exception as e:
            print(f"Could not get signature: {e}")

        # Get docstring
        doc = method.__doc__
        if doc:
            print(f"Docstring: {doc[:200]}")
        else:
            print("Docstring: None")

    except Exception as e:
        print(f"Error introspecting {method_name}: {e}")


# Initialize
connection = SolidEdgeConnection()
connection.connect(start_if_needed=True)
doc_manager = DocumentManager(connection)
doc_manager.create_part()
doc = doc_manager.get_active_document()
app = connection.application

print("=" * 80)
print("COMPREHENSIVE SOLID EDGE COM API INTROSPECTION")
print("=" * 80)
print("\nThis will take a few minutes to examine all available methods...")

# ============================================================================
# 1. MODELS COLLECTION - ALL "ADD" METHODS
# ============================================================================
print_section("MODELS COLLECTION - ALL 'Add' METHODS")

models = doc.Models
add_methods = sorted([m for m in dir(models) if m.startswith("Add")])
print(f"\nFound {len(add_methods)} 'Add' methods\n")

for method_name in add_methods:
    introspect_method(models, method_name, show_params=False)

# ============================================================================
# 2. MODELS COLLECTION - OTHER METHODS
# ============================================================================
print_section("MODELS COLLECTION - OTHER METHODS")

other_methods = sorted(
    [m for m in dir(models) if not m.startswith("_") and not m.startswith("Add")]
)
print(f"\nFound {len(other_methods)} other methods/properties\n")

# Just list them for now
for i, method in enumerate(other_methods):
    if i % 5 == 0:
        print()
    print(f"{method:30s}", end="")
print("\n")

# ============================================================================
# 3. WINDOW OBJECT - VIEW METHODS
# ============================================================================
print_section("WINDOW OBJECT - ALL METHODS")

try:
    window = app.ActiveWindow
    window_methods = sorted([m for m in dir(window) if not m.startswith("_")])
    print(f"\nFound {len(window_methods)} methods/properties\n")

    # Introspect key methods
    key_window_methods = [
        m
        for m in window_methods
        if any(x in m.lower() for x in ["view", "orient", "zoom", "fit", "display"])
    ]
    print(f"\nKey view-related methods ({len(key_window_methods)}):")
    for method_name in key_window_methods:
        introspect_method(window, method_name, show_params=True)

    # List all others
    print("\n\nAll window methods/properties:")
    for i, method in enumerate(window_methods):
        if i % 4 == 0:
            print()
        print(f"{method:35s}", end="")
    print("\n")
except Exception as e:
    print(f"Error accessing Window: {e}")

# ============================================================================
# 4. VIEW OBJECT - ALL METHODS
# ============================================================================
print_section("VIEW OBJECT - ALL METHODS")

try:
    view = app.ActiveWindow.View
    view_methods = sorted([m for m in dir(view) if not m.startswith("_")])
    print(f"\nFound {len(view_methods)} methods/properties\n")

    # Introspect all non-property methods
    for method_name in view_methods:
        try:
            attr = getattr(view, method_name)
            # Only introspect if it looks like a method (callable)
            if callable(attr):
                introspect_method(view, method_name, show_params=True)
        except Exception:
            pass

    # List all
    print("\n\nAll view methods/properties:")
    for i, method in enumerate(view_methods):
        if i % 4 == 0:
            print()
        print(f"{method:35s}", end="")
    print("\n")
except Exception as e:
    print(f"Error accessing View: {e}")

# ============================================================================
# 5. DOCUMENT OBJECT - ALL METHODS
# ============================================================================
print_section("DOCUMENT OBJECT - ALL METHODS")

doc_methods = sorted([m for m in dir(doc) if not m.startswith("_")])
print(f"\nFound {len(doc_methods)} methods/properties\n")

# Show key methods
key_doc_methods = [
    m
    for m in doc_methods
    if any(
        x in m.lower()
        for x in [
            "save",
            "close",
            "export",
            "property",
            "model",
            "profile",
        ]
    )
]
print(f"Key document methods ({len(key_doc_methods)}):")
for i, method in enumerate(key_doc_methods):
    if i % 4 == 0:
        print()
    print(f"{method:35s}", end="")
print("\n")

# ============================================================================
# 6. PROFILES COLLECTION
# ============================================================================
print_section("PROFILE COLLECTION - METHODS")

try:
    ref_planes = doc.RefPlanes
    ref_plane = ref_planes.Item(1)
    profile_sets = doc.ProfileSets
    profile_set = profile_sets.Add()
    profiles = profile_set.Profiles
    profile = profiles.Add(ref_plane)

    profile_methods = sorted([m for m in dir(profile) if not m.startswith("_")])
    print(f"\nFound {len(profile_methods)} methods/properties on Profile\n")

    # List all
    for i, method in enumerate(profile_methods):
        if i % 4 == 0:
            print()
        print(f"{method:35s}", end="")
    print("\n")

    # Introspect 2D geometry collections
    print("\n2D GEOMETRY COLLECTIONS:")
    for geom_type in ["Lines2d", "Circles2d", "Arcs2d", "Ellipses2d", "Splines2d"]:
        try:
            collection = getattr(profile, geom_type)
            methods = sorted([m for m in dir(collection) if m.startswith("Add")])
            print(f"\n{geom_type}: {methods}")
            for method_name in methods:
                introspect_method(collection, method_name, show_params=True)
        except Exception as e:
            print(f"Error accessing {geom_type}: {e}")

except Exception as e:
    print(f"Error accessing Profile: {e}")

# ============================================================================
# 7. CONSTANTS - ENUMERATE ALL
# ============================================================================
print_section("CONSTANTS - AVAILABLE ENUMS")

try:
    from win32com.client import constants

    constant_names = sorted([c for c in dir(constants) if not c.startswith("_")])
    print(f"\nFound {len(constant_names)} constants\n")

    # Group by prefix
    prefixes = {}
    for name in constant_names:
        prefix = name.split("_")[0] if "_" in name else name[:2]
        if prefix not in prefixes:
            prefixes[prefix] = []
        prefixes[prefix].append(name)

    # Show counts by prefix
    print("Constants by prefix:")
    for prefix, names in sorted(prefixes.items(), key=lambda x: -len(x[1]))[:20]:
        print(f"  {prefix}: {len(names)} constants")

    # Show some key constant groups
    print("\n\nKey constant groups:")
    for keyword in ["ViewOrientation", "DisplayStyle", "FeatureOperation", "ExtentType", "Extrude"]:
        matching = [c for c in constant_names if keyword.lower() in c.lower()]
        if matching:
            print(f"\n{keyword} constants ({len(matching)}):")
            for const in matching[:10]:  # First 10
                try:
                    value = getattr(constants, const)
                    print(f"  {const} = {value}")
                except Exception:
                    pass
            if len(matching) > 10:
                print(f"  ... and {len(matching) - 10} more")

except Exception as e:
    print(f"Error accessing constants: {e}")

# ============================================================================
# 8. FEATURE-SPECIFIC INTROSPECTION
# ============================================================================
print_section("SPECIFIC FEATURE METHODS - DETAILED")

print("\nAddFiniteRevolvedProtrusion:")
introspect_method(models, "AddFiniteRevolvedProtrusion", show_params=True)

print("\nAddBoxByTwoPoints:")
introspect_method(models, "AddBoxByTwoPoints", show_params=True)

print("\nAddRevolvedProtrusionSync:")
introspect_method(models, "AddRevolvedProtrusionSync", show_params=True)

print("\nAddLoftedProtrusion:")
introspect_method(models, "AddLoftedProtrusion", show_params=True)

print("\nAddSweptProtrusion:")
introspect_method(models, "AddSweptProtrusion", show_params=True)

# ============================================================================
# CLEANUP
# ============================================================================
doc_manager.close_document(save=False)

print("\n" + "=" * 80)
print("INTROSPECTION COMPLETE")
print("=" * 80)
print("\nOutput saved to console. Review the signatures and parameter details above.")
print("=" * 80)
