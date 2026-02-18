#!/usr/bin/env python
"""Show exactly what Python runtime introspection reveals about Solid Edge COM API"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))

import inspect

from solidedge_mcp.backends.connection import SolidEdgeConnection
from solidedge_mcp.backends.documents import DocumentManager

connection = SolidEdgeConnection()
connection.connect(start_if_needed=True)
doc_manager = DocumentManager(connection)
doc_manager.create_part()
doc = doc_manager.get_active_document()

models = doc.Models

print("=" * 80)
print("RUNTIME INTROSPECTION OF SOLID EDGE COM API")
print("=" * 80)

# ============================================================================
# 1. REVOLVE METHOD
# ============================================================================
print("\n" + "=" * 80)
print("AddRevolvedProtrusion")
print("=" * 80)

print("\n1. Signature:")
try:
    sig = inspect.signature(models.AddRevolvedProtrusion)
    print(f"{sig}")
except Exception as e:
    print(f"Error getting signature: {e}")

print("\n2. Parameters (detailed):")
try:
    sig = inspect.signature(models.AddRevolvedProtrusion)
    for param_name, param in sig.parameters.items():
        print(f"  {param_name}:")
        print(f"    - default: {param.default}")
        print(f"    - kind: {param.kind}")
        print(f"    - annotation: {param.annotation}")
except Exception as e:
    print(f"Error: {e}")

print("\n3. Docstring:")
print(models.AddRevolvedProtrusion.__doc__)

print("\n4. Type:")
print(type(models.AddRevolvedProtrusion))

# ============================================================================
# 2. BOX METHOD
# ============================================================================
print("\n" + "=" * 80)
print("AddBoxByCenter")
print("=" * 80)

if hasattr(models, "AddBoxByCenter"):
    print("\n1. Signature:")
    try:
        sig = inspect.signature(models.AddBoxByCenter)
        print(f"{sig}")
    except Exception as e:
        print(f"Error: {e}")

    print("\n2. Parameters (detailed):")
    try:
        sig = inspect.signature(models.AddBoxByCenter)
        for param_name, param in sig.parameters.items():
            print(f"  {param_name}:")
            print(f"    - default: {param.default}")
            print(f"    - kind: {param.kind}")
    except Exception as e:
        print(f"Error: {e}")

    print("\n3. Docstring:")
    print(models.AddBoxByCenter.__doc__)
else:
    print("Method does not exist!")
    print("\nSearching for 'Box' methods:")
    for attr in dir(models):
        if "Box" in attr:
            print(f"  - {attr}")

# ============================================================================
# 3. CYLINDER METHOD
# ============================================================================
print("\n" + "=" * 80)
print("AddCylinderByCenterAndRadius")
print("=" * 80)

if hasattr(models, "AddCylinderByCenterAndRadius"):
    print("\n1. Signature:")
    try:
        sig = inspect.signature(models.AddCylinderByCenterAndRadius)
        print(f"{sig}")
    except Exception as e:
        print(f"Error: {e}")

    print("\n2. Parameters (detailed):")
    try:
        sig = inspect.signature(models.AddCylinderByCenterAndRadius)
        for param_name, param in sig.parameters.items():
            print(f"  {param_name}:")
            print(f"    - default: {param.default}")
            print(f"    - kind: {param.kind}")
    except Exception as e:
        print(f"Error: {e}")

    print("\n3. Docstring:")
    print(models.AddCylinderByCenterAndRadius.__doc__)
else:
    print("Method does not exist!")
    print("\nSearching for 'Cylinder' methods:")
    for attr in dir(models):
        if "Cylinder" in attr:
            print(f"  - {attr}")

# ============================================================================
# 4. SPHERE METHOD
# ============================================================================
print("\n" + "=" * 80)
print("AddSphereByCenterAndRadius")
print("=" * 80)

if hasattr(models, "AddSphereByCenterAndRadius"):
    print("\n1. Signature:")
    try:
        sig = inspect.signature(models.AddSphereByCenterAndRadius)
        print(f"{sig}")
    except Exception as e:
        print(f"Error: {e}")

    print("\n2. Parameters (detailed):")
    try:
        sig = inspect.signature(models.AddSphereByCenterAndRadius)
        for param_name, param in sig.parameters.items():
            print(f"  {param_name}:")
            print(f"    - default: {param.default}")
            print(f"    - kind: {param.kind}")
    except Exception as e:
        print(f"Error: {e}")

    print("\n3. Docstring:")
    print(models.AddSphereByCenterAndRadius.__doc__)
else:
    print("Method does not exist!")
    print("\nSearching for 'Sphere' methods:")
    for attr in dir(models):
        if "Sphere" in attr:
            print(f"  - {attr}")

# ============================================================================
# 5. COMPARISON - WORKING EXTRUDE METHOD
# ============================================================================
print("\n" + "=" * 80)
print("AddFiniteExtrudedProtrusion (WORKING - for comparison)")
print("=" * 80)

print("\n1. Signature:")
try:
    sig = inspect.signature(models.AddFiniteExtrudedProtrusion)
    print(f"{sig}")
except Exception as e:
    print(f"Error: {e}")

print("\n2. Parameters (detailed) - showing first 5:")
try:
    sig = inspect.signature(models.AddFiniteExtrudedProtrusion)
    for i, (param_name, param) in enumerate(sig.parameters.items()):
        if i >= 5:
            print(f"  ... ({len(sig.parameters)} total parameters)")
            break
        print(f"  {param_name}:")
        print(f"    - default: {param.default}")
        print(f"    - kind: {param.kind}")
except Exception as e:
    print(f"Error: {e}")

# ============================================================================
# 6. LIST ALL AVAILABLE "Add" METHODS ON MODELS
# ============================================================================
print("\n" + "=" * 80)
print("ALL 'Add' METHODS ON Models COLLECTION")
print("=" * 80)

add_methods = [attr for attr in dir(models) if attr.startswith("Add")]
print(f"\nFound {len(add_methods)} 'Add' methods:\n")

for method in sorted(add_methods):
    print(f"  - {method}")

doc_manager.close_document(save=False)

print("\n" + "=" * 80)
print("INTROSPECTION COMPLETE")
print("=" * 80)
