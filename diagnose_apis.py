#!/usr/bin/env python
"""
Quick diagnostic script to discover correct COM API signatures.
"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from solidedge_mcp.backends.connection import SolidEdgeConnection
from solidedge_mcp.backends.documents import DocumentManager

# Connect
connection = SolidEdgeConnection()
connection.connect(start_if_needed=True)
doc_manager = DocumentManager(connection)

# Create a part
doc_manager.create_part()
doc = doc_manager.get_active_document()

print("=" * 70)
print("DIAGNOSING COM API SIGNATURES")
print("=" * 70)

# Check Models collection
print("\n1. Models collection methods:")
models = doc.Models
for attr in dir(models):
    if attr.startswith('Add'):
        print(f"  - {attr}")

# Check Profile methods for ellipse and spline
print("\n2. Creating a test sketch to check Profile methods...")
ref_planes = doc.RefPlanes
ref_plane = ref_planes.Item(1)
profile_sets = doc.ProfileSets
profile_set = profile_sets.Add()
profiles = profile_set.Profiles
profile = profiles.Add(ref_plane)

print("\n3. Profile.Ellipses2d methods:")
ellipses = profile.Ellipses2d
for attr in dir(ellipses):
    if attr.startswith('Add'):
        print(f"  - {attr}")

print("\n4. Profile.BSplineCurves2d methods:")
splines = profile.BSplineCurves2d
for attr in dir(splines):
    if attr.startswith('Add'):
        print(f"  - {attr}")

# Try to get method signatures using Python's help
print("\n5. Checking AddByCenterRadii signature...")
try:
    import inspect
    if hasattr(ellipses, 'AddByCenterRadii'):
        # Can't get signature from COM objects easily, but we can try
        print("  Method exists!")
except Exception as e:
    print(f"  Error: {e}")

print("\n6. Testing primitive creation signatures...")
print("  Testing AddBoxByCenter...")
try:
    # Try different signatures
    result = models.AddBoxByCenter(
        Origin=[0, 0, 0],
        Length=0.1,
        Width=0.1,
        Height=0.1
    )
    print("  SUCCESS with keyword args!")
except Exception as e:
    print(f"  Failed with keywords: {e}")
    try:
        # Try positional
        result = models.AddBoxByCenter(0, 0, 0, 0.1, 0.1, 0.1)
        print("  SUCCESS with positional args!")
    except Exception as e2:
        print(f"  Failed with positional: {e2}")

print("\n7. Checking ViewOrientationConstants...")
try:
    from solidedge_mcp.backends.constants import ViewOrientationConstants
    print("  Constants defined:")
    for attr in dir(ViewOrientationConstants):
        if not attr.startswith('_'):
            print(f"    - {attr} = {getattr(ViewOrientationConstants, attr)}")
except Exception as e:
    print(f"  Error: {e}")

print("\n" + "=" * 70)
print("DIAGNOSIS COMPLETE")
print("=" * 70)
