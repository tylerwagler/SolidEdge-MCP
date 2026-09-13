"""
Test mirror with a protrusion that extends outside the box.
"""

import sys

sys.path.insert(0, "src")

from solidedge_mcp.backends.connection import SolidEdgeConnection
from solidedge_mcp.backends.documents import DocumentManager
from solidedge_mcp.backends.features import FeatureManager
from solidedge_mcp.backends.query import QueryManager
from solidedge_mcp.backends.sketching import SketchManager

conn = SolidEdgeConnection()
conn.connect(True)
doc_mgr = DocumentManager(conn)
sketch_mgr = SketchManager(doc_mgr)
feat_mgr = FeatureManager(doc_mgr, sketch_mgr)
query_mgr = QueryManager(doc_mgr)

# Create a new part with a base box
doc_mgr.create_part()

# Box: 100mm x 100mm x 50mm centered at origin on XZ plane
# The box extends from (x=-0.05, z=-0.05) to (x=0.05, z=0.05) on the top plane
# And from y=0 to y=0.05 vertically
feat_mgr.create_box_by_two_points(-0.05, 0, -0.05, 0.05, 0.05, 0.05)
print("Box created")

# Create a protrusion OUTSIDE the box footprint
# Use the Front/XY plane (plane 2) for the sketch
# On Front plane: x goes left-right, y goes up-down
# Draw a circle at the right side of the box (x=0.05, y=0.025) extending beyond
result = sketch_mgr.create_sketch("Front")
print(f"Sketch: {result.get('status', result.get('error'))}")

# Draw a rectangle protrusion extending from the right face
# The box right face is at x=0.05 on Front plane
result = sketch_mgr.draw_rectangle(0.05, 0.01, 0.08, 0.04)  # Extends 30mm from the face
print(f"Rectangle: {result.get('status', result.get('error'))}")

result = sketch_mgr.close_sketch()
print(f"Close: {result.get('status', result.get('error'))}")

# Extrude this - on the Front plane, Normal direction extrudes in Z direction
# The box spans z=-0.05 to z=0.05, so extrude 0.1m in "Symmetric" to cover full width
result = feat_mgr.create_extrude(0.1, "Add", "Symmetric")
print(f"Extrude: {result}")

# Check features
result = query_mgr.list_features()
print("\nFeatures:")
for f in result.get("features", []):
    print(f"  {f['name']}")

# If there's an ExtrudedProtrusion, mirror it
feature_names = [f["name"] for f in result.get("features", [])]
protrusion_name = None
for name in feature_names:
    if "Protrusion" in name or "Extrude" in name:
        protrusion_name = name
        break

if protrusion_name:
    print(f"\nMirroring '{protrusion_name}' across Front/XY plane (2)...")
    result = feat_mgr.create_mirror(protrusion_name, 2)
    print(f"Mirror: {result}")

    result = query_mgr.list_features()
    print("\nFeatures after mirror:")
    for f in result.get("features", []):
        print(f"  {f['name']}")
else:
    print("\nNo protrusion feature found to mirror.")
    # Try creating using a different approach: sketch on Right plane
    print("Trying with Right plane sketch...")

    result = sketch_mgr.create_sketch("Right")
    print(f"Sketch on Right: {result.get('status', result.get('error'))}")

    # On Right/YZ plane: y goes left-right, z goes up-down
    # Draw a tab extending from the top of the box
    result = sketch_mgr.draw_rectangle(0.01, 0.05, 0.04, 0.08)
    print(f"Rectangle: {result.get('status', result.get('error'))}")

    result = sketch_mgr.close_sketch()
    print(f"Close: {result.get('status', result.get('error'))}")

    result = feat_mgr.create_extrude(0.1, "Add", "Symmetric")
    print(f"Extrude: {result}")

    result = query_mgr.list_features()
    print("\nFeatures:")
    for f in result.get("features", []):
        print(f"  {f['name']}")

print("\nDone!")
