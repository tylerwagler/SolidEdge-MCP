"""
End-to-end test of the mirror copy tool through backend classes.
Creates a box, a cylindrical protrusion off-center, then mirrors it.
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

# Create a new part
result = doc_mgr.create_part()
print(f"Create part: {result.get('status', result.get('error'))}")

# Create base box (100x100x50mm)
result = feat_mgr.create_box_by_two_points(-0.05, -0.05, 0, 0.05, 0.05, 0.05)
print(f"Create box: {result.get('status', result.get('error'))}")

# Create an off-center protrusion (cylinder sketch at x=0.03)
result = sketch_mgr.create_sketch("Top")
print(f"Create sketch: {result.get('status', result.get('error'))}")

result = sketch_mgr.draw_circle(0.03, 0.0, 0.008)  # 8mm radius, at x=30mm
print(f"Draw circle: {result.get('status', result.get('error'))}")

result = sketch_mgr.close_sketch()
print(f"Close sketch: {result.get('status', result.get('error'))}")

result = feat_mgr.create_extrude(0.02, "Add", "Normal")
print(f"Extrude cylinder: {result}")

# List features to see names
result = query_mgr.list_features()
print("\nFeatures before mirror:")
for f in result.get("features", []):
    print(f"  {f['name']}")

# Mirror the extrusion across the Right/YZ plane (index 3)
# This should create a mirror at x=-0.03
print("\nMirroring 'ExtrudedProtrusion_1' across Right/YZ plane (3)...")
result = feat_mgr.create_mirror("ExtrudedProtrusion_1", 3)
print(f"Mirror: {result}")

# List features after mirror
result = query_mgr.list_features()
print("\nFeatures after mirror:")
for f in result.get("features", []):
    print(f"  {f['name']}")

# Also try mirroring across Front/XY plane (index 2)
print("\nMirroring 'ExtrudedProtrusion_1' across Front/XY plane (2)...")
result = feat_mgr.create_mirror("ExtrudedProtrusion_1", 2)
print(f"Mirror: {result}")

# List features after second mirror
result = query_mgr.list_features()
print("\nFeatures after 2nd mirror:")
for f in result.get("features", []):
    print(f"  {f['name']}")

print("\nDONE - Check Solid Edge for symmetrical protrusions!")
