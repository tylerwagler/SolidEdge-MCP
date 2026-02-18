"""
Test mirror copy and thread creation.
Requires Solid Edge running with an open part that has a base feature.
"""

import sys

sys.path.insert(0, "src")
import traceback

import pythoncom
import win32com.client as win32
from win32com.client import VARIANT

app = win32.Dispatch("SolidEdge.Application")
doc = app.ActiveDocument
models = doc.Models

if models.Count == 0:
    print("ERROR: No model exists. Create a base feature first.")
    exit(1)

model = models.Item(1)

# ============================================================
# TEST: MirrorCopies.AddSync
# ============================================================
print("=" * 60)
print("TEST: MIRROR COPY")
print("=" * 60)

try:
    mc = model.MirrorCopies
    print(f"  MirrorCopies.Count = {mc.Count}")

    # AddSync(NumberOfFeatures, FeatureArray, MirrorPlane, MirrorOption)
    # We need: a feature to mirror and a mirror plane
    # Let's try using a RefPlane as the mirror plane

    # Get DesignEdgebarFeatures for the real feature tree
    features = doc.DesignEdgebarFeatures
    print(f"  DesignEdgebarFeatures.Count = {features.Count}")
    for i in range(1, min(features.Count + 1, 10)):
        f = features.Item(i)
        print(f"    [{i}] Name={f.Name}")

    # Try getting the model as a feature to mirror
    # The last model feature should be what we want
    body = model.Body
    print(f"  Body type: {type(body)}")

    # Get right plane for mirroring
    right_plane = doc.RefPlanes.Item(3)  # Right/YZ plane
    print("  Mirror plane (Right/YZ): obtained")

    # Try to use the entire model body as the feature
    # MirrorCopies.Add(PatternPlane, NumberOfFeatures, FeatureArray)
    print("\n  Trying MirrorCopies.Add(plane, 1, [model])...")
    try:
        mirror = mc.Add(right_plane, 1, [model])
        print("  SUCCESS! Mirror created")
    except Exception as e:
        print(f"  Failed: {e}")

    # Try AddSync
    print("\n  Trying MirrorCopies.AddSync(1, [model], plane, 0)...")
    try:
        # MirrorOption: 0 = default
        mirror = mc.AddSync(1, [model], right_plane, 0)
        print("  SUCCESS! Mirror created with AddSync")
    except Exception as e:
        print(f"  Failed: {e}")

    # Try with the first feature from DesignEdgebarFeatures
    if features.Count > 3:  # Skip ref planes (first 3)
        feat = features.Item(4)  # First real feature
        print(f"\n  Trying with feature '{feat.Name}'...")

        print("  MirrorCopies.Add(plane, 1, [feat])...")
        try:
            mirror = mc.Add(right_plane, 1, [feat])
            print("  SUCCESS! Mirror created")
        except Exception as e:
            print(f"  Failed: {e}")

        # Try with VARIANT wrapper
        print("  With VARIANT wrapper...")
        try:
            v_feats = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [feat])
            mirror = mc.Add(right_plane, 1, v_feats)
            print("  SUCCESS! Mirror created")
        except Exception as e:
            print(f"  Failed: {e}")

except Exception as e:
    print(f"  Error: {e}")
    traceback.print_exc()


# ============================================================
# TEST: Threads.Add
# ============================================================
print("\n" + "=" * 60)
print("TEST: THREADS")
print("=" * 60)

try:
    threads = model.Threads
    print(f"  Threads.Count = {threads.Count}")

    # Threads.Add(HoleData, NumberOfCylinders,
    #   CylinderArray, CylinderEndArray, Reserved1, Reserved2)
    # Need a HoleData and a cylindrical face
    # Let's first check if there are any cylindrical faces

    # Get faces of the body
    body = model.Body
    faces = body.Faces(6)  # igQueryAll = 6
    print(f"  Total faces: {faces.Count}")

    # Check face types
    for i in range(1, min(faces.Count + 1, 20)):
        face = faces.Item(i)
        try:
            ft = face.Type
            area = face.Area
            print(f"    Face[{i}]: Type={ft}, Area={area:.6f}")
        except Exception:
            try:
                area = face.Area
                print(f"    Face[{i}]: Area={area:.6f}")
            except Exception:
                print(f"    Face[{i}]: (cannot read properties)")

except Exception as e:
    print(f"  Error: {e}")
    traceback.print_exc()


# ============================================================
# TEST: Ribs.Add
# ============================================================
print("\n" + "=" * 60)
print("TEST: RIBS")
print("=" * 60)

try:
    ribs = model.Ribs
    print(f"  Ribs.Count = {ribs.Count}")
    # Ribs.Add(RibProfile, ProfileExtensionType,
    #   ThicknessType, MaterialSide, ThicknessSide,
    #   Thickness, FiniteDepth)
    # Need: a sketch profile that defines the rib cross-section
    # This is complex - needs a profile plus various enum constants
    print(
        "  Ribs.Add signature: (RibProfile, "
        "ProfileExtensionType, ThicknessType, "
        "MaterialSide, ThicknessSide, "
        "Thickness, FiniteDepth)"
    )
    print("  Requires specific enum values for extension type and thickness type")
except Exception as e:
    print(f"  Error: {e}")

# ============================================================
# TEST: DeleteFaces
# ============================================================
print("\n" + "=" * 60)
print("TEST: DELETE FACES")
print("=" * 60)

try:
    df = model.DeleteFaces
    print(f"  DeleteFaces.Count = {df.Count}")
    print("  DeleteFaces.Add(FaceSetToDelete) - needs face object from body.Faces()")
except Exception as e:
    print(f"  Error: {e}")

print("\n" + "=" * 60)
print("DONE")
print("=" * 60)
