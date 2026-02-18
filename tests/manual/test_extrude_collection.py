"""
Test extruded protrusion via collection-level API (model.ExtrudedProtrusions).
"""

import sys

sys.path.insert(0, "src")

import win32com.client as win32

app = win32.Dispatch("SolidEdge.Application")
doc = app.ActiveDocument
models = doc.Models

if models.Count == 0:
    print("ERROR: No model exists")
    exit(1)

model = models.Item(1)

# Get ExtrudedProtrusions type info
print("ExtrudedProtrusions collection:")
ep = model.ExtrudedProtrusions
print(f"  Count = {ep.Count}")

# Get type info
try:
    early = win32.gencache.EnsureDispatch(ep)
    ti = early._oleobj_.GetTypeInfo()
    ta = ti.GetTypeAttr()
    print(f"  GUID: {ta[0]}")
    print(f"  Functions: {ta[6]}")
    for i in range(ta[6]):
        fd = ti.GetFuncDesc(i)
        names = ti.GetNames(fd[0])
        if names[0].startswith(("Add", "Create")):
            params_str = ", ".join(names[1:]) if len(names) > 1 else ""
            print(f"  ** {names[0]}({params_str})")
except Exception as e:
    print(f"  TypeInfo error: {e}")

# Create a sketch profile
ps = doc.ProfileSets.Add()
top_plane = doc.RefPlanes.Item(1)
profile = ps.Profiles.Add(top_plane)
circles = profile.Circles2d
circle = circles.AddByCenterRadius(0.03, 0.0, 0.008)
profile.End(0)

igRight = 2  # Material side

# Try AddFinite
print("\nTrying ExtrudedProtrusions.AddFinite(profile, igRight, 0.02)...")
try:
    feat = ep.AddFinite(profile, igRight, 0.02)
    print("  SUCCESS! Feature created")
except Exception as e:
    print(f"  Failed: {e}")

# Try AddFiniteMulti (like cutouts use)
print("\nTrying AddFiniteMulti(1, (profile,), igRight, 0.02)...")
try:
    # Re-create profile since previous attempt may have consumed it
    ps2 = doc.ProfileSets.Add()
    profile2 = ps2.Profiles.Add(top_plane)
    circles2 = profile2.Circles2d
    circle2 = circles2.AddByCenterRadius(-0.03, 0.0, 0.008)
    profile2.End(0)

    feat = ep.AddFiniteMulti(1, (profile2,), igRight, 0.02)
    print("  SUCCESS! Feature created")
except Exception as e:
    print(f"  Failed: {e}")

# Check features now
print(f"\nModels.Count = {models.Count}")
model = models.Item(1)
print(f"Features.Count = {model.Features.Count}")
for j in range(1, model.Features.Count + 1):
    f = model.Features.Item(j)
    print(f"  [{j}] {f.Name}")

print(f"\nDesignEdgebarFeatures.Count = {doc.DesignEdgebarFeatures.Count}")
for i in range(1, doc.DesignEdgebarFeatures.Count + 1):
    f = doc.DesignEdgebarFeatures.Item(i)
    print(f"  [{i}] {f.Name}")

print(f"\nBody volume: {model.Body.Volume}")
