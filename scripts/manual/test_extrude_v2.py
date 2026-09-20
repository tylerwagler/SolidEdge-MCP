"""Test ExtrudedProtrusions.AddFinite with correct 4 params."""

import sys

sys.path.insert(0, "src")
import win32com.client as win32

app = win32.Dispatch("SolidEdge.Application")
doc = app.ActiveDocument
models = doc.Models
model = models.Item(1)
ep = model.ExtrudedProtrusions

igRight = 2
igLeft = 1

# Create a sketch profile
ps = doc.ProfileSets.Add()
top_plane = doc.RefPlanes.Item(1)
profile = ps.Profiles.Add(top_plane)
profile.Circles2d.AddByCenterRadius(0.03, 0.0, 0.008)
profile.End(0)

# Try AddFinite with 4 params: Profile, ProfileSide, ProfilePlaneSide, Depth
print("Try 1: AddFinite(profile, igRight, igRight, 0.02)...")
try:
    feat = ep.AddFinite(profile, igRight, igRight, 0.02)
    print("  SUCCESS!")
except Exception as e:
    print(f"  Failed: {e}")

    # Try with different profile side
    ps2 = doc.ProfileSets.Add()
    profile2 = ps2.Profiles.Add(top_plane)
    profile2.Circles2d.AddByCenterRadius(0.03, 0.0, 0.008)
    profile2.End(0)

    print("Try 2: AddFinite(profile, igLeft, igRight, 0.02)...")
    try:
        feat = ep.AddFinite(profile2, igLeft, igRight, 0.02)
        print("  SUCCESS!")
    except Exception as e2:
        print(f"  Failed: {e2}")

        # Try with tuple
        ps3 = doc.ProfileSets.Add()
        profile3 = ps3.Profiles.Add(top_plane)
        profile3.Circles2d.AddByCenterRadius(0.03, 0.0, 0.008)
        profile3.End(0)

        print("Try 3: AddFiniteMulti(1, (profile,), igRight, 0.02)...")
        try:
            feat = ep.AddFiniteMulti(1, (profile3,), igRight, 0.02)
            print("  SUCCESS!")
        except Exception as e3:
            print(f"  Failed: {e3}")

            # OK let's try the models-level API that we had before
            ps4 = doc.ProfileSets.Add()
            profile4 = ps4.Profiles.Add(top_plane)
            profile4.Circles2d.AddByCenterRadius(0.03, 0.0, 0.008)
            profile4.End(0)

            print("Try 4: models.AddFiniteExtrudedProtrusion(1, (profile,), igRight, 0.02)...")
            try:
                m = models.AddFiniteExtrudedProtrusion(1, (profile4,), igRight, 0.02)
                print(f"  Result type: {type(m)}")
                print(f"  Models.Count = {models.Count}")
                for i in range(1, models.Count + 1):
                    mdl = models.Item(i)
                    print(f"    Model[{i}]: {mdl.Name}")
            except Exception as e4:
                print(f"  Failed: {e4}")

# Print final state
print("\nDesignEdgebarFeatures:")
for i in range(1, doc.DesignEdgebarFeatures.Count + 1):
    f = doc.DesignEdgebarFeatures.Item(i)
    print(f"  [{i}] {f.Name}")

print(f"\nBody volume: {model.Body.Volume}")
print(f"ExtrudedProtrusions.Count = {ep.Count}")
