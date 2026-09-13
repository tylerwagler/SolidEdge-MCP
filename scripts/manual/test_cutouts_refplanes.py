"""
Investigate ExtrudedCutouts, RevolvedCutouts, RefPlanes, and Threads type signatures.
"""

import traceback

import win32com.client as win32

app = win32.Dispatch("SolidEdge.Application")
doc = app.ActiveDocument
models = doc.Models

if models.Count == 0:
    print("ERROR: No model/feature exists.")
    exit(1)

model = models.Item(1)


def dump_typeinfo(name, obj):
    """Get TypeInfo for a COM collection."""
    print(f"\n{'=' * 60}")
    print(f"{name} TYPE INFO")
    print(f"{'=' * 60}")
    try:
        early = win32.gencache.EnsureDispatch(obj)
        ti = early._oleobj_.GetTypeInfo()
        ta = ti.GetTypeAttr()
        print(f"  GUID: {ta[0]}")
        print(f"  Functions: {ta[6]}")
        for i in range(ta[6]):
            fd = ti.GetFuncDesc(i)
            names = ti.GetNames(fd[0])
            # fd[2] = params count, fd[8] = param types
            param_types = fd[2]
            ret_type = fd[8]  # return type
            invoke_kind = fd[3]  # 1=method, 2=propget, 4=propput
            params_str = ", ".join(names[1:]) if len(names) > 1 else ""
            if names[0].startswith(("Add", "Create")):
                # Get detailed param info
                print(f"  ** [{i}] {names[0]}({params_str})")
                print(f"       invkind={invoke_kind}, nParams={param_types}, retType={ret_type}")
                # Try to get param type info from fd
                if hasattr(fd, "__len__") and len(fd) > 2:
                    cparams = fd[5] if len(fd) > 5 else "?"
                print(f"       Full FuncDesc: invkind={fd[3]}, cParams={cparams}")
            elif not names[0].startswith(("Query", "AddRef", "Release", "_")):
                print(f"  [{i}] {names[0]}({params_str}) invkind={invoke_kind}")
    except Exception as e:
        print(f"  TypeInfo error: {e}")
        traceback.print_exc()


# Investigate ExtrudedCutouts
dump_typeinfo("ExtrudedCutouts", model.ExtrudedCutouts)

# Investigate RevolvedCutouts
dump_typeinfo("RevolvedCutouts", model.RevolvedCutouts)

# Investigate NormalCutouts
dump_typeinfo("NormalCutouts", model.NormalCutouts)

# Investigate Threads
dump_typeinfo("Threads", model.Threads)

# Investigate Slots
dump_typeinfo("Slots", model.Slots)

# Investigate RefPlanes
dump_typeinfo("RefPlanes", doc.RefPlanes)

# Investigate Ribs
dump_typeinfo("Ribs", model.Ribs)

# Investigate DeleteFaces
dump_typeinfo("DeleteFaces", model.DeleteFaces)

# Now try creating a cutout!
print("\n" + "=" * 60)
print("ATTEMPTING EXTRUDED CUTOUT")
print("=" * 60)

try:
    # Create a sketch profile on top of the existing box
    ps = doc.ProfileSets.Add()
    top_plane = doc.RefPlanes.Item(1)  # Top/XZ plane
    profile = ps.Profiles.Add(top_plane)

    # Draw a small circle for the cutout
    circles = profile.Circles2d
    circle = circles.AddByCenterRadius(0.0, 0.0, 0.01)  # 10mm radius hole

    profile.End(0)  # Close profile

    # Try ExtrudedCutouts.AddFinite
    cutouts = model.ExtrudedCutouts
    igRight = 2  # ProfilePlaneSide
    depth = 0.02  # 20mm

    print("  Trying ExtrudedCutouts.AddFinite(profile, igRight, depth)...")
    try:
        cutout = cutouts.AddFinite(profile, igRight, depth)
        print("  SUCCESS! Cutout created")
    except Exception as e:
        print(f"  Failed (3 params): {e}")

        # Try with more params
        print("  Trying AddFinite with 4 params (1, (profile,), igRight, depth)...")
        try:
            cutout = cutouts.AddFinite(1, (profile,), igRight, depth)
            print("  SUCCESS! Cutout created")
        except Exception as e2:
            print(f"  Failed (4 params): {e2}")

            # Try AddFiniteMulti
            print("  Trying AddFiniteMulti(1, (profile,), igRight, depth)...")
            try:
                cutout = cutouts.AddFiniteMulti(1, (profile,), igRight, depth)
                print("  SUCCESS! Cutout created with AddFiniteMulti")
            except Exception as e3:
                print(f"  Failed AddFiniteMulti: {e3}")

except Exception as e:
    print(f"  Error: {e}")
    traceback.print_exc()

# Try creating an offset reference plane
print("\n" + "=" * 60)
print("ATTEMPTING OFFSET REFERENCE PLANE")
print("=" * 60)

try:
    rp = doc.RefPlanes
    top_plane = rp.Item(1)  # Top plane

    print("  Trying AddParallelByDistance(RefPlane, Distance, NormalSide)...")
    try:
        new_plane = rp.AddParallelByDistance(top_plane, 0.05, 2)  # 50mm above, igRight=2
        print(f"  SUCCESS! Offset plane created. Total planes: {rp.Count}")
    except Exception as e:
        print(f"  Failed (3 params): {e}")

        # Try different param counts
        print("  Trying AddParallelByDistance(RefPlane, Distance, NormalSide, bLocal)...")
        try:
            new_plane = rp.AddParallelByDistance(top_plane, 0.05, 2, False)
            print(f"  SUCCESS! Offset plane created. Total planes: {rp.Count}")
        except Exception as e2:
            print(f"  Failed (4 params): {e2}")
except Exception as e:
    print(f"  Error: {e}")
    traceback.print_exc()

print("\n" + "=" * 60)
print("DONE")
print("=" * 60)
