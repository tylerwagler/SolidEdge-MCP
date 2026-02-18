"""
Investigate feature collections on model objects.
Tests which collections are accessible and what Add methods they support.
Requires Solid Edge running with an open part that has a base feature.
"""

import win32com.client as win32

app = win32.Dispatch("SolidEdge.Application")
doc = app.ActiveDocument
models = doc.Models

if models.Count == 0:
    print("ERROR: No model/feature exists. Create a base feature first.")
    exit(1)

model = models.Item(1)
body = model.Body

# Collections to investigate
collections_to_check = [
    "MirrorCopies",
    "Drafts",
    "Threads",
    "Ribs",
    "Slots",
    "Blends",
    "EmbossFeatures",
    "Lips",
    "Splits",
    "Thickens",
    "CopySurfaces",
    "DeleteFaces",
    "Dimples",
    "DrawnCutouts",
    "FlatPatterns",
    "Gussets",
    "Louvers",
    "NormalCutouts",
    "TappedHoles",
    "WebFeatures",
    "WrapFeatures",
    "SmartFrames",
    "SmartPatterns",
    "BendReliefs",
    "Beads",
    "ExtrudedCutouts",
    "RevolvedCutouts",
    "CutoutPatterns",
    "FeaturePatterns",
]

print("=" * 60)
print("FEATURE COLLECTION INVESTIGATION")
print("=" * 60)

found_collections = {}

for name in collections_to_check:
    try:
        coll = getattr(model, name)
        count = coll.Count if hasattr(coll, "Count") else "?"

        # Get all available methods
        methods = []
        for attr in dir(coll):
            if attr.startswith("Add") or attr.startswith("Create"):
                methods.append(attr)

        print(f"\n  {name}: Count={count}")
        if methods:
            print(f"    Add/Create methods: {methods}")
        else:
            print("    No Add/Create methods found in dir()")
            # Try common method names anyway
            for m in ["Add", "AddFinite", "AddSimple", "AddByRectangular", "AddByCircular"]:
                try:
                    _ = getattr(coll, m)
                    methods.append(m)
                except Exception:
                    pass
            if methods:
                print(f"    But accessible via getattr: {methods}")

        found_collections[name] = methods
    except Exception as e:
        err = str(e)
        if len(err) > 80:
            err = err[:80] + "..."
        print(f"\n  {name}: NOT ACCESSIBLE - {err}")

# Also check Relations3d on assembly
print("\n" + "=" * 60)
print("RELATIONS3D INVESTIGATION (if assembly)")
print("=" * 60)

try:
    rels = doc.Relations3d
    print(f"  Relations3d: Count={rels.Count}")
    methods = [m for m in dir(rels) if m.startswith("Add")]
    print(f"  Add methods: {methods}")
except Exception as e:
    print(f"  Not available (not assembly): {e}")

# Check for RefPlanes features
print("\n" + "=" * 60)
print("REFERENCE PLANE CREATION")
print("=" * 60)

try:
    rp = doc.RefPlanes
    print(f"  RefPlanes: Count={rp.Count}")
    methods = [m for m in dir(rp) if m.startswith("Add")]
    print(f"  Add methods: {methods}")
except Exception as e:
    print(f"  Error: {e}")

# Check for Variables
print("\n" + "=" * 60)
print("VARIABLES")
print("=" * 60)

try:
    vrs = doc.Variables
    print(f"  Variables: Count={vrs.Count}")
    for i in range(1, min(vrs.Count + 1, 10)):
        v = vrs.Item(i)
        try:
            print(f"    [{i}] {v.Name} = {v.Value}")
        except Exception:
            try:
                print(f"    [{i}] Formula={v.Formula}")
            except Exception:
                print(f"    [{i}] (cannot read)")
except Exception as e:
    print(f"  Error: {e}")

# Try MirrorCopies in detail
print("\n" + "=" * 60)
print("MIRROR COPY DETAILED INVESTIGATION")
print("=" * 60)

try:
    mc = model.MirrorCopies
    print(f"  MirrorCopies.Count = {mc.Count}")

    # Check if we can get type info
    try:
        import win32com.client as wc

        mc_early = wc.gencache.EnsureDispatch(mc)
        # Get type info
        ti = mc_early._oleobj_.GetTypeInfo()
        ta = ti.GetTypeAttr()
        print(f"  Type: {ta[0]}")
        print(f"  Funcs: {ta[6]}")
        for i in range(ta[6]):
            fd = ti.GetFuncDesc(i)
            names = ti.GetNames(fd[0])
            print(f"    [{i}] {names[0]}({', '.join(names[1:])}) -> invkind={fd[3]}")
    except Exception as e2:
        print(f"  TypeInfo error: {e2}")
        # Try raw method access
        for m in ["Add", "AddMirror", "AddMirrorCopy"]:
            try:
                method = getattr(mc, m)
                print(f"  mc.{m} accessible!")
            except Exception:
                pass
except Exception as e:
    print(f"  Error: {e}")

# Try Drafts in detail
print("\n" + "=" * 60)
print("DRAFTS DETAILED INVESTIGATION")
print("=" * 60)

try:
    drafts = model.Drafts
    print(f"  Drafts.Count = {drafts.Count}")

    try:
        import win32com.client as wc

        dr_early = wc.gencache.EnsureDispatch(drafts)
        ti = dr_early._oleobj_.GetTypeInfo()
        ta = ti.GetTypeAttr()
        print(f"  Type: {ta[0]}")
        print(f"  Funcs: {ta[6]}")
        for i in range(ta[6]):
            fd = ti.GetFuncDesc(i)
            names = ti.GetNames(fd[0])
            params_str = ", ".join(names[1:]) if len(names) > 1 else ""
            print(f"    [{i}] {names[0]}({params_str}) -> invkind={fd[3]}")
    except Exception as e2:
        print(f"  TypeInfo error: {e2}")
except Exception as e:
    print(f"  Error: {e}")

print("\n" + "=" * 60)
print("DONE")
print("=" * 60)
