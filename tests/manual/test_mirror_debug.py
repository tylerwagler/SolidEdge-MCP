"""Debug: understand how features appear in the tree after box + extrude."""

import sys

sys.path.insert(0, "src")
import win32com.client as win32

app = win32.Dispatch("SolidEdge.Application")
doc = app.ActiveDocument
models = doc.Models

print(f"Models.Count = {models.Count}")
for i in range(1, models.Count + 1):
    m = models.Item(i)
    print(f"  Model[{i}]: Name={m.Name if hasattr(m, 'Name') else '?'}")

    # Check features within this model
    if hasattr(m, "Features"):
        feats = m.Features
        print(f"    Features.Count = {feats.Count}")
        for j in range(1, min(feats.Count + 1, 20)):
            f = feats.Item(j)
            print(f"      [{j}] {f.Name}")

    # Check specific collections
    for coll_name in [
        "ExtrudedProtrusions",
        "ExtrudedCutouts",
        "Rounds",
        "Chamfers",
        "Holes",
        "MirrorCopies",
    ]:
        try:
            coll = getattr(m, coll_name)
            if coll.Count > 0:
                print(f"    {coll_name}.Count = {coll.Count}")
        except Exception:
            pass

print(f"\nDesignEdgebarFeatures.Count = {doc.DesignEdgebarFeatures.Count}")
for i in range(1, doc.DesignEdgebarFeatures.Count + 1):
    f = doc.DesignEdgebarFeatures.Item(i)
    print(f"  [{i}] {f.Name}")

# Also check if there are multiple bodies
if models.Count > 0:
    model = models.Item(1)
    body = model.Body
    print(f"\nBody: Volume = {body.Volume}")
