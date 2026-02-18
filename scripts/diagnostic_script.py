"""
Standalone diagnostic script to explore Solid Edge COM API
"""

import json
import sys

import win32com.client


def check_collections(doc):
    """Check which collections are available on the document"""
    collection_names = [
        "Models",
        "ExtrudedCutouts",
        "Cutouts",
        "Features",
        "ProfileSets",
        "Profiles",
        "ExtrudedProtrusions",
        "Holes",
        "Rounds",
        "Chamfers",
        "Patterns",
        "RibWebs",
        "Threads",
        "Constructions",
        "RefPlanes",
        "UserDefinedPatterns",
        "Assemblies",
        "Occurrences",
        "Sketches",
        "Shells",
        "ExtrudedProtrusions",
        "RevolvedProtrusions",
        "LoftedProtrusions",
        "SweptProtrusions",
        "Gussets",
        "Beads",
        "Dimples",
    ]

    available = {}

    for name in collection_names:
        try:
            coll = getattr(doc, name, None)
            if coll is not None:
                # Try to get count
                try:
                    count = coll.Count
                except Exception:
                    count = "N/A"

                # Get Add methods
                try:
                    methods = [
                        m for m in dir(coll) if m.startswith("Add") and not m.startswith("_")
                    ]
                    available[name] = {"exists": True, "count": count, "add_methods": methods}
                except Exception:
                    available[name] = {"exists": True, "count": count, "add_methods": []}
        except Exception:
            pass

    return available


def main():
    print("Connecting to Solid Edge...")
    try:
        app = win32com.client.Dispatch("SolidEdge.Application")
        app.Visible = True
        print(f"Connected to Solid Edge {app.Version}")
    except Exception as e:
        print(f"Failed to connect: {e}")
        return 1

    # Create a new part document
    print("\nCreating part document...")
    try:
        doc = app.Documents.Add("SolidEdge.PartDocument")
        print("Part document created")
    except Exception as e:
        print(f"Failed to create document: {e}")
        return 1

    # Check collections
    print("\nChecking available collections...")
    collections = check_collections(doc)

    print(f"\n{'=' * 70}")
    print("AVAILABLE COLLECTIONS")
    print(f"{'=' * 70}\n")

    for name, info in sorted(collections.items()):
        print(f"\n{name}:")
        print(f"  Count: {info['count']}")
        if info["add_methods"]:
            print(f"  Add Methods ({len(info['add_methods'])}):")
            for method in info["add_methods"]:
                print(f"    - {method}")
        else:
            print("  Add Methods: None found")

    # Save to file
    with open("collections_diagnostic.json", "w") as f:
        json.dump(collections, f, indent=2)
    print("\n\nFull results saved to collections_diagnostic.json")

    return 0


if __name__ == "__main__":
    sys.exit(main())
