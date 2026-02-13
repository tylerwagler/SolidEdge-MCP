"""Sheet metal tools for Solid Edge MCP.

Note: Most sheet metal features (dimple, bead, louver, gusset, etch, flange,
base_flange, base_tab, etc.) are registered in features.py since they share
the FeatureManager backend. This module is reserved for sheet-metal-only
operations that don't fit in features.py (currently none beyond what's there).
"""


def register(mcp):
    """Register sheet metal tools with the MCP server."""
    # All sheet metal tools are registered via features.py and documents.py.
    # This module exists as an organizational placeholder.
    pass
