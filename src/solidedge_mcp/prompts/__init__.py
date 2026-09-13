"""MCP prompt templates and the server instructions string.

Prompts are conversation starters an MCP client can offer the user. They
carry the workflow knowledge that used to live only in CLAUDE.md, so an LLM
driving this server from any client sees the same guidance.
"""

from __future__ import annotations

from typing import Any

# Shown to the client at initialize time. Keep it short: it is in context on
# every request. Detail lives in the solidedge://guide/* resources.
SERVER_INSTRUCTIONS = """\
Solid Edge CAD automation over COM. Windows only; Solid Edge must be installed.

Conventions (apply to every tool):
- Lengths are METERS (10 mm = 0.01). Angles are DEGREES.
- Reference plane indices are 1-based: 1=Top(XY), 2=Right(YZ), 3=Front(XZ).
- Face, edge, feature, and component indices are 0-based.
- Results are dicts. A failure is {"error": "...", ...}; check for the key.

Start every session with manage_connection(action="connect"), then create or
open a document. Solid features need a closed sketch: create_sketch ->
draw_* -> close_sketch -> create_extrude / create_revolve / ... The closed
profile is consumed by the next feature call.

Read model state through solidedge:// resources (feature list, mass
properties, document info) instead of tools where one exists. Read
solidedge://guide/workflows for step-by-step recipes and
solidedge://guide/conventions for the full rules.
"""

WORKFLOWS_GUIDE = """\
# Solid Edge MCP workflows

All lengths in meters, angles in degrees. Plane indices 1-based
(1=Top/XY, 2=Right/YZ, 3=Front/XZ); everything else 0-based.

## Extruded part (100 mm square, 50 mm tall)
1. manage_connection(action="connect")
2. manage_document(action="create", doc_type="part")
3. create_sketch(plane="Top")
4. draw_rectangle(x1=0, y1=0, x2=0.1, y2=0.1)
5. close_sketch()
6. create_extrude(method="finite", distance=0.05)
7. manage_document(action="save", file_path="C:/temp/box.par")
8. export_file(format="step", file_path="C:/temp/box.step")

## Revolved part
1-2 as above, then create_sketch(plane="Front")
3. draw_line x4 forming a closed profile touching the axis (x=0)
4. close_sketch()
5. create_revolve(method="finite", angle=360)

## Cut a hole through an existing body
1. create_sketch(plane="Top")
2. draw_circle(cx=0.05, cy=0.05, radius=0.005)
3. close_sketch()
4. create_extruded_cutout(method="through_all")
   Front-plane cutouts: use direction="Symmetric" to cut both ways.
   One closed profile per cutout sketch; SE rejects disjoint profiles.

## Assembly
1. manage_document(action="create", doc_type="assembly")
2. add_assembly_component(file_path="C:/parts/base.par", x=0, y=0, z=0)
3. add_assembly_component(file_path="C:/parts/top.par", x=0, y=0, z=0.1)
4. solidedge://assembly/components  (resource) to get indices
5. manage_relation(action="add_mate", ...) etc.
Assembly-level pattern/mirror are not available via COM (E_ACCESSDENIED);
pattern at part level or place components individually.

## Query
- solidedge://document/info, solidedge://features, solidedge://mass-properties
- measure(...) for distances/angles between entities
- diagnose_api(...) to inspect what a COM object exposes when a call fails
"""

CONVENTIONS_GUIDE = """\
# Solid Edge MCP conventions

## Units
- Length: meters. Convert mm by dividing by 1000.
- Angle: degrees on every tool (converted to radians internally).
- Mass properties return kg, m^3, m^2 with mm/cm variants where provided.

## Indices
- Reference planes: 1-based. 1=Top (XY, normal +Z), 2=Right (YZ, normal +X),
  3=Front (XZ, normal +Y). A plane-index parameter defaulting to 0 means
  "not provided" for methods that do not need it.
- Faces, edges, vertices, features, components, relations: 0-based.
- Sketch constraint element references use 1-based indices: [["line", 1], ["line", 2]].

## Sketch and feature lifecycle
- create_sketch opens a profile on a plane and makes it active.
- draw_* tools add 2D geometry to the active profile.
- close_sketch validates the profile (must be closed for solids) and queues
  it for the next feature. Do not call close_sketch twice.
- Solid features consume the queued profile. Loft/sweep consume several.
- After a feature call the sketch is no longer active; create a new one.

## Directions
- direction="Normal" follows the plane normal; "Reverse" the opposite;
  "Symmetric" both ways. On the Front plane the COM normal is inverted
  relative to world +Y; prefer Symmetric for cutouts there.

## Errors
- Every result is a dict. {"error": "..."} means failure. COM failures include
  "hresult" (hex) and, when Solid Edge went away, "disconnected": true.
  Reconnect with manage_connection(action="connect").
- Set SOLIDEDGE_MCP_DEBUG=1 in the server environment to include tracebacks.

## Known COM limits (Solid Edge 2025/2026)
- AssemblyFeaturesPatterns.Add and AssemblyFeaturesMirrors.Add return
  E_ACCESSDENIED; assembly-level pattern/mirror tools report unsupported.
- One closed profile per extruded-cutout sketch.
- Shell (thin-wall) requires interactive face selection; not exposed.
"""


def register_prompts(mcp: Any) -> None:
    """Register conversation-starter prompts."""

    @mcp.prompt(name="new_part", tags={"part"})
    def new_part(description: str) -> str:
        """Model a new part from a plain-language description."""
        return (
            "You are driving Solid Edge through the solidedge MCP server. "
            "First read solidedge://guide/workflows and solidedge://guide/conventions. "
            "Then model this part, one feature at a time, checking each result dict "
            "for an 'error' key before continuing. Units are meters and degrees. "
            "When done, read solidedge://mass-properties and summarise the result.\n\n"
            f"Part to model:\n{description}"
        )

    @mcp.prompt(name="design_review", tags={"part", "assembly"})
    def design_review(focus: str = "general") -> str:
        """Review the active document: feature tree, geometry sanity, mass properties."""
        return (
            "Review the active Solid Edge document. Read solidedge://document/info, "
            "solidedge://features, and solidedge://mass-properties. If it is an assembly "
            "also read solidedge://assembly/components and run "
            "query_component(property='interference'). Report: what the model is, "
            "suspicious features (failed, suppressed, zero-volume), and concrete "
            f"improvements. Focus: {focus}."
        )

    @mcp.prompt(name="manufacturability_check", tags={"part", "sheet_metal"})
    def manufacturability_check(process: str = "CNC machining") -> str:
        """Check the active part for manufacturability issues for a given process."""
        return (
            "Assess the active Solid Edge part for manufacturability by "
            f"{process}. Read solidedge://features and solidedge://mass-properties, "
            "then use query_face and query_edge on representative faces/edges to find "
            "thin walls, deep narrow pockets, sharp internal corners, undercuts, and "
            "unreachable features. List each issue with the feature/face index, why it "
            "is a problem for this process, and a suggested change."
        )

    @mcp.prompt(name="troubleshoot_feature", tags={"diagnostics"})
    def troubleshoot_feature(error_text: str) -> str:
        """Diagnose a failed feature or COM call and propose a fix."""
        return (
            "A Solid Edge MCP call failed with this error:\n\n"
            f"{error_text}\n\n"
            "Read solidedge://guide/conventions. Check the sketch state "
            "(solidedge://sketch/info): is a closed profile queued, on which plane? "
            "Use diagnose_api or diagnose_feature_tool to inspect the COM object "
            "involved. Explain the most likely cause and the exact corrected call."
        )
