"""Shared fixtures and safety guards for integration tests.

These tests drive a real, running Solid Edge. That session may already hold
documents belonging to whoever is at the keyboard, so the fixtures here refuse
to close anything they did not create, and the session guard reports loudly if
a pre-existing document disappears.

Run with:  uv run pytest -m integration
"""

from __future__ import annotations

import contextlib
import warnings
from dataclasses import dataclass
from typing import Any

import pytest

pytestmark = pytest.mark.integration


def _open_document_names(app: Any) -> set[str]:
    """Names of every document currently open in the application."""
    names: set[str] = set()
    try:
        documents = app.Documents
        for i in range(1, documents.Count + 1):
            try:
                names.add(str(documents.Item(i).Name))
            except Exception:  # a document mid-close can raise
                continue
    except Exception:
        pass
    return names


@pytest.fixture(scope="session")
def solid_edge():
    """Connect to a running Solid Edge, or skip. Never starts one.

    Starting Solid Edge from a test run is slow and leaves an application
    behind, so integration tests attach to an existing session only.
    """
    from solidedge_mcp.backends.connection import SolidEdgeConnection

    conn = SolidEdgeConnection()
    result = conn.connect(start_if_needed=False)
    if result.get("status") != "connected":
        pytest.skip(f"Solid Edge is not running: {result.get('error', result)}")

    app = conn.get_application()
    preexisting = _open_document_names(app)
    display_alerts = None
    with contextlib.suppress(Exception):
        display_alerts = app.DisplayAlerts

    yield conn

    # close_document(save=False) disables alerts; put the user's setting back.
    if display_alerts is not None:
        with contextlib.suppress(Exception):
            app.DisplayAlerts = display_alerts

    lost = preexisting - _open_document_names(app)
    if lost:
        warnings.warn(
            f"Integration run closed pre-existing documents: {sorted(lost)}. "
            "This is a bug in a fixture or test, not in Solid Edge.",
            stacklevel=1,
        )


@pytest.fixture(scope="session")
def preexisting_documents(solid_edge) -> frozenset[str]:
    """Documents that were already open before the test session started."""
    return frozenset(_open_document_names(solid_edge.get_application()))


@dataclass(frozen=True)
class Stack:
    """Every manager the production server builds, wired the same way.

    ``managers`` keeps returning its three-tuple for the tests written against
    it; new tests take ``stack`` and get the assembly, query, export and view
    managers too.
    """

    connection: Any
    doc: Any
    sketch: Any
    feature: Any
    assembly: Any
    query: Any
    export: Any
    view: Any


@pytest.fixture(scope="module")
def stack(solid_edge) -> Stack:
    """The full manager stack, built in the order managers.py builds it."""
    from solidedge_mcp.backends.assembly import AssemblyManager
    from solidedge_mcp.backends.documents import DocumentManager
    from solidedge_mcp.backends.export import ExportManager, ViewModel
    from solidedge_mcp.backends.features import FeatureManager
    from solidedge_mcp.backends.query import QueryManager
    from solidedge_mcp.backends.sketching import SketchManager

    doc_mgr = DocumentManager(solid_edge)
    sketch_mgr = SketchManager(doc_mgr)
    doc_mgr.sketch_manager = sketch_mgr
    return Stack(
        connection=solid_edge,
        doc=doc_mgr,
        sketch=sketch_mgr,
        feature=FeatureManager(doc_mgr, sketch_mgr),
        assembly=AssemblyManager(doc_mgr, sketch_mgr),
        query=QueryManager(doc_mgr),
        export=ExportManager(doc_mgr),
        view=ViewModel(doc_mgr),
    )


@pytest.fixture(scope="module")
def managers(stack):
    """Document, sketch, and feature managers bound to the live application."""
    return stack.doc, stack.sketch, stack.feature


@contextlib.contextmanager
def _scratch_document(doc_mgr, solid_edge, preexisting_documents, create, kind):
    """Create a scratch document; close it afterwards, and only it.

    Three guards, and every document fixture has to keep all three or it can
    close a document belonging to whoever is at the keyboard:

    1. The name the creator reports must not already be open, so the name the
       teardown matches on identifies this fixture's document and nothing else.
    2. At teardown the manager must still point at that same document. A proxy
       that is gone is left alone.
    3. Anything else is refused with a warning, never closed.
    """
    before = _open_document_names(solid_edge.get_application())
    result = create()
    assert "error" not in result, f"Failed to create {kind}: {result}"

    created_name = result.get("name")
    assert created_name, f"{kind} creator returned no name: {result}"
    assert created_name not in before, (
        f"{kind} creator reported {created_name!r}, which was already open. "
        "Refusing to run: teardown could close someone else's document."
    )

    yield created_name

    active = doc_mgr.active_document
    if active is None:
        return
    try:
        active_name = str(active.Name)
    except Exception:
        # The proxy is gone; nothing safe to do.
        doc_mgr.active_document = None
        return

    if active_name != created_name or active_name in preexisting_documents:
        warnings.warn(
            f"Not closing {active_name!r}: it is not the document this test created "
            f"({created_name!r}). Close it by hand if it is scratch.",
            stacklevel=1,
        )
        doc_mgr.active_document = None
        return

    doc_mgr.close_document(save=False)


@pytest.fixture
def new_part(stack, solid_edge, preexisting_documents):
    """A scratch part, closed afterwards and only it."""
    with _scratch_document(
        stack.doc, solid_edge, preexisting_documents, stack.doc.create_part, "part"
    ) as name:
        yield name


@pytest.fixture
def new_assembly(stack, solid_edge, preexisting_documents):
    """A scratch assembly, closed afterwards and only it."""
    with _scratch_document(
        stack.doc, solid_edge, preexisting_documents, stack.doc.create_assembly, "assembly"
    ) as name:
        yield name


@pytest.fixture
def new_sheet_metal(stack, solid_edge, preexisting_documents):
    """A scratch sheet-metal part, closed afterwards and only it."""
    with _scratch_document(
        stack.doc, solid_edge, preexisting_documents, stack.doc.create_sheet_metal, "sheet metal"
    ) as name:
        yield name


@pytest.fixture
def new_draft(stack, solid_edge, preexisting_documents):
    """A scratch draft, closed afterwards and only it."""
    with _scratch_document(
        stack.doc, solid_edge, preexisting_documents, stack.doc.create_draft, "draft"
    ) as name:
        yield name


@pytest.fixture(scope="module")
def saved_part(stack, solid_edge, preexisting_documents, tmp_path_factory):
    """A 0.08 x 0.048 x 0.03 m box saved to disk, for assembly and draft tests.

    add_assembly_component and create_drawing take a file path, so a part has
    to exist on disk. The path is fresh under pytest's temp tree, so no
    overwrite prompt can arise -- Solid Edge answers an existing path with a
    modal that hangs the whole server. The scratch part is closed once saved;
    the file itself is pytest's to clean up.
    """
    path = tmp_path_factory.mktemp("parts") / "box.par"
    assert not path.exists()
    with _scratch_document(
        stack.doc, solid_edge, preexisting_documents, stack.doc.create_part, "part"
    ):
        stack.sketch.create_sketch("Top")
        stack.sketch.draw_rectangle(0, 0, 0.08, 0.048)
        stack.sketch.close_sketch()
        built = stack.feature.create_extrude(0.03)
        assert built.get("status") == "created", built
        saved = stack.doc.save_document(str(path))
        assert "error" not in saved, saved
        # SaveAs renames the document to the file name, so the factory's
        # identity check would rightly refuse to close it at teardown and the
        # scratch part would stay open for the whole module. It is closed
        # here, by the fixture that renamed it, once the rename is confirmed.
        renamed = str(stack.doc.get_active_document().Name)
        assert renamed == path.name, f"saved as {path.name!r} but the document is {renamed!r}"
        closed = stack.doc.close_document(save=False)
        assert "error" not in closed, closed
    assert path.exists(), f"save_document reported success but {path} is not on disk"
    assert path.name not in _open_document_names(solid_edge.get_application()), (
        f"{path.name} is still open after the saved-part fixture closed it"
    )
    return path
