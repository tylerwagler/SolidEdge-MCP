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


@pytest.fixture(scope="module")
def managers(solid_edge):
    """Document, sketch, and feature managers bound to the live application."""
    from solidedge_mcp.backends.documents import DocumentManager
    from solidedge_mcp.backends.features import FeatureManager
    from solidedge_mcp.backends.sketching import SketchManager

    doc_mgr = DocumentManager(solid_edge)
    sketch_mgr = SketchManager(doc_mgr)
    doc_mgr.sketch_manager = sketch_mgr
    feature_mgr = FeatureManager(doc_mgr, sketch_mgr)
    return doc_mgr, sketch_mgr, feature_mgr


@pytest.fixture
def new_part(managers, solid_edge, preexisting_documents):
    """Create a scratch part; close it afterwards, and only it.

    The teardown closes the document only when the manager still points at the
    document this fixture created and that document was not already open when
    the session began. Anything else is left untouched.
    """
    doc_mgr, _, _ = managers

    before = _open_document_names(solid_edge.get_application())
    result = doc_mgr.create_part()
    assert "error" not in result, f"Failed to create part: {result}"

    created_name = result.get("name")
    assert created_name, f"create_part returned no name: {result}"
    assert created_name not in before, (
        f"create_part reported {created_name!r}, which was already open. "
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
