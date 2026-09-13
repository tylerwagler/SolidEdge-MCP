"""Unit tests for the verifies_geometry decorator (honest feature status).

The decorator keys on the body's FACE COUNT (and Models.Count for the first
solid), because a failed feature still adds a feature-tree node but leaves the
body's faces unchanged.
"""

from unittest.mock import MagicMock

from solidedge_mcp.backends.features._base import FeatureManagerBase, verifies_geometry


class _Counter:
    def __init__(self, count):
        self.Count = count


class _Body:
    def __init__(self, face_count):
        self.face_count = face_count

    def Faces(self, _query):
        return _Counter(self.face_count)


class _Model:
    def __init__(self, body):
        self.Body = body


class _Models:
    def __init__(self, count, body):
        self.Count = count
        self._body = body

    def Item(self, _i):
        return _Model(self._body)


class _FakeDoc:
    def __init__(self, models_count, face_count, faces_int=True):
        body = _Body(face_count if faces_int else MagicMock())
        self.Models = _Models(models_count, body)
        self.body = body


class _Mgr(FeatureManagerBase):
    def __init__(self, doc):
        dm = MagicMock()
        dm.get_active_document.return_value = doc
        super().__init__(dm, MagicMock())
        self._doc = doc

    @verifies_geometry
    def make(self, *, new_faces):
        # Simulate the COM feature changing (or not) the body's face count.
        self._doc.body.face_count = new_faces
        return {"status": "created", "type": "test"}

    @verifies_geometry
    def make_error(self):
        return {"error": "boom"}


def test_passes_through_when_faces_change():
    mgr = _Mgr(_FakeDoc(models_count=1, face_count=6))
    result = mgr.make(new_faces=8)  # a round added faces
    assert result["status"] == "created"


def test_downgrades_to_error_when_faces_unchanged():
    mgr = _Mgr(_FakeDoc(models_count=1, face_count=6))
    result = mgr.make(new_faces=6)  # no-op feature
    assert "error" in result
    assert result["faces_before"] == 6
    assert result["faces_after"] == 6
    assert result["attempted"] == "test"


def test_base_feature_first_solid_counts_as_success():
    # Models 0 -> 1 (faces 0 -> 6): the base solid appeared.
    doc = _FakeDoc(models_count=0, face_count=0)
    mgr = _Mgr(doc)

    @verifies_geometry
    def make_base(self):
        doc.Models.Count = 1
        doc.body.face_count = 6
        return {"status": "created", "type": "base"}

    result = make_base(mgr)
    assert result["status"] == "created"


def test_existing_error_is_not_rewritten():
    mgr = _Mgr(_FakeDoc(models_count=1, face_count=6))
    assert mgr.make_error() == {"error": "boom"}


def test_conservative_when_faces_not_readable():
    mgr = _Mgr(_FakeDoc(models_count=1, face_count=6, faces_int=False))
    result = mgr.make(new_faces=6)
    assert result["status"] == "created"  # cannot prove failure -> pass through


def test_no_geometry_created_logic():
    f = FeatureManagerBase._no_geometry_created
    assert f((1, 6), (1, 6)) is True      # body unchanged
    assert f((1, 6), (1, 8)) is False     # faces changed
    assert f((0, 0), (1, 6)) is False     # base solid appeared
    assert f((None, None), (None, None)) is False  # unknown -> no claim
