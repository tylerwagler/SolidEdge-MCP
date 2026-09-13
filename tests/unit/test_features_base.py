"""Unit tests for the verifies_geometry decorator (honest feature status).

The decorator keys on the TOTAL FACE COUNT across all bodies in Models (and
Models.Count for the first solid), because a failed feature still adds a
feature-tree node but leaves the bodies' faces unchanged.
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
    """Fake Models collection holding one body per model, 1-indexed like COM."""

    def __init__(self, count, bodies):
        self.Count = count
        self._bodies = list(bodies)
        self.item_calls = []

    def Item(self, i):
        self.item_calls.append(i)
        if i < 1 or i > len(self._bodies):
            raise IndexError(f"Models.Item({i}) out of range (have {len(self._bodies)})")
        return _Model(self._bodies[i - 1])


class _FakeDoc:
    """Document with ``models_count`` bodies; ``face_count`` is body 1's faces,
    ``extra_face_counts`` gives the faces of bodies 2..N."""

    def __init__(self, models_count, face_count, faces_int=True, extra_face_counts=()):
        body = _Body(face_count if faces_int else MagicMock())
        bodies = [body] + [_Body(fc) for fc in extra_face_counts]
        self.Models = _Models(models_count, bodies)
        self.body = body
        self.bodies = bodies


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
    assert f((1, 6), (1, 6)) is True  # body unchanged
    assert f((1, 6), (1, 8)) is False  # faces changed
    assert f((0, 0), (1, 6)) is False  # base solid appeared
    assert f((None, None), (None, None)) is False  # unknown -> no claim


# ---------------------------------------------------------------------------
# Multi-body: the snapshot must SUM face counts over Models.Item(1..Count)
# ---------------------------------------------------------------------------


def test_snapshot_sums_faces_across_all_bodies():
    doc = _FakeDoc(models_count=3, face_count=6, extra_face_counts=(10, 4))
    mgr = _Mgr(doc)
    assert mgr._geometry_snapshot() == (3, 20)
    # Every body was visited, 1-indexed, exactly once.
    assert doc.Models.item_calls == [1, 2, 3]


def test_single_body_snapshot_unchanged():
    mgr = _Mgr(_FakeDoc(models_count=1, face_count=6))
    assert mgr._geometry_snapshot() == (1, 6)


def test_change_on_second_body_counts_as_geometry_created():
    # A multi-body cutout that only touches body 2 (body 1 untouched) must NOT
    # be rewritten as "no geometry was created".
    doc = _FakeDoc(models_count=2, face_count=6, extra_face_counts=(6,))
    mgr = _Mgr(doc)

    @verifies_geometry
    def cut_second_body(self):
        doc.bodies[1].face_count = 9  # body 1 stays at 6 faces
        return {"status": "created", "type": "extruded_cutout_multi_body"}

    result = cut_second_body(mgr)
    assert result["status"] == "created"
    assert "error" not in result


def test_change_on_last_of_many_bodies_counts_as_geometry_created():
    doc = _FakeDoc(models_count=4, face_count=6, extra_face_counts=(6, 6, 6))
    mgr = _Mgr(doc)

    @verifies_geometry
    def cut_last_body(self):
        doc.bodies[3].face_count = 7
        return {"status": "created", "type": "hole_multi_body"}

    assert cut_last_body(mgr)["status"] == "created"


def test_multi_body_no_op_is_still_downgraded():
    # No body changed -> the honest-status downgrade still fires, reporting
    # the summed totals.
    doc = _FakeDoc(models_count=2, face_count=6, extra_face_counts=(10,))
    mgr = _Mgr(doc)
    result = mgr.make(new_faces=6)
    assert "error" in result
    assert result["faces_before"] == 16
    assert result["faces_after"] == 16
    assert result["models_before"] == 2


def test_multi_body_offsetting_changes_are_not_masked_into_error_by_mistake():
    # Sanity: a change to body 1 alone is still recognised in a multi-body doc.
    doc = _FakeDoc(models_count=2, face_count=6, extra_face_counts=(10,))
    mgr = _Mgr(doc)
    assert mgr.make(new_faces=8)["status"] == "created"


def test_conservative_when_any_body_unreadable():
    # Body 2's face count is a MagicMock (not an int): the whole total must be
    # None, never a partial sum from body 1 alone.
    doc = _FakeDoc(models_count=2, face_count=6, extra_face_counts=(MagicMock(),))
    mgr = _Mgr(doc)
    assert mgr._geometry_snapshot() == (2, None)
    assert mgr.make(new_faces=6)["status"] == "created"  # cannot prove -> pass through


def test_conservative_when_a_body_raises():
    # Models.Count claims 3 bodies but Item(3) raises: total is None, not 16.
    doc = _FakeDoc(models_count=3, face_count=6, extra_face_counts=(10,))
    mgr = _Mgr(doc)
    assert mgr._geometry_snapshot() == (3, None)
    assert mgr.make(new_faces=6)["status"] == "created"


def test_unittest_mock_document_bails_out():
    # A MagicMock document is never measured, regardless of what it reports.
    doc = MagicMock()
    doc.Models.Count = 1
    doc.Models.Item.return_value.Body.Faces.return_value.Count = 6
    mgr = _Mgr(doc)
    assert mgr._geometry_snapshot() == (None, None)
    doc.Models.Item.assert_not_called()
