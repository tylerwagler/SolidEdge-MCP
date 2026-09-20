"""Unit tests for the shared input-validation helpers."""

from __future__ import annotations

# ============================================================================
# OVERWRITE GUARD
# ============================================================================


class TestGuardOverwrite:
    """Every Solid Edge call that writes a path must go through this.

    Solid Edge answers an existing target with a modal "This file exists. Do
    you want to overwrite it?" prompt. DisplayAlerts does not suppress it, and
    Solid Edge has a single UI thread, so the COM call never returns and the
    whole server stops answering. Reproduced and photographed on Solid Edge
    2026; before the guard, an export over an existing PDF hung the session.
    """

    def test_a_free_path_may_be_written(self, tmp_path):
        from solidedge_mcp.backends.validation import guard_overwrite

        assert guard_overwrite(str(tmp_path / "new.step"), overwrite=False) is None

    def test_an_existing_file_is_refused_by_default(self, tmp_path):
        from solidedge_mcp.backends.validation import guard_overwrite

        target = tmp_path / "taken.step"
        target.write_text("x", encoding="utf-8")

        err = guard_overwrite(str(target), overwrite=False)

        assert err is not None
        assert "already exists" in err["error"]
        assert "overwrite=true" in err["error"]
        assert err["exists"] is True
        # Nothing is removed when the write is refused.
        assert target.exists()

    def test_overwrite_clears_the_way(self, tmp_path):
        from solidedge_mcp.backends.validation import guard_overwrite

        target = tmp_path / "taken.step"
        target.write_text("x", encoding="utf-8")

        assert guard_overwrite(str(target), overwrite=True) is None
        # Removed up front, so Solid Edge is never asked about it.
        assert not target.exists()

    def test_a_file_that_cannot_be_removed_is_reported(self, tmp_path, monkeypatch):
        from solidedge_mcp.backends import validation

        target = tmp_path / "locked.step"
        target.write_text("x", encoding="utf-8")

        def refuse(self):
            raise OSError("in use by another process")

        monkeypatch.setattr(validation.Path, "unlink", refuse)

        err = validation.guard_overwrite(str(target), overwrite=True)

        assert err is not None
        assert "in use by another process" in err["error"]

    def test_a_directory_in_the_way_is_refused(self, tmp_path):
        from solidedge_mcp.backends.validation import guard_overwrite

        target = tmp_path / "adirectory"
        target.mkdir()

        err = guard_overwrite(str(target), overwrite=False)

        assert err is not None
        assert target.exists()
