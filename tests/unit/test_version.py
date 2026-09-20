"""The version is published in three places; a release with two of them stale is a
release nobody can pin. ``pyproject.toml`` is what the package reports, ``__version__``
is what code reads, and the top ``CHANGELOG.md`` entry is what people read."""

from __future__ import annotations

import pathlib
import re
import tomllib

import solidedge_mcp

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent


def _pyproject_version() -> str:
    data = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    return str(data["project"]["version"])


def test_init_matches_pyproject():
    assert solidedge_mcp.__version__ == _pyproject_version()


def test_changelog_leads_with_the_current_version():
    text = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    match = re.search(r"^## (\S+)", text, re.MULTILINE)
    assert match, "CHANGELOG.md has no '## <version>' entry"
    assert match.group(1) == _pyproject_version()


def test_version_is_pep440():
    assert re.fullmatch(r"\d+\.\d+\.\d+(?:(?:a|b|rc)\d+)?", _pyproject_version())
