"""Tests for shared backend validation helpers."""

from math import inf, nan

import pytest

from solidedge_mcp.backends.validation import validate_numerics, validate_path


def test_validate_numerics_accepts_ints_and_floats() -> None:
    assert validate_numerics(x=0, y=1.25, z=-3) is None


@pytest.mark.parametrize("value", ["1", None, True, False, nan, inf, -inf])
def test_validate_numerics_rejects_invalid_values(value: object) -> None:
    result = validate_numerics(distance=value)

    assert result is not None
    assert "distance" in result["error"]


def test_validate_path_allows_existing_path(tmp_path) -> None:
    input_file = tmp_path / "part.par"
    input_file.write_text("solid edge data")

    normalized, error = validate_path(str(input_file), must_exist=True)

    assert normalized == str(input_file)
    assert error is None


def test_validate_path_rejects_missing_required_path(tmp_path) -> None:
    missing_file = tmp_path / "missing.par"

    normalized, error = validate_path(str(missing_file), must_exist=True)

    assert normalized == str(missing_file)
    assert error is not None
    assert "does not exist" in error["error"]


def test_validate_path_allows_new_file_when_parent_exists(tmp_path) -> None:
    output_file = tmp_path / "output.par"

    normalized, error = validate_path(str(output_file), must_exist=False)

    assert normalized == str(output_file)
    assert error is None


def test_validate_path_rejects_missing_output_parent(tmp_path) -> None:
    output_file = tmp_path / "missing-dir" / "output.par"

    normalized, error = validate_path(str(output_file), must_exist=False)

    assert normalized == str(output_file)
    assert error is not None
    assert "parent directory" in error["error"]


def test_validate_path_rejects_empty_path() -> None:
    normalized, error = validate_path("", must_exist=True)

    assert normalized == ""
    assert error is not None
    assert "required" in error["error"]
