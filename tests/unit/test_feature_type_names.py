"""The feature-type name table must match the type library it came from.

Feature.Type is a FeatureTypeConstants value. A caller cannot do anything
with a bare 462094706, so list_features reports "extruded protrusion" beside
it. The table was generated from the scraped dump; this checks it did not
drift, and skips when the dump is absent as the other type-library tests do.
"""

from __future__ import annotations

import json
import pathlib

import pytest

from solidedge_mcp.backends.constants import FEATURE_TYPE_NAMES

REPO_ROOT = pathlib.Path(__file__).resolve().parents[2]
DUMP = REPO_ROOT / "reference" / "typelib_dump.json"


@pytest.fixture(scope="module")
def enum_values() -> dict[str, int]:
    data = json.loads(DUMP.read_text(encoding="utf-8"))
    return data["typelibs"]["Program/constant.tlb"]["enums"]["FeatureTypeConstants"]


needs_dump = pytest.mark.skipif(not DUMP.exists(), reason="reference/typelib_dump.json missing")


class TestTableMatchesTheTypeLibrary:
    @needs_dump
    def test_every_value_is_covered(self, enum_values):
        missing = sorted(set(enum_values.values()) - set(FEATURE_TYPE_NAMES))
        assert not missing, (
            f"FeatureTypeConstants gained {len(missing)} value(s) the table does not "
            f"name: {missing[:8]}. Regenerate the table."
        )

    @needs_dump
    def test_no_invented_values(self, enum_values):
        extra = sorted(set(FEATURE_TYPE_NAMES) - set(enum_values.values()))
        assert not extra, f"table names values no enum has: {extra[:8]}"

    @needs_dump
    def test_a_known_name_maps_to_its_value(self, enum_values):
        value = enum_values["igExtrudedProtrusionFeatureObject"]
        assert FEATURE_TYPE_NAMES[value] == "extruded protrusion"


class TestNamesAreReadable:
    def test_no_name_keeps_its_prefix_or_suffix(self):
        for value, name in FEATURE_TYPE_NAMES.items():
            assert not name.startswith("ig"), f"{value}: {name}"
            assert "FeatureObject" not in name, f"{value}: {name}"
            assert name == name.lower(), f"{value}: {name}"
            assert name.strip(), f"{value} has an empty name"

    def test_the_table_is_not_empty(self):
        assert len(FEATURE_TYPE_NAMES) > 100


class TestDescribeType:
    def test_a_known_type_gets_a_name_and_keeps_the_code(self):
        from solidedge_mcp.backends.features._base import _describe_type

        assert _describe_type(462094706) == {
            "type": "extruded protrusion",
            "type_code": 462094706,
        }

    def test_an_unknown_number_is_passed_through(self):
        from solidedge_mcp.backends.features._base import _describe_type

        assert _describe_type(12345) == {"type": 12345}

    def test_something_that_is_not_a_number_reads_as_unknown(self):
        from solidedge_mcp.backends.features._base import _describe_type

        assert _describe_type(None) == {"type": "Unknown"}
        assert _describe_type("x") == {"type": "Unknown"}
