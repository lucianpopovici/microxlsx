"""
Unit tests for microxlsx.utils
"""
import pytest
from microxlsx.utils import cell_to_indices, indices_to_cell


class TestCellToIndices:
    def test_a1(self):
        assert cell_to_indices("A1") == (0, 0)

    def test_b2(self):
        assert cell_to_indices("B2") == (1, 1)

    def test_z1(self):
        assert cell_to_indices("Z1") == (0, 25)

    def test_aa1(self):
        assert cell_to_indices("AA1") == (0, 26)

    def test_ab10(self):
        assert cell_to_indices("AB10") == (9, 27)

    def test_lowercase(self):
        assert cell_to_indices("a1") == (0, 0)

    def test_large_row(self):
        assert cell_to_indices("A100") == (99, 0)

    def test_invalid_raises_value_error(self):
        with pytest.raises(ValueError, match="Invalid cell reference"):
            cell_to_indices("invalid")

    def test_no_column_raises_value_error(self):
        with pytest.raises(ValueError):
            cell_to_indices("123")

    def test_no_row_raises_value_error(self):
        with pytest.raises(ValueError):
            cell_to_indices("ABC")


class TestIndicesToCell:
    def test_0_0(self):
        assert indices_to_cell(0, 0) == "A1"

    def test_1_1(self):
        assert indices_to_cell(1, 1) == "B2"

    def test_0_25(self):
        assert indices_to_cell(0, 25) == "Z1"

    def test_0_26(self):
        assert indices_to_cell(0, 26) == "AA1"

    def test_0_27(self):
        assert indices_to_cell(0, 27) == "AB1"

    def test_large_row(self):
        assert indices_to_cell(99, 0) == "A100"


class TestRoundtrip:
    @pytest.mark.parametrize("ref", ["A1", "B2", "Z1", "AA1", "AB10", "AZ99", "ZZ100"])
    def test_cell_to_indices_and_back(self, ref):
        assert indices_to_cell(*cell_to_indices(ref)) == ref
