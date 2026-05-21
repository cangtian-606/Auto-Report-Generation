"""DataMapper unit tests"""

import sys
from pathlib import Path

import pandas as pd
import pytest

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))

from src.mapper import DataMapper


class TestConvertValue:
    def test_zero_int_kept_as_zero(self):
        mapper = DataMapper({})
        result = mapper._convert_value(0)
        assert result == 0

    def test_zero_float_kept_as_zero(self):
        mapper = DataMapper({})
        result = mapper._convert_value(0.0)
        assert result == 0.0

    def test_positive_int_kept(self):
        mapper = DataMapper({})
        result = mapper._convert_value(100)
        assert result == 100

    def test_negative_int_kept(self):
        mapper = DataMapper({})
        result = mapper._convert_value(-50)
        assert result == -50

    def test_true_string_returns_bool(self):
        mapper = DataMapper({})
        result = mapper._convert_value("TRUE")
        assert result is True

    def test_false_string_returns_bool(self):
        mapper = DataMapper({})
        result = mapper._convert_value("FALSE")
        assert result is False

    def test_none_returns_empty_string(self):
        mapper = DataMapper({})
        result = mapper._convert_value(None)
        assert result == ""

    def test_whitespace_true_returns_bool(self):
        mapper = DataMapper({})
        result = mapper._convert_value("  TRUE  ")
        assert result is True

    def test_ordinal_string_kept(self):
        mapper = DataMapper({})
        result = mapper._convert_value("hello")
        assert result == "hello"


class TestBuildContext:
    def test_date_sheet_builds_nested_context(self):
        raw_data = {"date.全局信息": {"公司名": "Test公司"}}
        mapper = DataMapper(raw_data)
        ctx = mapper.build_context()
        assert ctx["date"]["全局信息"]["公司名"] == "Test公司"

    def test_form_sheet_builds_table_context(self):
        df = pd.DataFrame({"名称": ["A", "B"], "金额": [100, 200]})
        raw_data = {"form.商品列表": df}
        mapper = DataMapper(raw_data)
        ctx = mapper.build_context()
        assert len(ctx["form"]["商品列表"]) == 2
        assert ctx["form"]["商品列表"][0]["名称"] == "A"

    def test_non_prefixed_sheet_goes_to_date(self):
        raw_data = {"随便Sheet": {"key1": "val1"}}
        mapper = DataMapper(raw_data)
        ctx = mapper.build_context()
        assert ctx["date"]["key1"] == "val1"

    def test_non_prefixed_dataframe_goes_to_form(self):
        df = pd.DataFrame({"col": [1, 2]})
        raw_data = {"任意表格": df}
        mapper = DataMapper(raw_data)
        ctx = mapper.build_context()
        assert ctx["form"]["任意表格"] == [{"col": 1}, {"col": 2}]


class TestSetNested:
    def test_two_level_nesting(self):
        mapper = DataMapper({})
        d = {}
        mapper._set_nested(d, "a.b", "val")
        assert d["a"]["b"] == "val"

    def test_three_level_nesting(self):
        mapper = DataMapper({})
        d = {}
        mapper._set_nested(d, "x.y.z", 42)
        assert d["x"]["y"]["z"] == 42
