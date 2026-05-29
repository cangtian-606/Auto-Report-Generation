"""DataMapper unit tests"""

import sys
from pathlib import Path

import pandas as pd
import pytest

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))

from src.processing.mapper import DataMapper
from src.exceptions import DataReadError


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
    def test_dict_sheet_builds_context(self):
        raw_data = {"全局信息": {"公司名": "Test公司"}}
        mapper = DataMapper(raw_data)
        ctx = mapper.build_context()
        assert ctx["全局信息"]["公司名"] == "Test公司"

    def test_df_sheet_builds_table_context(self):
        df = pd.DataFrame({"名称": ["A", "B"], "金额": [100, 200]})
        raw_data = {"商品列表": df}
        mapper = DataMapper(raw_data)
        ctx = mapper.build_context()
        assert len(ctx["商品列表"]) == 2
        assert ctx["商品列表"][0]["名称"] == "A"

    def test_non_prefixed_dict_keeps_name_as_key(self):
        raw_data = {"随便Sheet": {"key1": "val1"}}
        mapper = DataMapper(raw_data)
        ctx = mapper.build_context()
        assert ctx["随便Sheet"]["key1"] == "val1"

    def test_non_prefixed_dataframe_keeps_name_as_key(self):
        df = pd.DataFrame({"col": [1, 2]})
        raw_data = {"任意表格": df}
        mapper = DataMapper(raw_data)
        ctx = mapper.build_context()
        assert ctx["任意表格"] == [{"col": 1}, {"col": 2}]


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


class TestNestingDateToDate:
    def test_child_kv_sheet_auto_mounts(self):
        raw_data = {
            "公司": {"公司名": "Test公司"},
            "公司.地址": {"省": "重庆", "市": "璧山区"},
        }
        mapper = DataMapper(raw_data)
        ctx = mapper.build_context()
        assert ctx["公司"]["公司名"] == "Test公司"
        assert ctx["公司"]["地址"]["省"] == "重庆"
        assert ctx["公司"]["地址"]["市"] == "璧山区"

    def test_two_level_date_nesting(self):
        raw_data = {
            "a": {"k": "v1"},
            "a.b": {"k2": "v2"},
            "a.b.c": {"k3": "v3"},
        }
        mapper = DataMapper(raw_data)
        ctx = mapper.build_context()
        assert ctx["a"]["k"] == "v1"
        assert ctx["a"]["b"]["k2"] == "v2"
        assert ctx["a"]["b"]["c"]["k3"] == "v3"


class TestNestingDateToForm:
    def test_child_table_sheet_auto_mounts(self):
        df = pd.DataFrame({"股东": ["张伟"], "认缴额": [600]})
        raw_data = {
            "公司": {"公司名": "Test公司"},
            "公司.股东出资": df,
        }
        mapper = DataMapper(raw_data)
        ctx = mapper.build_context()
        assert ctx["公司"]["公司名"] == "Test公司"
        assert len(ctx["公司"]["股东出资"]) == 1
        assert ctx["公司"]["股东出资"][0]["股东"] == "张伟"


class TestNestingFormToForm:
    def test_child_table_matches_by_parent_column(self):
        parent_df = pd.DataFrame({"公司简称": ["重庆晟和泰", "安徽富军"], "信用代码": ["91", "92"]})
        child_df = pd.DataFrame({
            "_parent_公司简称": ["重庆晟和泰", "重庆晟和泰", "安徽富军"],
            "股东": ["张伟", "李芳", "王强"],
            "认缴额": [600, 250, 500],
        })
        raw_data = {
            "项目公司": parent_df,
            "项目公司.股东出资": child_df,
        }
        mapper = DataMapper(raw_data)
        ctx = mapper.build_context()
        companies = ctx["项目公司"]
        assert len(companies) == 2
        cq = companies[0]
        assert cq["公司简称"] == "重庆晟和泰"
        assert len(cq["股东出资"]) == 2
        assert cq["股东出资"][0]["股东"] == "张伟"
        assert cq["股东出资"][1]["股东"] == "李芳"
        ah = companies[1]
        assert ah["公司简称"] == "安徽富军"
        assert len(ah["股东出资"]) == 1
        assert ah["股东出资"][0]["股东"] == "王强"

    def test_child_orphan_raises_error(self):
        parent_df = pd.DataFrame({"公司简称": ["重庆晟和泰"]})
        child_df = pd.DataFrame({
            "_parent_公司简称": ["不存在的公司"],
            "股东": ["某人"],
        })
        raw_data = {
            "项目公司": parent_df,
            "项目公司.股东出资": child_df,
        }
        mapper = DataMapper(raw_data)
        with pytest.raises(DataReadError):
            mapper.build_context()

    def test_parent_row_without_child_warns_and_uses_empty(self):
        parent_df = pd.DataFrame({"公司简称": ["重庆晟和泰", "安徽富军"]})
        child_df = pd.DataFrame({
            "_parent_公司简称": ["重庆晟和泰"],
            "股东": ["张伟"],
        })
        raw_data = {
            "项目公司": parent_df,
            "项目公司.股东出资": child_df,
        }
        mapper = DataMapper(raw_data)
        context = mapper.build_context()
        rows = context["项目公司"]
        assert len(rows) == 2
        assert rows[0]["股东出资"] == [{"股东": "张伟"}]
        assert rows[1]["股东出资"] == []
        assert rows[1]["公司简称"] == "安徽富军"

    def test_multi_level_nesting_form(self):
        parent_df = pd.DataFrame({"公司简称": ["重庆晟和泰"]})
        child_df = pd.DataFrame({
            "_parent_公司简称": ["重庆晟和泰", "重庆晟和泰"],
            "股东": ["张伟", "李芳"],
        })
        grandchild_df = pd.DataFrame({
            "_parent_股东": ["张伟", "张伟", "李芳"],
            "出资方式": ["货币", "实物", "货币"],
            "金额": [300, 300, 250],
        })
        raw_data = {
            "项目公司": parent_df,
            "项目公司.股东出资": child_df,
            "项目公司.股东出资.出资明细": grandchild_df,
        }
        mapper = DataMapper(raw_data)
        ctx = mapper.build_context()
        cq = ctx["项目公司"][0]
        assert cq["公司简称"] == "重庆晟和泰"
        zw = cq["股东出资"][0]
        assert zw["股东"] == "张伟"
        assert len(zw["出资明细"]) == 2
        assert zw["出资明细"][0]["出资方式"] == "货币"
        assert zw["出资明细"][1]["出资方式"] == "实物"
        lf = cq["股东出资"][1]
        assert lf["股东"] == "李芳"
        assert len(lf["出资明细"]) == 1
        assert lf["出资明细"][0]["出资方式"] == "货币"


class TestListConvertValue:
    def test_list_items_convert_none_to_empty(self):
        raw_data = {"列表": [{"字段": None}, {"字段": "val"}]}
        mapper = DataMapper(raw_data)
        ctx = mapper.build_context()
        assert ctx["列表"][0]["字段"] == ""
        assert ctx["列表"][1]["字段"] == "val"

    def test_list_items_convert_true_string_to_bool(self):
        raw_data = {"列表": [{"标志": "TRUE"}, {"标志": "FALSE"}]}
        mapper = DataMapper(raw_data)
        ctx = mapper.build_context()
        assert ctx["列表"][0]["标志"] is True
        assert ctx["列表"][1]["标志"] is False

    def test_list_items_keep_numbers(self):
        raw_data = {"列表": [{"金额": 100}, {"金额": 0.0}]}
        mapper = DataMapper(raw_data)
        ctx = mapper.build_context()
        assert ctx["列表"][0]["金额"] == 100
        assert ctx["列表"][1]["金额"] == 0.0
