"""Jinja2 filter unit tests"""

import sys
from datetime import datetime
from pathlib import Path

import pytest

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))

from src.render.filters import (
    filter_money,
    filter_percent,
    filter_num,
    filter_date,
    filter_default_dash,
    filter_default,
    filter_int,
)


class TestFilterMoney:
    def test_none_returns_empty(self):
        assert filter_money(None) == ''

    def test_empty_string_returns_empty(self):
        assert filter_money('') == ''

    def test_zero_returns_empty(self):
        assert filter_money(0) == ''

    def test_positive_number_formats(self):
        assert filter_money(1234567.89) == '1,234,567.89'

    def test_invalid_returns_original_string(self):
        assert filter_money('abc') == 'abc'


class TestFilterPercent:
    def test_none_returns_empty(self):
        assert filter_percent(None) == ''

    def test_zero_returns_empty(self):
        assert filter_percent(0) == ''

    def test_decimal_formats(self):
        assert filter_percent(0.25) == '25.00%'

    def test_whole_number_formats(self):
        assert filter_percent(1) == '100.00%'


class TestFilterNum:
    def test_none_returns_empty(self):
        assert filter_num(None) == ''

    def test_zero_returns_empty(self):
        assert filter_num(0) == ''

    def test_integer_formats(self):
        assert filter_num(1234567) == '1,234,567'

    def test_float_formats(self):
        assert filter_num(1234567.89) == '1,234,567.89'


class TestFilterDate:
    def test_none_returns_empty(self):
        assert filter_date(None) == ''

    def test_iso_date_formats(self):
        assert filter_date('2025-01-15') == '2025年01月15日'

    def test_slash_date_formats(self):
        assert filter_date('2025/01/15') == '2025年01月15日'

    def test_chinese_date_formats(self):
        assert filter_date('2025年01月15日') == '2025年01月15日'

    def test_custom_format(self):
        assert filter_date('2025-01-15', '%Y/%m/%d') == '2025/01/15'

    def test_datetime_object(self):
        assert filter_date(datetime(2025, 1, 15)) == '2025年01月15日'

    def test_unrecognized_returns_original(self):
        assert filter_date('not-a-date') == 'not-a-date'


class TestFilterDefaultDash:
    def test_none_returns_empty(self):
        assert filter_default_dash(None) == ''

    def test_empty_string_returns_empty(self):
        assert filter_default_dash('') == ''

    def test_zero_kept_as_string(self):
        assert filter_default_dash(0) == '0'

    def test_normal_value_returns_string(self):
        assert filter_default_dash('hello') == 'hello'


class TestFilterDefault:
    def test_none_returns_custom_default(self):
        assert filter_default(None, '-') == '-'

    def test_empty_string_returns_custom_default(self):
        assert filter_default('', '无') == '无'

    def test_no_custom_default_returns_empty(self):
        assert filter_default(None) == ''

    def test_normal_value_returns_string(self):
        assert filter_default('hello') == 'hello'

    def test_zero_kept_as_string(self):
        assert filter_default(0) == '0'


class TestFilterInt:
    def test_none_returns_empty(self):
        assert filter_int(None) == ''

    def test_float_truncates(self):
        assert filter_int(123.45) == '123'

    def test_integer_kept(self):
        assert filter_int(100) == '100'
