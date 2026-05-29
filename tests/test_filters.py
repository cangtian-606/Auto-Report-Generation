"""Jinja2 filter unit tests"""

import sys
from datetime import datetime
from pathlib import Path

import pytest

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))

from src.render.filters import (
    filter_percent,
    filter_num,
    filter_date,
    filter_default_dash,
    filter_default,
    filter_int,
)


class TestFilterNum:
    def test_none_returns_empty(self):
        assert filter_num(None) == ''

    def test_empty_string_returns_empty(self):
        assert filter_num('') == ''

    def test_zero_returns_empty(self):
        assert filter_num(0) == ''

    def test_default_two_decimals(self):
        assert filter_num(1234567.89) == '1,234,567.89'

    def test_integer_with_default_decimals(self):
        assert filter_num(1234567) == '1,234,567.00'

    def test_zero_decimals(self):
        assert filter_num(1234567.89, 0) == '1,234,568'

    def test_four_decimals(self):
        assert filter_num(12345.6789, 4) == '12,345.6789'

    def test_negative_decimals_rounds(self):
        assert filter_num(123456, -2) == '123,500'

    def test_negative_decimals_1(self):
        assert filter_num(123456, -1) == '123,460'

    def test_unit_wan(self):
        assert filter_num(12345678, 2, '万') == '1,234.57'

    def test_unit_wan_zero_decimals(self):
        assert filter_num(12345678, 0, '万') == '1,235'

    def test_unit_yi(self):
        assert filter_num(1234567890, 2, '亿') == '12.35'

    def test_unit_qian(self):
        assert filter_num(12345678, 2, '千') == '12,345.68'

    def test_unit_wan_negative_decimals(self):
        assert filter_num(12345678, -2, '万') == '1,200'

    def test_unit_me_wan_negative_3(self):
        assert filter_num(12345678, -3, '万') == '1,000'

    def test_unknown_unit_treated_as_1(self):
        assert filter_num(1234567.89, 2, '未知') == '1,234,567.89'

    def test_invalid_value_returns_original(self):
        assert filter_num('abc') == 'abc'


class TestFilterPercent:
    def test_none_returns_empty(self):
        assert filter_percent(None) == ''

    def test_zero_returns_empty(self):
        assert filter_percent(0) == ''

    def test_decimal_formats(self):
        assert filter_percent(0.25) == '25.00%'

    def test_whole_number_formats(self):
        assert filter_percent(1) == '100.00%'


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

    def test_preset_short(self):
        assert filter_date('2025-01-15', 'short') == '2025/01/15'

    def test_preset_month(self):
        assert filter_date('2025-01-15', 'month') == '2025年01月'

    def test_preset_chinese(self):
        assert filter_date('2025-01-15', 'chinese') == '二〇二五年一月十五日'

    def test_preset_chinese_leap_day(self):
        assert filter_date('2024-02-29', 'chinese') == '二〇二四年二月二十九日'

    def test_preset_chinese_single_digit(self):
        assert filter_date('2025-03-05', 'chinese') == '二〇二五年三月五日'

    def test_preset_chinese_datetime(self):
        assert filter_date(datetime(2025, 10, 20), 'chinese') == '二〇二五年十月二十日'

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
