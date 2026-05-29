#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Jinja2 自定义过滤器"""

from datetime import datetime as dt
from typing import Any

_UNITS = {
    '十': 10, '百': 100, '千': 1000,
    '万': 10000, '十万': 100000, '百万': 1000000,
    '千万': 10000000, '亿': 100000000, '十亿': 1000000000,
}

_DATE_PRESETS = {
    'short': '%Y/%m/%d',
    'month': '%Y年%m月',
    'chinese': 'chinese',
}

_CN_NUMS = '〇一二三四五六七八九'
_CN_TENS = ['', '十', '二十', '三十']


def _to_chinese_num(n: int) -> str:
    if 1 <= n <= 9:
        return _CN_NUMS[n]
    if n == 10:
        return '十'
    if 11 <= n <= 19:
        return '十' + _CN_NUMS[n - 10]
    if 20 <= n <= 31:
        tens = n // 10
        ones = n % 10
        if ones == 0:
            return _CN_TENS[tens]
        return _CN_TENS[tens] + _CN_NUMS[ones]
    return str(n)


def _to_chinese_date(d: dt) -> str:
    year_chars = ''.join(_CN_NUMS[int(c)] for c in str(d.year))
    return f"{year_chars}年{_to_chinese_num(d.month)}月{_to_chinese_num(d.day)}日"


def _is_empty(value: Any) -> bool:
    return value is None or value == ''


def filter_percent(value: Any) -> str:
    if _is_empty(value):
        return ''
    try:
        num = float(value)
        if num == 0:
            return ''
        return f"{num * 100:.2f}%"
    except (ValueError, TypeError):
        return str(value)


def filter_num(value: Any, decimals: int = 2, unit: str = '') -> str:
    if _is_empty(value):
        return ''
    try:
        num = float(value)
    except (ValueError, TypeError):
        return str(value)
    if num == 0:
        return ''
    divisor = _UNITS.get(unit, 1)
    num = num / divisor
    if decimals >= 0:
        return f"{num:,.{decimals}f}"
    return f"{int(round(num, decimals)):,}"


def filter_date(value: Any, fmt: str = '%Y年%m月%d日') -> str:
    if _is_empty(value):
        return ''
    actual_fmt = _DATE_PRESETS.get(fmt, fmt)
    if isinstance(value, dt):
        if actual_fmt == 'chinese':
            return _to_chinese_date(value)
        return value.strftime(actual_fmt)
    for pattern in ('%Y-%m-%d', '%Y/%m/%d', '%Y年%m月%d日'):
        try:
            parsed = dt.strptime(str(value), pattern)
            if actual_fmt == 'chinese':
                return _to_chinese_date(parsed)
            return parsed.strftime(actual_fmt)
        except ValueError:
            continue
    return str(value)


def filter_default_dash(value: Any) -> str:
    if _is_empty(value):
        return ''
    return str(value)


def filter_default(value: Any, default: str = '') -> str:
    if _is_empty(value):
        return default
    return str(value)


def filter_int(value: Any) -> str:
    if _is_empty(value):
        return ''
    try:
        return str(int(float(value)))
    except (ValueError, TypeError):
        return str(value)


def filter_paragraphs(value: Any) -> str:
    if _is_empty(value):
        return ''
    return str(value).replace('\n', '\a')


FILTERS = {
    'money': filter_num,
    'percent': filter_percent,
    'num': filter_num,
    'date': filter_date,
    'default_dash': filter_default_dash,
    'default': filter_default,
    'int': filter_int,
    'paragraphs': filter_paragraphs,
}
