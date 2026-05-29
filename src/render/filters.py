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
    if isinstance(value, dt):
        return value.strftime(fmt)
    for pattern in ('%Y-%m-%d', '%Y/%m/%d', '%Y年%m月%d日'):
        try:
            return dt.strptime(str(value), pattern).strftime(fmt)
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
