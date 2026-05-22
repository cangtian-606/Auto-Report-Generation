#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Jinja2 自定义过滤器"""

from datetime import datetime
from typing import Any


def filter_money(value: Any) -> str:
    if value is None or value == '':
        return ''
    try:
        num = float(value)
        if num == 0:
            return ''
        return f"{num:,.2f}"
    except (ValueError, TypeError):
        return str(value)


def filter_percent(value: Any) -> str:
    if value is None or value == '':
        return ''
    try:
        num = float(value)
        if num == 0:
            return ''
        return f"{num * 100:.2f}%"
    except (ValueError, TypeError):
        return str(value)


def filter_num(value: Any) -> str:
    if value is None or value == '':
        return ''
    try:
        num = float(value)
        if num == 0:
            return ''
        if num == int(num):
            return f"{int(num):,}"
        return f"{num:,.2f}"
    except (ValueError, TypeError):
        return str(value)


def filter_date(value: Any, fmt: str = '%Y年%m月%d日') -> str:
    if value is None or value == '':
        return ''
    for pattern in ('%Y-%m-%d', '%Y/%m/%d', '%Y年%m月%d日'):
        try:
            dt = datetime.strptime(str(value), pattern)
            return dt.strftime(fmt)
        except ValueError:
            continue
    return str(value)


def filter_default_dash(value: Any) -> str:
    if value is None or value == '' or value == 0:
        return '-'
    return str(value)


def filter_default(value: Any, default: str = '') -> str:
    if value is None or value == '':
        return default
    return str(value)


def filter_int(value: Any) -> str:
    if value is None or value == '':
        return ''
    try:
        return str(int(float(value)))
    except (ValueError, TypeError):
        return str(value)


def filter_str(value: Any) -> str:
    if value is None:
        return ''
    return str(value)


FILTERS = {
    'money': filter_money,
    'percent': filter_percent,
    'num': filter_num,
    'date': filter_date,
    'default_dash': filter_default_dash,
    'default': filter_default,
    'int': filter_int,
    'str': filter_str,
}
