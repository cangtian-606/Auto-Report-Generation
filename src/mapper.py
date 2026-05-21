#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""数据映射器"""

from typing import Any, Dict

import pandas as pd


class DataMapper:
    """数据映射器"""

    def __init__(self, raw_data: Dict[str, Any]) -> None:
        self.raw_data = raw_data

    def build_context(self) -> Dict[str, Any]:
        context: Dict[str, Any] = {
            'date': {},
            'form': {},
        }

        for sheet_name, data in self.raw_data.items():
            if isinstance(data, dict):
                if sheet_name.startswith('date.'):
                    sheet_prefix = sheet_name[5:]
                    for key, value in data.items():
                        key = key.strip()
                        full_key = f"{sheet_prefix}.{key}"
                        self._set_nested(context['date'], full_key, self._convert_value(value))
                else:
                    for key, value in data.items():
                        context['date'][key.strip()] = self._convert_value(value)

            elif isinstance(data, pd.DataFrame):
                if sheet_name.startswith('form.'):
                    context['form'][sheet_name[5:]] = data.to_dict('records')
                else:
                    context['form'][sheet_name] = data.to_dict('records')

        return context

    def _set_nested(self, d: Dict[str, Any], path: str, value: Any) -> None:
        parts = path.split('.')
        current = d
        for part in parts[:-1]:
            if part not in current:
                current[part] = {}
            current = current[part]
        current[parts[-1]] = value

    def _convert_value(self, value: Any) -> Any:
        if value is None:
            return ''
        if isinstance(value, str):
            upper = value.strip().upper()
            if upper == 'TRUE':
                return True
            if upper == 'FALSE':
                return False
        return value
