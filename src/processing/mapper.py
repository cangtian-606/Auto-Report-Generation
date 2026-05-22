#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""数据映射器"""

from typing import Any, Dict, List

import pandas as pd

from ..exceptions import DataReadError


class DataMapper:
    """数据映射器"""

    def __init__(self, raw_data: Dict[str, Any]) -> None:
        self.raw_data = raw_data

    def build_context(self) -> Dict[str, Any]:
        context: Dict[str, Any] = {}

        sorted_sheets = sorted(
            self.raw_data.items(),
            key=lambda x: x[0].count('.'),
        )

        for sheet_name, data in sorted_sheets:
            if isinstance(data, dict):
                self._process_dict_sheet(context, sheet_name, data)
            elif isinstance(data, pd.DataFrame):
                self._process_df_sheet(context, sheet_name, data)
            elif isinstance(data, list):
                context[sheet_name] = data

        return context

    def _process_dict_sheet(self, context, sheet_name, data):
        if '.' in sheet_name:
            self._mount_or_build_date_dict(context, sheet_name, data)
        else:
            context[sheet_name] = {}
            for key, value in data.items():
                context[sheet_name][key.strip()] = self._convert_value(value)

    def _process_df_sheet(self, context, sheet_name, data):
        if '.' in sheet_name:
            if self._extract_parent_col(data) is not None:
                self._mount_or_build_form_df(context, sheet_name, data)
            else:
                self._mount_or_build_date_df(context, sheet_name, data)
        else:
            context[sheet_name] = data.to_dict('records')

    def _mount_or_build_date_dict(self, context, full_path, data: dict):
        parts = full_path.rsplit('.', 1)

        if len(parts) == 1:
            sheet_prefix = parts[0]
            for key, value in data.items():
                full_key = f"{sheet_prefix}.{key.strip()}"
                self._set_nested(context, full_key, self._convert_value(value))
            return

        parent_path, child_key = parts
        parent = self._get_nested(context, parent_path)

        if isinstance(parent, dict):
            parent[child_key] = {}
            for key, value in data.items():
                parent[child_key][key.strip()] = self._convert_value(value)
        else:
            for key, value in data.items():
                full_key = f"{full_path}.{key.strip()}"
                self._set_nested(context, full_key, self._convert_value(value))

    def _mount_or_build_form_df(self, context, full_path, data: pd.DataFrame):
        parts = full_path.rsplit('.', 1)

        if len(parts) == 1:
            context[parts[0]] = data.to_dict('records')
            return

        parent_path, child_key = parts
        records = data.to_dict('records')

        parent_col = self._extract_parent_col(data)
        if parent_col is None:
            raise DataReadError(
                f"子表 '{full_path}' 缺少 _parent_ 开头列，无法关联父表 '{parent_path}'"
            )

        parent_target = self._resolve_form_parent(context, parent_path)
        if parent_target is None:
            raise DataReadError(
                f"找不到父表 '{parent_path}'，无法挂载子表 '{full_path}'"
            )

        self._attach_children_by_column_flat(
            parent_target, records, parent_col, child_key, full_path
        )

    def _mount_or_build_date_df(self, context, full_path, data: pd.DataFrame):
        parts = full_path.rsplit('.', 1)

        if len(parts) == 1:
            context[parts[0]] = data.to_dict('records')
            return

        parent_path, child_key = parts
        records = data.to_dict('records')
        parent = self._get_nested(context, parent_path)

        if isinstance(parent, dict):
            parent[child_key] = records
        else:
            raise DataReadError(
                f"找不到父表 '{parent_path}'，无法挂载子表 '{full_path}'"
            )

    def _extract_parent_col(self, df: pd.DataFrame):
        for col in df.columns:
            col_str = str(col).strip()
            if col_str.startswith('_parent_'):
                return col_str[8:]
        return None

    def _resolve_form_parent(self, domain_root: Dict[str, Any], parent_path: str) -> List[dict]:
        parts = parent_path.split('.')
        current = domain_root
        for i, part in enumerate(parts):
            if isinstance(current, dict) and part in current:
                current = current[part]
            elif isinstance(current, list):
                remaining = parts[i:]
                parents = []
                for row in current:
                    inner = row
                    for rp in remaining:
                        if isinstance(inner, dict) and rp in inner:
                            inner = inner[rp]
                        else:
                            inner = None
                            break
                    if isinstance(inner, list):
                        parents.extend(inner)
                return parents
            else:
                return []
        if isinstance(current, list):
            return current
        return []

    def _attach_children_by_column_flat(self, parent_rows: List[dict],
                                         raw_records: List[dict],
                                         parent_col: str,
                                         child_key: str,
                                         table_name: str):
        if not parent_rows:
            raise DataReadError(f"父表 '{table_name}' 无可用行")

        parent_value_set = {str(r.get(parent_col, '')) for r in parent_rows}

        for rec in raw_records:
            match_val = str(rec.get(f'_parent_{parent_col}', ''))
            if match_val not in parent_value_set:
                raise DataReadError(
                    f"子表 '{table_name}' 中 '{match_val}' 在父表 '{parent_col}' 列中找不到匹配"
                )

        for row in parent_rows:
            row[child_key] = []

        for rec in raw_records:
            pval = str(rec.get(f'_parent_{parent_col}', ''))
            for row in parent_rows:
                if str(row.get(parent_col, '')) == pval:
                    clean = {k: v for k, v in rec.items() if not k.startswith('_parent_')}
                    row[child_key].append(clean)
                    break

        unmatched = [r for r in parent_rows if not r.get(child_key)]
        if unmatched:
            raise DataReadError(
                f"父表有 {len(unmatched)} 行在子表 '{table_name}' 中无匹配数据: "
                f"{[r.get(parent_col, '?') for r in unmatched]}"
            )

    def _set_nested(self, d: Dict[str, Any], path: str, value: Any) -> None:
        parts = path.split('.')
        current = d
        for part in parts[:-1]:
            if part not in current:
                current[part] = {}
            current = current[part]
        current[parts[-1]] = value

    def _get_nested(self, d: Dict[str, Any], path: str) -> Any:
        parts = path.split('.')
        current = d
        for part in parts:
            if isinstance(current, dict) and part in current:
                current = current[part]
            elif isinstance(current, list):
                return None
            else:
                return None
        return current

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