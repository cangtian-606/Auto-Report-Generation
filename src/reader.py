#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Excel 数据读取器"""

import logging
from typing import Dict, Any

import pandas as pd

logger = logging.getLogger(__name__)


class ExcelDataReader:
    """Excel 数据读取器"""

    _KEY_NAMES = frozenset([
        "字段编码", "key", "名称", "变量", "field", "name", "code",
        "值", "value", "content", "内容",
    ])

    def __init__(self, file_path: str) -> None:
        self.file_path = file_path
        self.data: Dict[str, Any] = {}

    def read_all(self) -> Dict[str, Any]:
        xl = pd.ExcelFile(self.file_path, engine='openpyxl')
        logger.info(f"读取Excel: {self.file_path}")
        logger.info(f"工作表: {xl.sheet_names}")

        for sheet_name in xl.sheet_names:
            data = self._read_sheet(xl, sheet_name)
            self.data[sheet_name] = data

        return self.data

    def _read_sheet(self, xl, sheet_name: str) -> Any:
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None, nrows=5)

        if df.empty:
            return {}

        first_cell = str(df.iloc[0, 0]).strip().lower() if pd.notna(df.iloc[0, 0]) else ""

        if first_cell in self._KEY_NAMES:
            return self._read_key_value_sheet(xl, sheet_name)

        return self._read_table_sheet(xl, sheet_name)

    def _read_key_value_sheet(self, xl, sheet_name: str) -> Dict[str, Any]:
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
        result: Dict[str, Any] = {}
        old_format_keys = []

        for _, row in df.iterrows():
            key = row.iloc[0] if pd.notna(row.iloc[0]) else None
            value = row.iloc[1] if len(row) > 1 and pd.notna(row.iloc[1]) else None

            if key is not None:
                key_str = str(key).strip()
                if key_str.lower() in self._KEY_NAMES:
                    continue
                if not key_str:
                    continue
                if key_str.startswith(('date.', 'form.', 'g.', 'info.', 'flags.', 'notes.')):
                    old_format_keys.append(key_str)
                    continue
                result[key_str] = value

        if old_format_keys:
            logger.warning(f"  [{sheet_name}] 检测到旧格式字段编码（已自动忽略）: {old_format_keys}")

        logger.info(f"  [{sheet_name}] 键值对: {len(result)} 条")
        return result

    def _read_table_sheet(self, xl, sheet_name: str) -> pd.DataFrame:
        df = pd.read_excel(xl, sheet_name=sheet_name, engine='openpyxl')
        df.columns = [str(col).strip() for col in df.columns]
        logger.info(f"  [{sheet_name}] 表格: {len(df)} 行 x {len(df.columns)} 列")
        return df
