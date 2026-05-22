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
            df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
            self.data[sheet_name] = self._parse_sheet(sheet_name, df)

        return self.data

    def _parse_sheet(self, sheet_name: str, df: pd.DataFrame) -> Any:
        if df.empty:
            return {}

        first_cell = str(df.iloc[0, 0]).strip().lower() if pd.notna(df.iloc[0, 0]) else ""

        if first_cell in self._KEY_NAMES:
            return self._parse_key_value(df)
        return self._parse_table(sheet_name, df)

    def _parse_key_value(self, df: pd.DataFrame) -> Dict[str, Any]:
        result: Dict[str, Any] = {}
        for _, row in df.iterrows():
            key = row.iloc[0] if pd.notna(row.iloc[0]) else None
            value = row.iloc[1] if len(row) > 1 and pd.notna(row.iloc[1]) else None
            if key is not None:
                key_str = str(key).strip()
                if key_str.lower() in self._KEY_NAMES:
                    continue
                if not key_str:
                    continue
                result[key_str] = value
        return result

    def _parse_table(self, sheet_name: str, df: pd.DataFrame) -> pd.DataFrame:
        df.columns = [str(col).strip() for col in df.iloc[0]]
        df = df.iloc[1:].reset_index(drop=True)
        logger.info(f"  [{sheet_name}] 表格: {len(df)} 行 x {len(df.columns)} 列")
        return df
