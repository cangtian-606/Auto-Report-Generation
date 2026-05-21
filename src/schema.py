#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""数据验证器"""

import json
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd

logger = logging.getLogger(__name__)


@dataclass
class FieldSpec:
    required: bool = False
    field_type: Optional[str] = None
    choices: Optional[List[Any]] = None


@dataclass
class SheetSchema:
    required: bool = False
    fields: Dict[str, FieldSpec] = field(default_factory=dict)


@dataclass
class DataSchema:
    sheets: Dict[str, SheetSchema] = field(default_factory=dict)

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'DataSchema':
        schema = cls()
        for sheet_name, sheet_spec in d.get('sheets', {}).items():
            sheet_schema = SheetSchema(
                required=sheet_spec.get('required', False),
            )
            for field_name, field_spec in sheet_spec.get('fields', {}).items():
                if isinstance(field_spec, dict):
                    sheet_schema.fields[field_name] = FieldSpec(
                        required=field_spec.get('required', False),
                        field_type=field_spec.get('field_type'),
                        choices=field_spec.get('choices'),
                    )
                elif isinstance(field_spec, bool):
                    sheet_schema.fields[field_name] = FieldSpec(required=field_spec)
            schema.sheets[sheet_name] = sheet_schema
        return schema

    @classmethod
    def from_json_file(cls, path: str) -> 'DataSchema':
        with open(path, 'r', encoding='utf-8') as f:
            return cls.from_dict(json.load(f))


class SchemaValidator:
    """Schema 验证器"""

    _BOOL_VALUES = {'true', 'false'}
    _TYPE_CHECKS = {
        'str': lambda v: isinstance(v, str) and v != '',
        'int': lambda v: isinstance(v, (int, float)) and not isinstance(v, bool),
        'float': lambda v: isinstance(v, (int, float)) and not isinstance(v, bool),
        'bool': lambda v: isinstance(v, bool),
        'bool_str': lambda v: isinstance(v, str) and v.strip().upper() in {'TRUE', 'FALSE'},
    }

    def __init__(self, schema: Optional[DataSchema] = None) -> None:
        self.schema = schema

    def load_from_dict(self, d: Dict[str, Any]) -> None:
        self.schema = DataSchema.from_dict(d)

    def load_from_file(self, path: str) -> None:
        self.schema = DataSchema.from_json_file(path)

    def validate(self, raw_data: Dict[str, Any]) -> List[str]:
        errors: List[str] = []

        if self.schema is None:
            return errors

        for sheet_name, sheet_schema in self.schema.sheets.items():
            if sheet_name not in raw_data:
                if sheet_schema.required:
                    errors.append(f"必填 Sheet 缺失: {sheet_name}")
                continue

            data = raw_data[sheet_name]

            if isinstance(data, dict):
                for field_name, field_spec in sheet_schema.fields.items():
                    value = data.get(field_name)

                    if field_spec.required and (value is None or value == ''):
                        errors.append(f"必填字段缺失: {sheet_name}.{field_name}")
                        continue

                    if value is None or value == '':
                        continue

                    if field_spec.field_type:
                        check = self._TYPE_CHECKS.get(field_spec.field_type)
                        if check and not check(value):
                            errors.append(
                                f"字段类型错误: {sheet_name}.{field_name} "
                                f"应为 {field_spec.field_type}，实际为 {type(value).__name__}"
                            )

                    if field_spec.choices is not None:
                        if value not in field_spec.choices:
                            errors.append(
                                f"字段值非法: {sheet_name}.{field_name} "
                                f"值应为 {field_spec.choices} 之一，实际为 {value}"
                            )

            elif isinstance(data, pd.DataFrame):
                if data.empty and sheet_schema.required:
                    errors.append(f"必填 Sheet 无数据行: {sheet_name}")
                    continue

                columns = set(str(c).strip() for c in data.columns)
                for field_name, field_spec in sheet_schema.fields.items():
                    if field_spec.required and field_name not in columns:
                        errors.append(f"必填列缺失: {sheet_name}.{field_name}")

        return errors
