#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""数据验证器 — 校验 Mapper 产出的 context"""

import json
import logging
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


@dataclass
class FieldSpec:
    required: bool = False
    field_type: Optional[str] = None
    choices: Optional[List[Any]] = None


@dataclass
class EntrySchema:
    required: bool = False
    fields: Dict[str, FieldSpec] = field(default_factory=dict)
    min_items: int = 0


@dataclass
class DataSchema:
    entries: Dict[str, EntrySchema] = field(default_factory=dict)

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'DataSchema':
        schema = cls()
        source = d.get('entries', d.get('sheets', {}))
        for entry_name, entry_spec in source.items():
            entry_schema = EntrySchema(
                required=entry_spec.get('required', False),
                min_items=entry_spec.get('min_items', 0),
            )
            for field_name, field_spec in entry_spec.get('fields', {}).items():
                if isinstance(field_spec, dict):
                    entry_schema.fields[field_name] = FieldSpec(
                        required=field_spec.get('required', False),
                        field_type=field_spec.get('field_type'),
                        choices=field_spec.get('choices'),
                    )
                elif isinstance(field_spec, bool):
                    entry_schema.fields[field_name] = FieldSpec(required=field_spec)
            schema.entries[entry_name] = entry_schema
        return schema

    @classmethod
    def from_json_file(cls, path: str) -> 'DataSchema':
        with open(path, 'r', encoding='utf-8') as f:
            return cls.from_dict(json.load(f))


class SchemaValidator:
    """Schema 验证器 — 基于 Mapper 产出的 context 进行校验"""

    _TYPE_CHECKS = {
        'str': lambda v: isinstance(v, str) and v != '',
        'int': lambda v: isinstance(v, int) and not isinstance(v, bool),
        'float': lambda v: isinstance(v, float) and not isinstance(v, bool),
        'bool': lambda v: isinstance(v, bool),
    }

    def __init__(self, schema: Optional[DataSchema] = None) -> None:
        self.schema = schema

    def load_from_dict(self, d: Dict[str, Any]) -> None:
        self.schema = DataSchema.from_dict(d)

    def load_from_file(self, path: str) -> None:
        self.schema = DataSchema.from_json_file(path)

    def validate(self, context: Dict[str, Any]) -> List[str]:
        errors: List[str] = []

        if self.schema is None:
            return errors

        for entry_name, entry_schema in self.schema.entries.items():
            if entry_name not in context:
                if entry_schema.required:
                    errors.append(f"必填条目缺失: {entry_name}")
                continue

            data = context[entry_name]

            if isinstance(data, list):
                self._validate_list(entry_name, data, entry_schema, errors)
            elif isinstance(data, dict):
                self._validate_dict(entry_name, data, entry_schema, errors)

        return errors

    def _validate_dict(self, name: str, data: dict,
                       schema: EntrySchema, errors: List[str]) -> None:
        for field_name, field_spec in schema.fields.items():
            value = data.get(field_name)
            if field_spec.required and (value is None or value == ''):
                errors.append(f"必填字段缺失: {name}.{field_name}")
                continue
            if value is None or value == '':
                continue
            self._check_type_and_choices(name, field_name, value, field_spec, errors)

    def _validate_list(self, name: str, data: list,
                       schema: EntrySchema, errors: List[str]) -> None:
        if not data and schema.required:
            errors.append(f"必填列表无数据: {name}")
            return

        if schema.min_items and len(data) < schema.min_items:
            errors.append(
                f"列表条目不足: {name} 至少需要 {schema.min_items} 条，"
                f"实际 {len(data)} 条"
            )

        for idx, item in enumerate(data):
            if not isinstance(item, dict):
                continue
            prefix = f"{name}[{idx}]"
            for field_name, field_spec in schema.fields.items():
                value = item.get(field_name)
                if field_spec.required and (value is None or value == ''):
                    errors.append(f"必填字段缺失: {prefix}.{field_name}")
                    continue
                if value is None or value == '':
                    continue
                self._check_type_and_choices(
                    prefix, field_name, value, field_spec, errors
                )

    def _check_type_and_choices(self, prefix: str, field_name: str,
                                 value: Any, field_spec: FieldSpec,
                                 errors: List[str]) -> None:
        if field_spec.field_type:
            check = self._TYPE_CHECKS.get(field_spec.field_type)
            if check and not check(value):
                errors.append(
                    f"字段类型错误: {prefix}.{field_name} "
                    f"应为 {field_spec.field_type}，实际为 {type(value).__name__}"
                )

        if field_spec.choices is not None:
            if value not in field_spec.choices:
                errors.append(
                    f"字段值非法: {prefix}.{field_name} "
                    f"值应为 {field_spec.choices} 之一，实际为 {value}"
                )
