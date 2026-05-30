#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""RenderContext — 渲染上下文

将 context 构建、扁平化、校验聚合为一个深度模块。
调用方只需一次 build() 调用即可得到完整、可用的渲染上下文。
"""

import logging
from typing import Any, Dict, Optional, List

from ..processing.mapper import DataMapper

logger = logging.getLogger(__name__)


def _flatten_nested_lists(context: Dict[str, Any]) -> None:
    for key, value in list(context.items()):
        if not isinstance(value, list) or not value:
            continue
        if not isinstance(value[0], dict):
            continue
        sub_keys = [k for k in value[0] if isinstance(value[0][k], list)]
        for sub_key in sub_keys:
            flat_key = f"{key}.{sub_key}"
            if flat_key not in context:
                flat_list: List[Any] = []
                for item in value:
                    if sub_key in item:
                        flat_list.extend(item[sub_key])
                if flat_list:
                    context[flat_key] = flat_list

    _fix_dotted_resolution(context)


def _fix_dotted_resolution(context: Dict[str, Any]) -> None:
    """为被 non-dict 父键阻挡的点号键创建 Jinja2 别名。

    当 context 同时存在 "项目建设情况" (list) 和 "项目建设情况.固定资产" (list) 时，
    Jinja2 的 {{ 项目建设情况.固定资产 }} 会先解析 "项目建设情况" → list，
    再尝试 list["固定资产"] → 失败，永远不会触达扁平键。

    修复：将点替换为双下划线作为别名，模板用 {{ X__Y }} 访问。
    """
    for flat_key, value in list(context.items()):
        if '.' not in flat_key:
            continue
        first_part = flat_key.split('.')[0]
        if first_part not in context:
            continue
        parent = context[first_part]
        if isinstance(parent, dict):
            continue
        alias = flat_key.replace('.', '__')
        if alias not in context:
            context[alias] = value


class RenderContext:
    """渲染上下文 — 从 raw_data 构建完整的 Jinja2 渲染上下文。

    Usage:
        rctx = RenderContext(raw_data)
        context = rctx.build()          # 构建 + 扁平化
        rctx.validate(schema_path)      # 可选校验
    """

    def __init__(self, raw_data: Dict[str, Any]) -> None:
        self._raw_data = raw_data
        self._context: Optional[Dict[str, Any]] = None

    def build(self) -> Dict[str, Any]:
        mapper = DataMapper(self._raw_data)
        self._context = mapper.build_context()
        _flatten_nested_lists(self._context)
        return self._context

    @property
    def context(self) -> Dict[str, Any]:
        if self._context is None:
            raise RuntimeError("请先调用 build() 构建上下文")
        return self._context

    def validate(self, schema_path: Optional[str] = None,
                 strict: bool = False) -> List[str]:
        from ..processing.schema import SchemaValidator
        validator = SchemaValidator()
        if schema_path:
            try:
                validator.load_from_file(schema_path)
            except FileNotFoundError:
                from ..exceptions import DataReadError
                raise DataReadError(f"Schema 文件不存在: {schema_path}")
            except Exception as e:
                from ..exceptions import DataReadError
                raise DataReadError(f"Schema 文件解析失败: {e}") from e

        errors = validator.validate(self.context)
        if errors:
            for e in errors:
                logger.error("[数据验证] %s", e)
            if strict:
                from ..exceptions import ValidationError
                raise ValidationError(f"数据验证失败: {len(errors)} 个错误")
        return errors
