#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 文档生成器

将含变量占位符的 docxtpl 模板与数据结合，生成填充好的 Word 文档。
"""

import os
import sys
import logging
from typing import Dict, List, Any, Optional

try:
    from docxtpl import DocxTemplate
except ImportError:
    print("错误：请先安装 docxtpl: pip install docxtpl")
    sys.exit(1)

try:
    from jinja2 import Environment
except ImportError:
    print("错误：请先安装 jinja2: pip install jinja2")
    sys.exit(1)

from ..reader.xlsx import ExcelDataReader
from ..reader.yaml import YamlDataReader
from ..processing.mapper import DataMapper
from .filters import FILTERS
from ..exceptions import TemplateSyntaxError, ValidationError

logger = logging.getLogger(__name__)


class DocumentGenerator:
    """docxtpl 文档生成器"""

    _template_cache: Dict[str, DocxTemplate] = {}

    def __init__(self, template_path: str) -> None:
        self.template_path = template_path

    def _load_template(self) -> DocxTemplate:
        if self.template_path not in self._template_cache:
            logger.info(f"加载模板: {self.template_path}")
            try:
                self._template_cache[self.template_path] = DocxTemplate(self.template_path)
            except Exception as e:
                logger.error(f"加载模板失败: {e}")
                raise
        return self._template_cache[self.template_path]

    def render(self, context: Dict[str, Any], output_path: str,
               strict: bool = False) -> bool:
        try:
            doc = self._load_template()
        except Exception:
            return False

        jinja_env = Environment()
        for name, func in FILTERS.items():
            jinja_env.filters[name] = func

        logger.info("执行渲染...")
        try:
            doc.render(context, jinja_env=jinja_env)
        except Exception as e:
            logger.error(f"渲染失败: {e}")
            if strict:
                raise
            return False

        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        doc.save(output_path)
        logger.info(f"保存文档: {output_path}")
        return True

    def get_undeclared_variables(self, context: Dict[str, Any]) -> set:
        try:
            doc = self._load_template()
            jinja_env = Environment()
            for name, func in FILTERS.items():
                jinja_env.filters[name] = func
            return doc.get_undeclared_template_variables(context=context, jinja_env=jinja_env)
        except Exception as e:
            logger.error(f"获取未定义变量失败: {e}")
            return set()

    def check_syntax(self, context: Dict[str, Any],
                     strict: bool = False,
                     report_unused: bool = False) -> bool:
        undeclared = self.get_undeclared_variables(context)
        has_issue = False

        if undeclared:
            has_issue = True
            logger.warning(f"模板中有 {len(undeclared)} 个变量在数据中未找到:")
            for var in sorted(undeclared)[:10]:
                logger.warning(f"  - {var}")
            if len(undeclared) > 10:
                logger.warning(f"  ... 还有 {len(undeclared) - 10} 个")
            if strict:
                raise TemplateSyntaxError(f"未声明变量: {undeclared}")

        if report_unused:
            unused = self._find_unused_data(context)
            if unused:
                has_issue = True
                logger.info(f"数据中有 {len(unused)} 个字段未被模板使用:")
                for item in sorted(unused)[:10]:
                    logger.info(f"  - {item}")
                if len(unused) > 10:
                    logger.info(f"  ... 还有 {len(unused) - 10} 个")

        return not has_issue

    def _find_unused_data(self, context: Dict[str, Any]) -> List[str]:
        unused: List[str] = []
        for key, value in context.items():
            if isinstance(value, dict):
                for field_key in value.keys():
                    unused.append(f"{key}.{field_key}")
        return unused


def _build_context(data_path: str) -> Dict[str, Any]:
    ext = os.path.splitext(data_path)[1].lower()
    if ext in ('.yaml', '.yml'):
        reader = YamlDataReader(data_path)
    else:
        reader = ExcelDataReader(data_path)
    raw_data = reader.read_all()
    mapper = DataMapper(raw_data)
    context = mapper.build_context()
    logger.info(f"上下文统计: {len(context)} 个顶级条目")
    return context


def _validate_context(context: Dict[str, Any],
                      schema_path: Optional[str],
                      strict_validate: bool) -> None:
    from ..processing.schema import SchemaValidator
    validator = SchemaValidator()
    if schema_path:
        validator.load_from_file(schema_path)
    errors = validator.validate(context)
    if errors:
        for e in errors:
            logger.error(f"[数据验证] {e}")
        if strict_validate:
            raise ValidationError(f"数据验证失败: {len(errors)} 个错误")


def _check_and_render(context: Dict[str, Any],
                      template_path: str, output_path: str,
                      strict: bool, check_vars: bool,
                      check_syntax: bool, report_unused: bool) -> bool:
    gen = DocumentGenerator(template_path)

    if check_vars:
        undeclared = gen.get_undeclared_variables(context)
        if undeclared:
            logger.warning(f"发现 {len(undeclared)} 个未定义的模板变量:")
            for var in sorted(undeclared)[:10]:
                logger.warning(f"  - {var}")
            if len(undeclared) > 10:
                logger.warning(f"  ... 还有 {len(undeclared) - 10} 个")
            if strict:
                raise TemplateSyntaxError(f"未声明变量: {undeclared}")

    if check_syntax:
        gen.check_syntax(context, strict=strict, report_unused=report_unused)

    return gen.render(context, output_path, strict=strict)


def generate(data_path: str, template_path: str, output_path: str,
             strict: bool = False, check_vars: bool = True,
             validate: bool = False, strict_validate: bool = False,
             schema_path: Optional[str] = None,
             check_syntax: bool = False,
             report_unused: bool = False) -> bool:
    context = _build_context(data_path)

    if validate:
        _validate_context(context, schema_path, strict_validate)

    return _check_and_render(
        context, template_path, output_path,
        strict=strict, check_vars=check_vars,
        check_syntax=check_syntax, report_unused=report_unused,
    )
