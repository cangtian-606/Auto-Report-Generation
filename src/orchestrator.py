#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""渲染编排 — 读取数据 → 校验 → 渲染的完整流程"""

import os
import logging
import tempfile
from typing import Dict, Any, Optional

from .exceptions import DataReadError, ValidationError, TemplateSyntaxError
from .reader import create_reader
from .processing.mapper import DataMapper
from .processing.table_preprocessors import TcInheritancePreprocessor, TrInheritancePreprocessor
from .render.generator import DocumentGenerator

logger = logging.getLogger(__name__)


def _build_context(data_path: str) -> Dict[str, Any]:
    try:
        reader = create_reader(data_path)
        raw_data = reader.read_all()
        mapper = DataMapper(raw_data)
        context = mapper.build_context()
    except DataReadError:
        raise
    except FileNotFoundError:
        raise DataReadError(f"文件不存在: {data_path}") from None
    except Exception as e:
        raise DataReadError(f"读取数据失败: {e}") from e
    logger.debug("上下文统计: %d 个顶级条目", len(context))
    return context


def _validate_context(context: Dict[str, Any],
                      schema_path: Optional[str],
                      strict_validate: bool) -> None:
    from .processing.schema import SchemaValidator
    validator = SchemaValidator()
    if schema_path:
        try:
            validator.load_from_file(schema_path)
        except FileNotFoundError:
            raise DataReadError(f"Schema 文件不存在: {schema_path}")
        except Exception as e:
            raise DataReadError(f"Schema 文件解析失败: {e}") from e
    errors = validator.validate(context)
    if errors:
        for e in errors:
            logger.error("[数据验证] %s", e)
        if strict_validate:
            raise ValidationError(f"数据验证失败: {len(errors)} 个错误")


def _preprocess_template(template_path: str) -> str:
    """对模板依次执行 tr/tc 继承预处理，返回临时文件路径。"""
    preprocessors = [
        ("tr", TrInheritancePreprocessor()),
        ("tc", TcInheritancePreprocessor()),
    ]
    current = template_path
    for tag, preprocessor in preprocessors:
        suffix = os.path.splitext(current)[1] or ".docx"
        tmp_fd, tmp_path = tempfile.mkstemp(suffix=suffix, prefix=f"{tag}_pre_")
        os.close(tmp_fd)
        n = preprocessor.process(current, tmp_path)
        if n > 0:
            logger.debug("%s 继承预处理: %d 个表格已注入循环标签", tag, n)
            if current != template_path:
                os.unlink(current)
            current = tmp_path
        else:
            os.unlink(tmp_path)
    return current


def _check_and_render(context: Dict[str, Any],
                      template_path: str, output_path: str,
                      strict: bool, check_vars: bool,
                      check_syntax: bool, report_unused: bool) -> bool:
    render_template_path = _preprocess_template(template_path)
    try:
        gen = DocumentGenerator(render_template_path)

        if check_vars:
            undeclared = gen.get_undeclared_variables(context)
            if undeclared is not None:
                if undeclared:
                    logger.warning("发现 %d 个未定义的模板变量:", len(undeclared))
                    for var in sorted(undeclared)[:10]:
                        logger.warning("  - %s", var)
                    if len(undeclared) > 10:
                        logger.warning("  ... 还有 %d 个", len(undeclared) - 10)
                    if strict:
                        raise TemplateSyntaxError(f"未声明变量: {undeclared}")

        if check_syntax:
            gen.check_syntax(context, strict=strict, report_unused=report_unused)

        return gen.render(context, output_path, strict=strict)
    finally:
        if render_template_path != template_path and os.path.exists(render_template_path):
            os.unlink(render_template_path)


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
