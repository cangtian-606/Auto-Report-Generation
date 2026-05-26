#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""渲染编排 — 读取数据 → 校验 → 渲染的完整流程"""

import os
import logging
import tempfile
from typing import Dict, Any, Optional, List, Tuple

from .exceptions import DataReadError, ValidationError, TemplateSyntaxError
from .reader import create_reader
from .processing.mapper import DataMapper
from .processing.table_preprocessors import TcInheritancePreprocessor, TrInheritancePreprocessor
from .render.generator import DocumentGenerator
from .logging_config import timed_stage, get_stage_times, get_run_id, reset_stage_times

logger = logging.getLogger(__name__)

SEP = "─" * 54


def _template_stats(template_path: str) -> Tuple[int, int]:
    import zipfile
    from lxml import etree
    import re
    W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    tables = 0
    tags = 0
    try:
        with zipfile.ZipFile(template_path, "r") as z:
            xml = z.read("word/document.xml")
        root = etree.fromstring(xml)
        for tbl in root.iter(f"{{{W_NS}}}tbl"):
            tables += 1
        tag_re = re.compile(r"\{[{%].*?[%}]\}")
        for elem in root.iter():
            if elem.text and tag_re.search(elem.text):
                tags += 1
            if elem.tail and tag_re.search(elem.tail):
                tags += 1
    except Exception:
        pass
    return tables, tags




def _flatten_nested_lists(context: Dict[str, Any]) -> None:
    """将嵌套在列表型主表中的子表扁平化为顶级键，供模板 {%tr for i in xxx.子表 %} 访问。"""
    for key, value in list(context.items()):
        if not isinstance(value, list) or not value:
            continue
        if not isinstance(value[0], dict):
            continue
        sub_keys = [k for k in value[0] if isinstance(value[0][k], list)]
        for sub_key in sub_keys:
            flat_key = f"{key}.{sub_key}"
            if flat_key not in context:
                flat_list = []
                for item in value:
                    if sub_key in item:
                        flat_list.extend(item[sub_key])
                if flat_list:
                    context[flat_key] = flat_list


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
            if current != template_path:
                os.unlink(current)
            current = tmp_path
        else:
            os.unlink(tmp_path)
    return current


def generate(data_path: str, template_path: str, output_path: str,
             strict: bool = False, check_vars: bool = True,
             validate: bool = False, strict_validate: bool = False,
             schema_path: Optional[str] = None,
             check_syntax: bool = False,
             report_unused: bool = False) -> bool:
    from .logging_config import init_run_id
    init_run_id()
    reset_stage_times()

    data_name = os.path.basename(data_path)
    template_name = os.path.basename(template_path)
    output_name = os.path.basename(output_path)

    logger.info("开始生成: %s → %s", data_name, output_name)

    # ---- 阶段1: 读取数据 ----
    with timed_stage("读数据"):
        reader = create_reader(data_path)
        raw_data = reader.read_all()

    sheet_count = len(raw_data)

    # ---- 阶段2: 构建上下文 ----
    with timed_stage("构建"):
        mapper = DataMapper(raw_data)
        context = mapper.build_context()
        _flatten_nested_lists(context)

    kv_count = sum(1 for v in context.values() if isinstance(v, dict))
    table_count = sum(1 for v in context.values() if isinstance(v, list))
    row_count = sum(len(v) for v in context.values() if isinstance(v, list))
    flat_count = sum(1 for k in context if "." in k)

    logger.info("读取数据: %d Sheet · %d 键值对 · %d 表格 · %d 行",
                sheet_count, kv_count, table_count, row_count)
    logger.info("构建上下文: %d 条目 · %d 子表扁平化", len(context), flat_count)

    if validate:
        with timed_stage("数据校验"):
            _validate_context(context, schema_path, strict_validate)

    # ---- 阶段3: 模板预处理 ----
    with timed_stage("预处理"):
        render_template_path = _preprocess_template(template_path)

    # ---- 阶段4: 渲染 ----
    template_tables, template_tags = _template_stats(render_template_path)
    logger.info("模板预处理完成")

    undeclared_count = 0
    with timed_stage("渲染"):
        gen = DocumentGenerator(render_template_path)

        if check_vars:
            undeclared = gen.get_undeclared_variables(context)
            if undeclared is not None and undeclared:
                undeclared_count = len(undeclared)
                names = ", ".join(sorted(undeclared)[:5])
                if len(undeclared) > 5:
                    names += f" ... 等{len(undeclared)}个"
                logger.warning("渲染: %d 个未定义 — %s", len(undeclared), names)
                if strict:
                    if render_template_path != template_path:
                        try:
                            os.unlink(render_template_path)
                        except OSError:
                            pass
                    raise TemplateSyntaxError(f"未声明变量: {undeclared}")

        if check_syntax:
            gen.check_syntax(context, strict=strict, report_unused=report_unused)

        success = gen.render(context, output_path, strict=strict)

    try:
        if render_template_path != template_path:
            os.unlink(render_template_path)
    except OSError:
        pass

    # ---- 汇总摘要 ----
    stages = get_stage_times()
    total = sum(stages.values())
    stage_parts = []
    for name in ("读数据", "构建", "预处理", "渲染"):
        t = stages.get(name)
        if t is not None:
            stage_parts.append(f"{name}{t:.1f}")
        else:
            stage_parts.append(f"{name}—")
    stage_str = " + ".join(stage_parts)

    logger.info(SEP)
    logger.info("  运行 ID: %s", get_run_id())
    logger.info("  数据文件: %s · %d Sheet · %d 行", data_name, sheet_count, row_count)
    logger.info("  模板文件: %s · %d 表格 · %d 标签", template_name, template_tables, template_tags)
    logger.info("  输出文件: %s", output_name)
    if success:
        result_str = "✓ 成功"
    else:
        result_str = "✗ 失败"
    logger.info("  执行结果: %s", result_str)
    logger.info("  阶段耗时: %.1fs = %s", total, stage_str)
    check_str = f"{undeclared_count} 未定义"
    logger.info("  变量检查: %s", check_str)

    return success
