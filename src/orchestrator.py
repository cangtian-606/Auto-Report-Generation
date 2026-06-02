#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""渲染编排 — 读取数据 → 校验 → 渲染的完整流程"""

import os
import logging
import tempfile
from typing import Dict, Any, Optional

from .exceptions import TemplateSyntaxError
from .reader import create_reader
from .processing.table_preprocessors import TcInheritancePreprocessor, TrInheritancePreprocessor
from .render.generator import DocumentGenerator
from .render.pipeline import RenderPipeline
from .render.context import RenderContext
from .render.analyzer import TemplateAnalyzer
from .logging_config import get_stage_times, get_run_id

logger = logging.getLogger(__name__)

SEP = "─" * 54


def _preprocess_template(template_path: str, pipeline: RenderPipeline) -> str:
    preprocessors = [
        ("tr", TrInheritancePreprocessor()),
        ("tc", TcInheritancePreprocessor()),
    ]
    current = template_path
    for tag, preprocessor in preprocessors:
        suffix = os.path.splitext(current)[1] or ".docx"
        tmp_fd, tmp_path = tempfile.mkstemp(suffix=suffix, prefix=f"{tag}_pre_")
        os.close(tmp_fd)
        pipeline.track_temp(tmp_path)
        n = preprocessor.process(current, tmp_path)
        if n > 0:
            if current != template_path:
                pipeline.track_temp(current)
            current = tmp_path
        else:
            os.unlink(tmp_path)
    return current


def _check_syntax(analyzer: TemplateAnalyzer, context: Dict[str, Any],
                 strict: bool, report_unused: bool) -> None:
    undeclared = analyzer.get_undeclared(context)
    if undeclared is not None and undeclared:
        logger.warning("模板中有 %d 个变量在数据中未找到:", len(undeclared))
        for var in sorted(undeclared)[:10]:
            logger.warning("  - %s", var)
        if len(undeclared) > 10:
            logger.warning("  ... 还有 %d 个", len(undeclared) - 10)
        if strict:
            raise TemplateSyntaxError(f"未声明变量: {undeclared}")

    if report_unused:
        unused = analyzer.get_unused(context)
        if unused:
            logger.info("数据中有 %d 个字段未被模板使用:", len(unused))
            for item in sorted(unused)[:10]:
                logger.info("  - %s", item)
            if len(unused) > 10:
                logger.info("  ... 还有 %d 个", len(unused) - 10)


def _print_summary(state: Dict[str, Any]) -> None:
    stages = get_stage_times()
    total = sum(stages.values())
    stage_parts = []
    for name in ("读数据", "构建", "预处理", "渲染"):
        t = stages.get(name)
        stage_parts.append(f"{name}{t:.1f}" if t is not None else f"{name}—")

    logger.info(SEP)
    logger.info("  运行 ID: %s", get_run_id())
    logger.info("  数据文件: %s · %d Sheet · %d 行",
                state["data_name"], state["sheet_count"], state["row_count"])
    logger.info("  模板文件: %s · %d 表格 · %d 标签",
                state["template_name"], state["template_tables"], state["template_tags"])
    logger.info("  输出文件: %s", state["output_name"])
    logger.info("  执行结果: %s", "✓ 成功" if state["success"] else "✗ 失败")
    logger.info("  阶段耗时: %.1fs = %s", total, " + ".join(stage_parts))
    logger.info("  变量检查: %d 未定义", state.get("undeclared_count", 0))


def generate(data_path: str, template_path: str, output_path: str,
             strict: bool = False, check_vars: bool = True,
             validate: bool = False, strict_validate: bool = False,
             schema_path: Optional[str] = None,
             check_syntax: bool = False,
             report_unused: bool = False) -> bool:

    data_name = os.path.basename(data_path)
    template_name = os.path.basename(template_path)
    output_name = os.path.basename(output_path)

    logger.info("开始生成: %s → %s", data_name, output_name)

    pipeline = RenderPipeline()

    # ---- 阶段1: 读取数据 ----
    def read_stage(state, pipe):
        reader = create_reader(data_path)
        return {"raw_data": reader.read_all()}

    # ---- 阶段2: 构建上下文 ----
    def build_stage(state, pipe):
        rctx = RenderContext(state["raw_data"])
        context = rctx.build()

        kv_count = sum(1 for v in context.values() if isinstance(v, dict))
        table_count = sum(1 for v in context.values() if isinstance(v, list))
        row_count = sum(len(v) for v in context.values() if isinstance(v, list))
        flat_count = sum(1 for k in context if "." in k)

        logger.info("读取数据: %d Sheet · %d 键值对 · %d 表格 · %d 行",
                    len(state["raw_data"]), kv_count, table_count, row_count)
        logger.info("构建上下文: %d 条目 · %d 子表扁平化", len(context), flat_count)

        return {
            "context": context,
            "rctx": rctx,
            "kv_count": kv_count,
            "table_count": table_count,
            "row_count": row_count,
            "flat_count": flat_count,
        }

    # ---- 阶段3: 模板预处理 ----
    def preprocess_stage(state, pipe):
        render_template_path = _preprocess_template(template_path, pipe)
        analyzer = TemplateAnalyzer(render_template_path)
        template_tables, template_tags = analyzer.get_stats()
        logger.info("模板预处理完成")
        return {
            "render_template_path": render_template_path,
            "analyzer": analyzer,
            "template_tables": template_tables,
            "template_tags": template_tags,
        }

    # ---- 阶段4: 渲染 ----
    def render_stage(state, pipe):
        render_template_path = state["render_template_path"]
        context = state["context"]
        analyzer = state["analyzer"]
        undeclared_count = 0

        if check_vars:
            undeclared = analyzer.get_undeclared(context)
            if undeclared is not None and undeclared:
                undeclared_count = len(undeclared)
                names = ", ".join(sorted(undeclared)[:5])
                if len(undeclared) > 5:
                    names += f" ... 等{len(undeclared)}个"
                logger.warning("渲染: %d 个未定义 — %s", len(undeclared), names)
                if strict:
                    pipe.cleanup_temps()
                    raise TemplateSyntaxError(f"未声明变量: {undeclared}")

        if check_syntax:
            _check_syntax(analyzer, context, strict, report_unused)

        gen = DocumentGenerator(
            render_template_path,
            source_template_path=template_path,
        )
        success = gen.render(context, output_path, strict=strict)
        return {"success": success, "undeclared_count": undeclared_count}

    pipeline.add_stage("读数据", read_stage)
    pipeline.add_stage("构建", build_stage)
    pipeline.add_stage("预处理", preprocess_stage)
    pipeline.add_stage("渲染", render_stage)

    try:
        state = pipeline.run()

        if validate:
            from .logging_config import timed_stage
            with timed_stage("数据校验"):
                state["rctx"].validate(schema_path, strict=strict_validate)

        state.update({
            "data_name": data_name,
            "sheet_count": len(state.get("raw_data", {})),
            "output_name": output_name,
            "template_name": template_name,
        })

        _print_summary(state)
        return state["success"]
    finally:
        pipeline.cleanup_temps()
