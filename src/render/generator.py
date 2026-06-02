#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 文档生成器

将含变量占位符的 docxtpl 模板与数据结合，生成填充好的 Word 文档。
"""

import os
import logging
from typing import Dict, Any, Optional

try:
    from docxtpl import DocxTemplate
except ImportError:
    raise ImportError("请先安装 docxtpl: pip install docxtpl") from None

from .analyzer import TemplateAnalyzer
from ..exceptions import TemplateSyntaxError

logger = logging.getLogger(__name__)


def _extract_error_info(exc: Exception) -> str:
    msg = str(exc)
    for line in msg.splitlines():
        line = line.strip()
        if line and "is undefined" in line:
            return line.split("is undefined")[0].strip().rstrip("'").lstrip("'")
    return msg.splitlines()[0].strip() if msg else "未知错误"


class DocumentGenerator:
    """docxtpl 文档生成器"""

    _template_cache: Dict[Any, DocxTemplate] = {}

    def __init__(
        self,
        template_path: str,
        source_template_path: Optional[str] = None,
    ) -> None:
        """文档生成器构造函数。

        Args:
            template_path: 实际要加载渲染的模板文件路径。预处理流程下指向临时文件。
            source_template_path: 用户提供的原始模板路径，仅用于日志/错误提示。
                若为 None，则回退到 template_path。
        """
        self.template_path = template_path
        self.source_template_path = source_template_path or template_path
        self._analyzer: Optional[TemplateAnalyzer] = None

    @property
    def analyzer(self) -> TemplateAnalyzer:
        if self._analyzer is None:
            self._analyzer = TemplateAnalyzer(self.template_path)
        return self._analyzer

    def _load_template(self) -> DocxTemplate:
        import os
        cache_key: Any
        try:
            mtime_ns = os.stat(self.template_path).st_mtime_ns
            cache_key = (self.template_path, mtime_ns)
        except Exception:
            # 若无法获取 mtime（例如临时路径删除），回退到纯路径
            cache_key = self.template_path
        if cache_key not in self._template_cache:
            # 日志显示用户原始路径，缓存键基于实际加载路径
            logger.debug("加载模板: %s", self.source_template_path)
            try:
                self._template_cache[cache_key] = DocxTemplate(self.template_path)
            except Exception:
                raise
        return self._template_cache[cache_key]

    def render(self, context: Dict[str, Any], output_path: str,
               strict: bool = False) -> bool:
        try:
            doc = self._load_template()
        except Exception:
            logger.exception("[模板加载] 渲染前置失败")
            return False

        logger.debug("执行渲染...")
        try:
            doc.render(context, jinja_env=self.analyzer.jinja_env)
        except Exception as e:
            error_info = _extract_error_info(e)
            logger.error("渲染: 失败 — %s", error_info)
            logger.debug("渲染异常详情", exc_info=True)
            if strict:
                raise TemplateSyntaxError(f"渲染失败: {error_info}") from e
            return False

        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        doc.save(output_path)
        logger.debug("保存文档: %s", output_path)
        return True
