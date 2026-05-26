#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""TemplateAnalyzer — 模板自省

聚合模板结构分析（表格数、标签数）和变量分析（未定义、未使用），
消除 SandboxedEnvironment 和过滤器重复创建。
"""

import re
import zipfile
import logging
from typing import Dict, List, Any, Optional, Set, Tuple

from lxml import etree

try:
    from docxtpl import DocxTemplate
except ImportError:
    raise ImportError("请先安装 docxtpl: pip install docxtpl") from None

try:
    from jinja2.sandbox import SandboxedEnvironment
except ImportError:
    raise ImportError("请先安装 jinja2: pip install jinja2") from None

from .filters import FILTERS

logger = logging.getLogger(__name__)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class TemplateAnalyzer:
    """模板分析器 — 加载模板一次，执行多种分析。

    Usage:
        analyzer = TemplateAnalyzer("template.docx")
        tables, tags = analyzer.get_stats()
        undeclared = analyzer.get_undeclared(context)
    """

    def __init__(self, template_path: str) -> None:
        self._template_path = template_path
        self._doc: Optional[DocxTemplate] = None
        self._jinja_env: Optional[SandboxedEnvironment] = None
        self._stats: Optional[Tuple[int, int]] = None

    def _ensure_loaded(self) -> None:
        if self._doc is None:
            self._doc = DocxTemplate(self._template_path)
            self._jinja_env = SandboxedEnvironment()
            for name, func in FILTERS.items():
                self._jinja_env.filters[name] = func

    def get_stats(self) -> Tuple[int, int]:
        if self._stats is not None:
            return self._stats

        tables = 0
        tags = 0
        try:
            with zipfile.ZipFile(self._template_path, "r") as z:
                xml = z.read("word/document.xml")
            root = etree.fromstring(xml)
            for _ in root.iter(f"{{{W_NS}}}tbl"):
                tables += 1
            tag_re = re.compile(r"\{[{%].*?[%}]\}")
            for elem in root.iter():
                if elem.text and tag_re.search(elem.text):
                    tags += 1
                if elem.tail and tag_re.search(elem.tail):
                    tags += 1
        except Exception:
            pass
        self._stats = (tables, tags)
        return self._stats

    def get_undeclared(self, context: Dict[str, Any]) -> Optional[Set[str]]:
        self._ensure_loaded()
        try:
            return self._doc.get_undeclared_template_variables(
                context=context, jinja_env=self._jinja_env
            )
        except Exception:
            logger.exception("获取未定义变量失败")
            return None

    def get_unused(self, context: Dict[str, Any]) -> List[str]:
        unused: List[str] = []
        template_vars = self._collect_template_vars()

        def traverse(value: Any, path_prefix: str = ""):
            if isinstance(value, dict):
                for key, val in value.items():
                    current_path = f"{path_prefix}.{key}" if path_prefix else key
                    if current_path not in template_vars:
                        if not isinstance(val, (dict, list)):
                            unused.append(current_path)
                    traverse(val, current_path)
            elif isinstance(value, list):
                for idx, item in enumerate(value):
                    if isinstance(item, (dict, list)):
                        list_path = f"{path_prefix}[{idx}]" if path_prefix else f"[{idx}]"
                        traverse(item, list_path)

        for key, value in context.items():
            current_path = key
            if isinstance(value, (dict, list)):
                traverse(value, current_path)
            else:
                if current_path not in template_vars:
                    unused.append(current_path)
        return unused

    def _collect_template_vars(self) -> Set[str]:
        self._ensure_loaded()
        var_pattern = re.compile(r"\{\{\s*(.+?)\s*(?:\||\}\})")
        template_vars: Set[str] = set()
        try:
            real_doc = self._doc.get_docx()
            for elem in real_doc.part._element.iter():
                if elem.text:
                    for match in var_pattern.finditer(elem.text):
                        var_path = match.group(1).strip()
                        if var_path:
                            template_vars.add(var_path)
                if elem.tail:
                    for match in var_pattern.finditer(elem.tail):
                        var_path = match.group(1).strip()
                        if var_path:
                            template_vars.add(var_path)
        except Exception:
            logger.exception("收集模板变量失败")
        return template_vars

    @property
    def doc(self) -> DocxTemplate:
        self._ensure_loaded()
        return self._doc

    @property
    def jinja_env(self) -> SandboxedEnvironment:
        self._ensure_loaded()
        return self._jinja_env
