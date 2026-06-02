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
        """报告 context 中未被模板引用的字段。

        重要约束：不下钻 list 内部。list 是否被使用以"列表整体路径"为单位判断，
        因为 list 元素是否被 {% for %} 块使用属于 Jinja2 循环变量作用域问题，
        纯路径匹配无法可靠判断（会引入大量误报）。
        """
        unused: List[str] = []
        template_vars = self._collect_template_vars()

        def walk(value: Any, path: str) -> None:
            if isinstance(value, dict):
                for key, val in value.items():
                    sub_path = f"{path}.{key}" if path else key
                    if isinstance(val, (dict, list)):
                        if sub_path not in template_vars:
                            unused.append(sub_path)
                        walk(val, sub_path)
                    else:
                        if sub_path not in template_vars:
                            unused.append(sub_path)
            elif isinstance(value, list):
                # list 不下钻：整列表已用则忽略，否则整列表报为未用
                if path not in template_vars:
                    unused.append(path)

        for key, value in context.items():
            if isinstance(value, (dict, list)):
                if key not in template_vars:
                    unused.append(key)
                walk(value, key)
            else:
                if key not in template_vars:
                    unused.append(key)
        return unused

    def _collect_template_vars(self) -> Set[str]:
        """收集模板中所有"被引用"的变量路径。

        收集两类：
        1. `{{ x.y }}` 表达式中的 x.y — 取首段路径用于"是否被引用"判断
        2. `{% for x in path %}` 中的 path — 标识被迭代的列表

        注意：正则只取"路径前缀"（不取完整表达式如 `x | filter`），
        以便与 context 中的纯路径对齐匹配。
        """
        self._ensure_loaded()
        var_pattern = re.compile(r"\{\{\s*([^\s|{}][^{}]*?)\s*(?:\||\}\})")
        for_pattern = re.compile(r"\{%\s*for\s+\w+\s+in\s+([^\s%]+)\s*%\}")
        template_vars: Set[str] = set()
        try:
            real_doc = self._doc.get_docx()
            for elem in real_doc.part._element.iter():
                for text in (elem.text, elem.tail):
                    if not text:
                        continue
                    for match in var_pattern.finditer(text):
                        var_path = match.group(1).strip()
                        if var_path:
                            template_vars.add(var_path)
                    for match in for_pattern.finditer(text):
                        list_path = match.group(1).strip()
                        if list_path:
                            template_vars.add(list_path)
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
