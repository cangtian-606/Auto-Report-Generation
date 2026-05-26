#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 文档生成器

将含变量占位符的 docxtpl 模板与数据结合，生成填充好的 Word 文档。
"""

import os
import logging
from typing import Dict, List, Any, Optional

try:
    from docxtpl import DocxTemplate
except ImportError:
    raise ImportError("请先安装 docxtpl: pip install docxtpl") from None

try:
    from jinja2.sandbox import SandboxedEnvironment
except ImportError:
    raise ImportError("请先安装 jinja2: pip install jinja2") from None

from .filters import FILTERS
from ..exceptions import TemplateSyntaxError

logger = logging.getLogger(__name__)


def _extract_error_info(exc: Exception) -> str:
    """从 Jinja2 异常中提取用户可读的错误信息。"""
    msg = str(exc)
    for line in msg.splitlines():
        line = line.strip()
        if line and "is undefined" in line:
            return line.split("is undefined")[0].strip().rstrip("'").lstrip("'")
    return msg.splitlines()[0].strip() if msg else "未知错误"


class DocumentGenerator:
    """docxtpl 文档生成器"""

    _template_cache: Dict[str, DocxTemplate] = {}

    def __init__(self, template_path: str) -> None:
        self.template_path = template_path

    def _load_template(self) -> DocxTemplate:
        if self.template_path not in self._template_cache:
            logger.debug("加载模板: %s", self.template_path)
            try:
                self._template_cache[self.template_path] = DocxTemplate(self.template_path)
            except Exception:
                raise
        return self._template_cache[self.template_path]

    def render(self, context: Dict[str, Any], output_path: str,
               strict: bool = False) -> bool:
        try:
            doc = self._load_template()
        except Exception:
            logger.exception("[模板加载] 渲染前置失败")
            return False

        jinja_env = SandboxedEnvironment()
        for name, func in FILTERS.items():
            jinja_env.filters[name] = func

        logger.debug("执行渲染...")
        try:
            doc.render(context, jinja_env=jinja_env)
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

    def get_undeclared_variables(self, context: Dict[str, Any]) -> Optional[set]:
        try:
            doc = self._load_template()
            jinja_env = SandboxedEnvironment()
            for name, func in FILTERS.items():
                jinja_env.filters[name] = func
            return doc.get_undeclared_template_variables(context=context, jinja_env=jinja_env)
        except Exception:
            logger.exception("获取未定义变量失败")
            return None

    def check_syntax(self, context: Dict[str, Any],
                     strict: bool = False,
                     report_unused: bool = False) -> bool:
        undeclared = self.get_undeclared_variables(context)
        has_issue = False

        if undeclared is not None:
            if undeclared:
                has_issue = True
                logger.warning(f"模板中有 {len(undeclared)} 个变量在数据中未找到:")
                for var in sorted(undeclared)[:10]:
                    logger.warning(f"  - {var}")
                if len(undeclared) > 10:
                    logger.warning(f"  ... 还有 {len(undeclared) - 10} 个")
                if strict:
                    raise TemplateSyntaxError(f"未声明变量: {undeclared}")
        else:
            has_issue = True

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
        try:
            doc = self._load_template()
            real_doc = doc.get_docx()
        except Exception:
            return unused

        import re

        texts = []
        for elem in real_doc.part._element.iter():
            if elem.text:
                texts.append(elem.text)
            if elem.tail:
                texts.append(elem.tail)

        var_pattern = re.compile(r'\{\{\s*(.+?)\s*(?:\||\}\})')
        template_vars: set = set()
        for text in texts:
            for match in var_pattern.finditer(text):
                var_path = match.group(1).strip()
                if var_path:
                    template_vars.add(var_path)

        def traverse(value: Any, path_prefix: str = ''):
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
