#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 文档生成器

将含变量占位符的 docxtpl 模板与 Excel 数据结合，生成填充好的 Word 文档。

Sheet类型：
- date.xxx: 键值对（字段编码只写字段名，自动组合为 date.xxx.字段名）
- form.xxx: 循环表格（列标题对应循环项属性名）
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

from .reader import ExcelDataReader
from .mapper import DataMapper
from .filters import FILTERS
from .exceptions import TemplateSyntaxError, ValidationError

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)


class DocumentGenerator:
    """docxtpl 文档生成器"""

    def __init__(self, template_path: str) -> None:
        self.template_path = template_path

    def render(self, context: Dict[str, Any], output_path: str,
               strict: bool = False) -> bool:
        logger.info(f"加载模板: {self.template_path}")

        try:
            doc = DocxTemplate(self.template_path)
        except Exception as e:
            logger.error(f"加载模板失败: {e}")
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
            doc = DocxTemplate(self.template_path)
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
        for domain_key, domain_value in context.items():
            if isinstance(domain_value, dict):
                for field_key in domain_value.keys():
                    path = f"{domain_key}.{field_key}"
                    if isinstance(domain_value[field_key], dict):
                        for sub_key in domain_value[field_key].keys():
                            unused.append(f"{path}.{sub_key}")
                    else:
                        unused.append(path)
        return unused


def generate(data_path: str, template_path: str, output_path: str,
             strict: bool = False, check_vars: bool = True,
             validate: bool = False, strict_validate: bool = False,
             schema_path: Optional[str] = None,
             check_syntax: bool = False,
             report_unused: bool = False) -> bool:
    reader = ExcelDataReader(data_path, strict=strict)
    raw_data = reader.read_all()

    mapper = DataMapper(raw_data)
    context = mapper.build_context()

    logger.info(f"上下文统计: date={len(context['date'])} 个键值对, "
                f"form={len(context['form'])} 个表格 Sheet")

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

    if validate:
        from .schema import SchemaValidator
        validator = SchemaValidator()
        if schema_path:
            validator.load_from_file(schema_path)
        errors = validator.validate(raw_data)
        if errors:
            for e in errors:
                logger.error(f"[数据验证] {e}")
            if strict_validate:
                raise ValidationError(f"数据验证失败: {len(errors)} 个错误")

    success = gen.render(context, output_path, strict=strict)
    return success
