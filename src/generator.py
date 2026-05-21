#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word 文档生成器

将含变量占位符的 docxtpl 模板与 Excel 数据结合，生成填充好的 Word 文档。

Sheet类型：
- date.xxx: 键值对（字段编码只写字段名，自动组合为 date.xxx.字段名）
- form.xxx: 循环表格（列标题对应循环项属性名）

过滤器：
- | money    金额千分位（1000 → 1,000.00）
- | percent  百分比（0.25 → 25.00%）
- | num      千分位数字（1000 → 1,000）
- | date     日期格式化（YYYY-MM-DD → YYYY年MM月DD日）
- | default  空值默认值

使用方法：
    python -m src.generator --data data.xlsx --template template.docx --output output.docx
    python -m src.generator --batch data/ --template template.docx --output-dir output/
"""

import os
import sys
import json
import argparse
import logging
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Optional, Callable

import pandas as pd

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
from .exceptions import TemplateSyntaxError

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
        self.filters = self._register_filters()

    def _register_filters(self) -> Dict[str, Callable]:
        return {
            'money': self._filter_money,
            'percent': self._filter_percent,
            'num': self._filter_num,
            'date': self._filter_date,
            'default_dash': self._filter_default_dash,
            'default': self._filter_default,
            'int': self._filter_int,
            'str': self._filter_str,
        }

    @staticmethod
    def _filter_money(value: Any) -> str:
        if value is None or value == '':
            return ''
        try:
            num = float(value)
            if num == 0:
                return ''
            return f"{num:,.2f}"
        except (ValueError, TypeError):
            return str(value)

    @staticmethod
    def _filter_percent(value: Any) -> str:
        if value is None or value == '':
            return ''
        try:
            num = float(value)
            if num == 0:
                return ''
            return f"{num * 100:.2f}%"
        except (ValueError, TypeError):
            return str(value)

    @staticmethod
    def _filter_num(value: Any) -> str:
        if value is None or value == '':
            return ''
        try:
            num = float(value)
            if num == 0:
                return ''
            if num == int(num):
                return f"{int(num):,}"
            return f"{num:,.2f}"
        except (ValueError, TypeError):
            return str(value)

    @staticmethod
    def _filter_date(value: Any, fmt: str = '%Y年%m月%d日') -> str:
        if value is None or value == '':
            return ''
        for pattern in ('%Y-%m-%d', '%Y/%m/%d', '%Y年%m月%d日'):
            try:
                dt = datetime.strptime(str(value), pattern)
                return dt.strftime(fmt)
            except ValueError:
                continue
        return str(value)

    @staticmethod
    def _filter_default_dash(value: Any) -> str:
        if value is None or value == '' or value == 0:
            return '-'
        return str(value)

    @staticmethod
    def _filter_default(value: Any, default: str = '') -> str:
        if value is None or value == '':
            return default
        return str(value)

    @staticmethod
    def _filter_int(value: Any) -> str:
        if value is None or value == '':
            return ''
        try:
            return str(int(float(value)))
        except (ValueError, TypeError):
            return str(value)

    @staticmethod
    def _filter_str(value: Any) -> str:
        if value is None:
            return ''
        return str(value)

    def render(self, context: Dict[str, Any], output_path: str,
               strict: bool = False) -> bool:
        logger.info(f"加载模板: {self.template_path}")

        try:
            doc = DocxTemplate(self.template_path)
        except Exception as e:
            logger.error(f"加载模板失败: {e}")
            return False

        jinja_env = Environment()
        for name, func in self.filters.items():
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
            for name, func in self.filters.items():
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
    reader = ExcelDataReader(data_path)
    raw_data = reader.read_all()

    mapper = DataMapper(raw_data)
    context = mapper.build_context()

    logger.info(f"上下文统计: date={len(context['date'])} 个键值对, "
                f"form={len(context['form'])} 个表格 Sheet")

    if check_vars:
        generator = DocumentGenerator(template_path)
        undeclared = generator.get_undeclared_variables(context)
        if undeclared:
            logger.warning(f"发现 {len(undeclared)} 个未定义的模板变量:")
            for var in sorted(undeclared)[:10]:
                logger.warning(f"  - {var}")
            if len(undeclared) > 10:
                logger.warning(f"  ... 还有 {len(undeclared) - 10} 个")
            if strict:
                sys.exit(1)

    if check_syntax:
        generator = DocumentGenerator(template_path)
        generator.check_syntax(context, strict=strict, report_unused=report_unused)

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
                sys.exit(1)

    generator = DocumentGenerator(template_path)
    success = generator.render(context, output_path, strict=strict)
    return success


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Document Generator - Generate documents from templates and Excel data",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python -m src.generator --data data.xlsx --template template.docx --output output.docx
  python -m src.generator --batch data/ --template template.docx --output-dir output/
  python -m src.generator --data data.xlsx --template template.docx --output o.docx --strict
  python -m src.generator --data data.xlsx --template template.docx --output o.docx --schema schema.json --validate
  python -m src.generator --data data.xlsx --template template.docx --output o.docx --check-syntax --report-unused
        """,
    )
    parser.add_argument("--data", "-d", help="Excel data file path")
    parser.add_argument("--template", "-t", help="Template file path (.docx)")
    parser.add_argument("--output", "-o", help="Output file path (for single mode)")
    parser.add_argument("--batch", "-b", help="Data files directory (batch mode)")
    parser.add_argument("--output-dir", help="Output directory (for batch mode)")
    parser.add_argument("--strict", action="store_true",
                        help="Strict mode: error on undefined variables")
    parser.add_argument("--no-check", action="store_true",
                        help="Skip variable check")
    parser.add_argument("--schema", help="Schema file path (.json)")
    parser.add_argument("--validate", action="store_true",
                        help="Enable data validation")
    parser.add_argument("--strict-validate", action="store_true",
                        help="Strict validation: exit on validation errors")
    parser.add_argument("--check-syntax", action="store_true",
                        help="Check template syntax before rendering")
    parser.add_argument("--report-unused", action="store_true",
                        help="Report unused data fields")

    args = parser.parse_args()

    if args.batch:
        if not args.template:
            logger.error("Batch mode requires --template")
            sys.exit(1)
        if not os.path.exists(args.template):
            logger.error(f"Template file not found: {args.template}")
            sys.exit(1)
        if not os.path.isdir(args.batch):
            logger.error(f"Data directory not found: {args.batch}")
            sys.exit(1)

        output_dir = args.output_dir or "output"
        os.makedirs(output_dir, exist_ok=True)

        _render_batch(args.batch, args.template, output_dir,
                      strict=args.strict,
                      check_vars=not args.no_check,
                      validate=args.validate,
                      strict_validate=args.strict_validate,
                      schema_path=args.schema)

    else:
        if not args.data or not args.template or not args.output:
            parser.print_help()
            sys.exit(1)

        if not os.path.exists(args.data):
            logger.error(f"Data file not found: {args.data}")
            sys.exit(1)
        if not os.path.exists(args.template):
            logger.error(f"Template file not found: {args.template}")
            sys.exit(1)

        success = generate(
            args.data,
            args.template,
            args.output,
            strict=args.strict,
            check_vars=not args.no_check,
            validate=args.validate,
            strict_validate=args.strict_validate,
            schema_path=args.schema,
            check_syntax=args.check_syntax,
            report_unused=args.report_unused,
        )

        if success:
            logger.info(f"\n✓ Render complete: {args.output}")
            sys.exit(0)
        else:
            logger.error(f"\n✗ Render failed")
            sys.exit(1)


def _render_batch(data_dir: str, template_path: str, output_dir: str,
                  strict: bool = False,
                  check_vars: bool = True,
                  validate: bool = False,
                  strict_validate: bool = False,
                  schema_path: Optional[str] = None) -> List[tuple]:
    results: List[tuple] = []

    excel_files: List[Path] = []
    for ext in ['*.xlsx', '*.xls']:
        excel_files.extend(Path(data_dir).glob(ext))
        excel_files.extend(Path(data_dir).glob(f"**/{ext}"))

    excel_files = sorted(set(excel_files))

    if not excel_files:
        logger.error(f"在 {data_dir} 中未找到Excel文件")
        return results

    logger.info(f"找到 {len(excel_files)} 个数据文件")

    for data_file in excel_files:
        logger.info(f"\n{'='*60}")
        logger.info(f"处理: {data_file.name}")

        base_name = data_file.stem
        if base_name.startswith('data_'):
            base_name = base_name[5:]
        output_path = os.path.join(output_dir, f"{base_name}_output.docx")

        try:
            success = generate(
                str(data_file),
                template_path,
                output_path,
                strict=strict,
                check_vars=check_vars,
                validate=validate,
                strict_validate=strict_validate,
                schema_path=schema_path,
            )
            results.append((str(data_file), success))

            if success:
                logger.info(f"✓ 成功: {output_path}")
            else:
                logger.error(f"✗ 失败: {data_file.name}")
        except Exception as e:
            logger.error(f"✗ 异常: {data_file.name} - {e}")
            results.append((str(data_file), False))

    success_count = sum(1 for _, s in results if s)
    logger.info(f"\n{'='*60}")
    logger.info(f"批量处理完成: {success_count}/{len(results)} 成功")

    return results


if __name__ == "__main__":
    main()
