#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""CLI 入口"""

import os
import sys
import argparse
import logging
from pathlib import Path
from typing import List, Optional

from .render.generator import generate
from .exceptions import TemplateSyntaxError, ValidationError, DataReadError

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Document Generator - Generate documents from templates and Excel/YAML data",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python -m src --data data.xlsx --template template.docx --output output.docx
  python -m src --data data.yaml --template template.docx --output output.docx
  python -m src --batch data/ --template template.docx --output-dir output/
  python -m src --data data.xlsx --template template.docx --output o.docx --strict
  python -m src --data data.xlsx --template template.docx --output o.docx --schema schema.json --validate
  python -m src --data data.xlsx --template template.docx --output o.docx --check-syntax --report-unused
        """,
    )
    parser.add_argument("--data", "-d", help="Data file path (.xlsx .xls .yaml .yml)")
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

        try:
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
        except (TemplateSyntaxError, ValidationError, DataReadError) as e:
            logger.error(f"\n✗ {e}")
            sys.exit(1)

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

    data_files: List[Path] = []
    for ext in ['*.xlsx', '*.xls', '*.yaml', '*.yml']:
        data_files.extend(Path(data_dir).glob(ext))
        data_files.extend(Path(data_dir).glob(f"**/{ext}"))

    data_files = sorted(set(data_files))

    if not data_files:
        logger.error(f"在 {data_dir} 中未找到数据文件")
        return results

    logger.info(f"找到 {len(data_files)} 个数据文件")

    for data_file in data_files:
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
