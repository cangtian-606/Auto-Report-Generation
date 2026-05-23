#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""CLI 入口"""

import os
import sys
import argparse
import logging
from pathlib import Path
from typing import List, Optional

from .orchestrator import generate
from .exceptions import TemplateError, ValidationError, DataReadError
from .logging_config import configure_logging
from .path_guard import validate_path

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
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="Enable verbose (DEBUG) terminal output")
    parser.add_argument("--quiet", "-q", action="store_true",
                        help="Suppress non-error terminal output (WARNING only)")
    parser.add_argument("--log-file", help="Write full debug log to file")

    args = parser.parse_args()

    # 日志配置: 用户错误→warning(stderr), 程序异常→error/exception(stderr), 进度→info(stdout)
    terminal_level = logging.INFO
    if args.verbose:
        terminal_level = logging.DEBUG
    if args.quiet:
        terminal_level = logging.WARNING

    configure_logging(
        terminal_level=terminal_level,
        log_file=args.log_file,
    )

    project_root = Path(__file__).parent.parent.absolute()
    allowed_data_dirs = [str(project_root / "data"), str(project_root)]
    allowed_template_dirs = [str(project_root / "templates"), str(project_root)]
    allowed_output_dirs = [str(project_root / "output"), str(project_root)]

    if args.batch:
        if not args.template:
            logger.warning("Batch mode requires --template")
            sys.exit(1)

        try:
            args.template = validate_path(args.template, allowed_template_dirs, must_exist=True)
            args.batch = validate_path(args.batch, allowed_data_dirs, must_exist=True)
            if args.schema:
                args.schema = validate_path(args.schema, allowed_data_dirs, must_exist=True)
        except PermissionError as e:
            logger.warning("路径校验失败: %s", e)
            sys.exit(1)
        except FileNotFoundError as e:
            logger.warning("文件不存在: %s", e)
            sys.exit(1)

        if not os.path.isdir(args.batch):
            logger.warning("不是目录: %s", args.batch)
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

        try:
            args.data = validate_path(args.data, allowed_data_dirs, must_exist=True)
            args.template = validate_path(args.template, allowed_template_dirs, must_exist=True)
            args.output = validate_path(args.output, allowed_output_dirs)
            if args.schema:
                args.schema = validate_path(args.schema, allowed_data_dirs, must_exist=True)
        except PermissionError as e:
            logger.warning("路径校验失败: %s", e)
            sys.exit(1)
        except FileNotFoundError as e:
            logger.warning("文件不存在: %s", e)
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
        except (TemplateError, ValidationError, DataReadError) as e:
            logger.error("\n✗ %s", e)
            sys.exit(1)

        if success:
            logger.info("\n✓ Render complete: %s", args.output)
            sys.exit(0)
        else:
            logger.error("\n✗ Render failed")
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
        logger.warning("在 %s 中未找到数据文件", data_dir)
        return results

    logger.info("找到 %d 个数据文件", len(data_files))

    for data_file in data_files:
        logger.debug("=" * 60)
        logger.info("处理: %s", data_file.name)

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
                logger.info("✓ 成功: %s", output_path)
            else:
                logger.error("✗ 失败: %s", data_file.name)
        except KeyboardInterrupt:
            logger.warning("\n用户中断，停止批量处理（已处理 %d 个文件）", len(results))
            raise
        except Exception:
            logger.exception("✗ 异常: %s", data_file.name)
            results.append((str(data_file), False))

    success_count = sum(1 for _, s in results if s)
    logger.debug("=" * 60)
    logger.info("批量处理完成: %d/%d 成功", success_count, len(results))

    return results


if __name__ == "__main__":
    main()
