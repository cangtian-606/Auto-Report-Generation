#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word模板渲染器

将含变量占位符的docxtpl模板与Excel数据结合，生成填充好的Word文档。

使用前提：
- 模板中已包含Jinja2变量占位符，如 {{ g.company_name }}、{{ notes.remark }} 等
- 数据来自Excel文件，结构与模板变量对应

使用方法：
    python template_renderer.py --data data.xlsx template.docx output.docx
    python template_renderer.py --batch data/ output/ template.docx

作者：通用模板工具
版本：1.0.0
"""

import os
import re
import sys
import json
import argparse
import logging
from pathlib import Path
from typing import Dict, List, Any, Optional
from dataclasses import dataclass

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


# ============================================================
# 日志配置
# ============================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
logger = logging.getLogger(__name__)


# ============================================================
# 数据类
# ============================================================

@dataclass
class RenderStats:
    """渲染统计"""
    template_vars: int = 0
    data_fields: int = 0
    unmatched_vars: int = 0
    matched_vars: int = 0


# ============================================================
# Excel数据读取器
# ============================================================

class ExcelDataReader:
    """
    Excel数据读取器
    
    支持两种Sheet格式：
    1. 键值对格式：字段编码 | 值
    2. 表格格式：第一行为列标题，后续为数据行
    """
    
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.data: Dict[str, Any] = {}
    
    def read_all(self) -> Dict[str, Any]:
        """
        读取Excel所有数据
        
        Returns:
            完整数据字典
        """
        xl = pd.ExcelFile(self.file_path, engine='openpyxl')
        logger.info(f"读取Excel: {self.file_path}")
        logger.info(f"工作表: {xl.sheet_names}")
        
        for sheet_name in xl.sheet_names:
            data = self._read_sheet(xl, sheet_name)
            self.data[sheet_name] = data
        
        return self.data
    
    def _read_sheet(self, xl, sheet_name: str) -> Any:
        """
        根据Sheet名称或内容识别类型并读取
        
        Args:
            xl: ExcelFile对象
            sheet_name: 工作表名称
            
        Returns:
            读取的数据（dict或DataFrame）
        """
        # 读取前几行判断类型
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None, nrows=5)
        
        if df.empty:
            return {}
        
        # 键值对格式识别：第一列是"字段编码"、"key"、"名称"等关键词
        first_cell = str(df.iloc[0, 0]).strip().lower() if pd.notna(df.iloc[0, 0]) else ""
        key_names = ["字段编码", "key", "名称", "变量", "field", "name", "code"]
        
        if first_cell in key_names:
            return self._read_key_value_sheet(xl, sheet_name)
        
        # 表格格式：第一行是列标题
        return self._read_table_sheet(xl, sheet_name)
    
    def _read_key_value_sheet(self, xl, sheet_name: str) -> Dict[str, Any]:
        """
        读取键值对格式的Sheet
        
        格式：
        | 字段编码 | 值 |
        | g.company_name | XX公司 |
        | g.date | 2025-01-15 |
        """
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
        result = {}
        
        for _, row in df.iterrows():
            key = row.iloc[0] if pd.notna(row.iloc[0]) else None
            value = row.iloc[1] if len(row) > 1 and pd.notna(row.iloc[1]) else None
            
            if key is not None:
                key_str = str(key).strip()
                # 跳过标题行
                if key_str.lower() in ["字段编码", "key", "名称", "变量", "field", "name", "code", 
                                        "值", "value", "content", "内容"]:
                    continue
                if not key_str:
                    continue
                result[key_str] = value
        
        logger.info(f"  [{sheet_name}] 键值对: {len(result)} 条")
        return result
    
    def _read_table_sheet(self, xl, sheet_name: str) -> pd.DataFrame:
        """
        读取表格格式的Sheet
        
        格式：
        | 名称 | 数量 | 金额 |
        | 产品A | 10 | 1000 |
        """
        df = pd.read_excel(xl, sheet_name=sheet_name, engine='openpyxl')
        
        # 清理列名
        df.columns = [str(col).strip() for col in df.columns]
        
        logger.info(f"  [{sheet_name}] 表格: {len(df)} 行 x {len(df.columns)} 列")
        return df


# ============================================================
# 数据映射器
# ============================================================

class DataMapper:
    """
    数据映射器
    
    将Excel读取的原始数据转换为docxtpl模板所需的context结构。
    """
    
    def __init__(self, raw_data: Dict[str, Any]):
        self.raw_data = raw_data
    
    def build_context(self) -> Dict[str, Any]:
        """
        构建模板渲染上下文

        支持两种 Sheet 类型：
        - date.xxx: 键值对 → context['date']
        - form.xxx: 表格数据 → context['form']

        字段名支持点号嵌套，如 `date.全局信息.公司名` → context['date']['全局信息']['公司名']

        Returns:
            docxtpl模板渲染上下文
        """
        context: Dict[str, Any] = {
            'date': {},
            'form': {},
        }

        for sheet_name, data in self.raw_data.items():
            # 键值对：date.xxx Sheet → date.xxx
            if isinstance(data, dict):
                for key, value in data.items():
                    if key.startswith('date.'):
                        key_suffix = key[5:]
                        self._set_nested(context['date'], key_suffix, self._convert_value(value))
                    else:
                        context['date'][key] = self._convert_value(value)

            # 表格数据：form.xxx Sheet → form.xxx
            elif isinstance(data, pd.DataFrame):
                if sheet_name.startswith('form.'):
                    context['form'][sheet_name[5:]] = data.to_dict('records')
                else:
                    context['form'][sheet_name] = data.to_dict('records')

        return context

    def _set_nested(self, d: Dict[str, Any], path: str, value: Any) -> None:
        """将键值按点号路径设置到嵌套字典中"""
        parts = path.split('.')
        current = d
        for part in parts[:-1]:
            if part not in current:
                current[part] = {}
            current = current[part]
        current[parts[-1]] = value
    
    def _convert_value(self, value: Any) -> Any:
        """转换值类型"""
        if value is None:
            return ''
        if isinstance(value, (int, float)):
            if value == 0:
                return ''
            return value
        return value
    
    def _convert_bool(self, value: Any) -> bool:
        """Convert boolean value. Only accepts TRUE/FALSE (case-insensitive)."""
        if isinstance(value, bool):
            return value
        if isinstance(value, str):
            return value.strip().upper() == 'TRUE'
        if isinstance(value, (int, float)):
            return bool(value)
        return False


# ============================================================
# 模板渲染器
# ============================================================

class TemplateRenderer:
    """
    docxtpl模板渲染器
    
    负责加载模板、注册过滤器、执行渲染、保存输出。
    """
    
    def __init__(self, template_path: str):
        self.template_path = template_path
        self.filters = self._register_filters()
    
    def _register_filters(self) -> Dict[str, callable]:
        """注册Jinja2自定义过滤器"""
        return {
            'money': self._filter_money,
            'percent': self._filter_percent,
            'default_dash': self._filter_default_dash,
            'default': self._filter_default,
            'int': self._filter_int,
            'str': self._filter_str,
        }
    
    @staticmethod
    def _filter_money(value: Any) -> str:
        """金额千分位格式化"""
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
        """百分比格式化"""
        if value is None or value == '':
            return ''
        try:
            return f"{float(value):.2f}%"
        except (ValueError, TypeError):
            return str(value)
    
    @staticmethod
    def _filter_default_dash(value: Any) -> str:
        """空值显示为横线"""
        if value is None or value == '' or value == 0:
            return '-'
        return str(value)
    
    @staticmethod
    def _filter_default(value: Any, default: str = '') -> str:
        """空值显示为指定默认值"""
        if value is None or value == '':
            return default
        return str(value)
    
    @staticmethod
    def _filter_int(value: Any) -> str:
        """转换为整数"""
        if value is None or value == '':
            return ''
        try:
            return str(int(float(value)))
        except (ValueError, TypeError):
            return str(value)
    
    @staticmethod
    def _filter_str(value: Any) -> str:
        """转换为字符串"""
        if value is None:
            return ''
        return str(value)
    
    def render(self, context: Dict[str, Any], output_path: str, 
                strict: bool = False) -> bool:
        """
        渲染模板并保存
        
        Args:
            context: 渲染上下文
            output_path: 输出文件路径
            strict: 是否启用严格模式（未定义变量报错）
            
        Returns:
            是否成功
        """
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
        """
        获取模板中未在context中定义的变量
        
        用于调试和检查数据完整性。
        """
        try:
            doc = DocxTemplate(self.template_path)
            return doc.get_undeclared_template_variables(context=context)
        except Exception as e:
            logger.error(f"获取未定义变量失败: {e}")
            return set()


# ============================================================
# 验证器
# ============================================================

class ContextValidator:
    """上下文验证器"""
    
    def __init__(self):
        self.stats = RenderStats()
    
    def validate(self, context: Dict[str, Any], undeclared_vars: set) -> List[str]:
        """
        验证上下文和变量匹配情况
        
        Returns:
            问题列表
        """
        issues = []
        
        # 统计上下文字段
        total_fields = 0
        for domain in context.values():
            if isinstance(domain, dict):
                total_fields += len(domain)
            elif isinstance(domain, list):
                total_fields += len(domain)
        self.stats.data_fields = total_fields
        
        # 统计未匹配变量
        self.stats.unmatched_vars = len(undeclared_vars)
        
        if undeclared_vars:
            logger.warning(f"发现 {len(undeclared_vars)} 个未定义的模板变量:")
            for var in sorted(undeclared_vars)[:10]:
                logger.warning(f"  - {var}")
            if len(undeclared_vars) > 10:
                logger.warning(f"  ... 还有 {len(undeclared_vars) - 10} 个")
            
            issues.append(f"{len(undeclared_vars)} 个变量未在数据中找到")
        
        return issues


# ============================================================
# CLI入口
# ============================================================

def render_single(data_path: str, template_path: str, output_path: str,
                  strict: bool = False, check_vars: bool = True) -> bool:
    """
    渲染单个文档
    
    Args:
        data_path: Excel数据文件路径
        template_path: docxtpl模板路径
        output_path: 输出文件路径
        strict: 是否严格模式
        check_vars: 是否检查未定义变量
        
    Returns:
        是否成功
    """
    # 1. 读取数据
    reader = ExcelDataReader(data_path)
    raw_data = reader.read_all()
    
    # 2. 构建上下文
    mapper = DataMapper(raw_data)
    context = mapper.build_context()
    
    logger.info(f"上下文统计: date={len(context['date'])} 个键值对, "
                f"form={len(context['form'])} 个表格 Sheet")
    
    # 3. 检查未定义变量
    if check_vars:
        renderer = TemplateRenderer(template_path)
        undeclared = renderer.get_undeclared_variables(context)
        validator = ContextValidator()
        issues = validator.validate(context, undeclared)
        
        if issues:
            logger.warning("变量检查发现以下问题:")
            for issue in issues:
                logger.warning(f"  - {issue}")
    
    # 4. 渲染
    renderer = TemplateRenderer(template_path)
    success = renderer.render(context, output_path, strict=strict)
    
    return success


def render_batch(data_dir: str, template_path: str, output_dir: str,
                 strict: bool = False) -> List[tuple]:
    """
    批量渲染文档
    
    Args:
        data_dir: 数据文件目录
        template_path: 模板文件路径
        output_dir: 输出目录
        strict: 是否严格模式
        
    Returns:
        结果列表 [(文件名, 成功标志), ...]
    """
    results = []
    
    # 查找所有Excel文件
    excel_files = []
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
        
        # 生成输出文件名
        base_name = data_file.stem
        # 去掉前缀如 "data_"、"data_公司A_"
        if base_name.startswith('data_'):
            base_name = base_name[5:]
        output_path = os.path.join(output_dir, f"{base_name}_output.docx")
        
        try:
            success = render_single(
                str(data_file), 
                template_path, 
                output_path,
                strict=strict,
                check_vars=False
            )
            results.append((str(data_file), success))
            
            if success:
                logger.info(f"✓ 成功: {output_path}")
            else:
                logger.error(f"✗ 失败: {data_file.name}")
        except Exception as e:
            logger.error(f"✗ 异常: {data_file.name} - {e}")
            results.append((str(data_file), False))
    
    # 汇总
    success_count = sum(1 for _, s in results if s)
    logger.info(f"\n{'='*60}")
    logger.info(f"批量处理完成: {success_count}/{len(results)} 成功")
    
    return results


def main():
    parser = argparse.ArgumentParser(
        description="Word Template Renderer - Generate documents from templates and Excel data",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Single file mode
  python -m src.renderer --data data.xlsx --template template.docx --output output.docx

  # Batch mode
  python -m src.renderer --batch data/ --template template.docx --output-dir output/

  # Strict mode (error on undefined variables)
  python -m src.renderer --strict --data data.xlsx --template template.docx --output output.docx

Excel data file format:
  Key-value format:
    | field_code | value |
    | g.company_name | XX Company |
    | g.date | 2025-01-15 |

  Table format (for loops):
    | name | quantity | amount |
    | Product A | 10 | 1000 |
    | Product B | 5 | 500 |
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
        
        render_batch(args.batch, args.template, output_dir, strict=args.strict)
    
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
        
        success = render_single(
            args.data,
            args.template,
            args.output,
            strict=args.strict,
            check_vars=not args.no_check
        )
        
        if success:
            logger.info(f"\n✓ Render complete: {args.output}")
            sys.exit(0)
        else:
            logger.error(f"\n✗ Render failed")
            sys.exit(1)


if __name__ == "__main__":
    main()
