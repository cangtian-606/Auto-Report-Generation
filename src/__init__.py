"""Document Generator - Word 文档生成器"""

from .reader import ExcelDataReader
from .mapper import DataMapper
from .yaml_reader import YamlDataReader
from .generator import DocumentGenerator, generate
from .schema import DataSchema, SchemaValidator
from .filters import FILTERS, filter_money, filter_percent, filter_num, filter_date
from .exceptions import (
    DataReadError,
    TemplateError,
    ValidationError,
    TemplateSyntaxError,
)

TemplateRenderer = DocumentGenerator
render_single = generate

__all__ = [
    'ExcelDataReader',
    'DataMapper',
    'YamlDataReader',
    'DocumentGenerator',
    'generate',
    'DataSchema',
    'SchemaValidator',
    'FILTERS',
    'filter_money',
    'filter_percent',
    'filter_num',
    'filter_date',
    'TemplateError',
    'DataReadError',
    'ValidationError',
    'TemplateSyntaxError',
    'TemplateRenderer',
    'render_single',
]
