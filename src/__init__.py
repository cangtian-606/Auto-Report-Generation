"""Document Generator - Word 文档生成器"""

from .reader.xlsx import ExcelDataReader
from .reader.yaml import YamlDataReader
from .processing.mapper import DataMapper
from .render.generator import DocumentGenerator, generate
from .processing.schema import DataSchema, SchemaValidator
from .render.filters import FILTERS, filter_money, filter_percent, filter_num, filter_date
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
