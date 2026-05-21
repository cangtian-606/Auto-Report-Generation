"""Document Generator - Word 文档生成器"""

from .reader import ExcelDataReader
from .mapper import DataMapper
from .generator import DocumentGenerator, generate
from .schema import DataSchema, SchemaValidator
from .exceptions import (
    RenderError,
    DataReadError,
    TemplateError,
    ValidationError,
    SchemaError,
    TemplateSyntaxError,
)

TemplateRenderer = DocumentGenerator
render_single = generate

__all__ = [
    'ExcelDataReader',
    'DataMapper',
    'DocumentGenerator',
    'generate',
    'DataSchema',
    'SchemaValidator',
    'RenderError',
    'DataReadError',
    'TemplateError',
    'ValidationError',
    'SchemaError',
    'TemplateSyntaxError',
    'TemplateRenderer',
    'render_single',
]
