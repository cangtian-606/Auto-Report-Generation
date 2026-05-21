#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""自定义异常类"""


class RenderError(Exception):
    """渲染错误基类"""
    pass


class DataReadError(RenderError):
    """数据读取错误"""
    pass


class TemplateError(RenderError):
    """模板错误"""
    pass


class ValidationError(RenderError):
    """数据验证错误"""
    pass


class SchemaError(ValidationError):
    """Schema 验证错误"""
    pass


class TemplateSyntaxError(TemplateError):
    """模板语法错误"""
    pass
