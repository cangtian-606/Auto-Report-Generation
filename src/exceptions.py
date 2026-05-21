#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""自定义异常类"""


class DataReadError(Exception):
    """数据读取错误"""
    pass


class TemplateError(Exception):
    """模板错误"""
    pass


class ValidationError(Exception):
    """数据验证错误"""
    pass


class TemplateSyntaxError(TemplateError):
    """模板语法错误"""
    pass
