#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试文档生成器功能
"""

import os
import tempfile
import pytest
from pathlib import Path
from docx import Document

from src.render.generator import DocumentGenerator


@pytest.fixture
def temp_dir():
    """临时目录 fixture"""
    with tempfile.TemporaryDirectory() as tmp:
        yield Path(tmp)


@pytest.fixture
def simple_template(temp_dir):
    """创建一个简单的测试模板"""
    tmpl_path = temp_dir / "test_template.docx"
    doc = Document()
    doc.add_paragraph("公司名称: {{ global.company_name }}")
    doc.add_paragraph("成立日期: {{ global.found_date }}")
    doc.add_paragraph("项目列表:")
    # 简化为不使用 tr 循环，避免 docxtpl 特定语法问题
    doc.add_paragraph("示例项目: {{ projects[0].name if projects else '' }}")
    doc.save(tmpl_path)
    return tmpl_path


class TestDocumentGenerator:
    """测试 DocumentGenerator 类"""

    def test_find_unused_data(self, simple_template):
        """测试 _find_unused_data 功能"""
        gen = DocumentGenerator(str(simple_template))
        ctx = {
            "global": {
                "company_name": "Test",
                "found_date": "2020-01-01",
                "unused_field": "value"
            },
            "projects": [
                {"name": "Proj1", "year": 2020, "unused_proj_field": "v"}
            ],
            "unused_section": "value"
        }

        unused = gen._find_unused_data(ctx)

        # 检查未使用字段是否被识别
        assert "global.unused_field" in unused
        assert "unused_section" in unused

        # 已使用的字段不应该出现在未使用列表
        assert "global.company_name" not in unused
        assert "global.found_date" not in unused

    def test_render(self, simple_template, temp_dir):
        """测试基本渲染功能"""
        gen = DocumentGenerator(str(simple_template))
        output_path = temp_dir / "output.docx"
        ctx = {
            "global": {
                "company_name": "Test Company",
                "found_date": "2020-01-01"
            },
            "projects": [
                {"name": "Project 1", "year": 2020},
                {"name": "Project 2", "year": 2021}
            ]
        }

        result = gen.render(ctx, str(output_path))

        assert result is True
        assert output_path.exists()
        stat = output_path.stat()
        assert stat.st_size > 0

    def test_check_syntax_no_errors(self, simple_template):
        """测试 check_syntax 在无错误情况下"""
        gen = DocumentGenerator(str(simple_template))
        ctx = {
            "global": {
                "company_name": "Test",
                "found_date": "2020-01-01"
            },
            "projects": [{"name": "P1", "year": 2020}]
        }

        result = gen.check_syntax(ctx, report_unused=False)
        assert result is True

    def test_template_cache(self, simple_template):
        """测试模板缓存机制"""
        gen1 = DocumentGenerator(str(simple_template))
        doc1 = gen1._load_template()
        doc2 = gen1._load_template()

        # 同一对象引用
        assert doc1 is doc2

        # 不同实例相同路径，使用同一缓存
        gen2 = DocumentGenerator(str(simple_template))
        doc3 = gen2._load_template()
        assert doc3 is doc1
