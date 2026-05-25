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

    def test_tc_column_loop(self, temp_dir):
        """测试 {%tc for %} 列循环渲染 — 表头 + 多行数据"""
        from docx import Document
        tmpl_path = temp_dir / "tc_template.docx"
        out_path = temp_dir / "tc_output.docx"

        doc = Document()
        table = doc.add_table(rows=3, cols=5, style="Table Grid")

        # Row 0: 表头 — tc 循环动态生成月份列
        h = table.rows[0]
        h.cells[0].text = "产品"
        h.cells[1].paragraphs[0].text = "{%tc for month in months %}"
        h.cells[2].paragraphs[0].text = "{{ month }}"
        h.cells[3].paragraphs[0].text = "{%tc endfor %}"
        h.cells[4].text = "合计"

        # Row 1: 数据行 A
        a = table.rows[1]
        a.cells[0].text = "产品A"
        a.cells[1].paragraphs[0].text = "{%tc for month in months %}"
        a.cells[2].paragraphs[0].text = "{{ data_a[month] | money }}"
        a.cells[3].paragraphs[0].text = "{%tc endfor %}"
        a.cells[4].text = "{{ total_a | money }}"

        # Row 2: 数据行 B
        b = table.rows[2]
        b.cells[0].text = "产品B"
        b.cells[1].paragraphs[0].text = "{%tc for month in months %}"
        b.cells[2].paragraphs[0].text = "{{ data_b[month] | money }}"
        b.cells[3].paragraphs[0].text = "{%tc endfor %}"
        b.cells[4].text = "{{ total_b | money }}"

        doc.save(tmpl_path)

        gen = DocumentGenerator(str(tmpl_path))
        ctx = {
            "months": ["1月", "2月", "3月"],
            "data_a": {"1月": 100, "2月": 200, "3月": 150},
            "total_a": 450,
            "data_b": {"1月": 80, "2月": 120, "3月": 90},
            "total_b": 290,
        }

        result = gen.render(ctx, str(out_path))
        assert result is True
        assert out_path.exists()

        out_doc = Document(str(out_path))
        cells = [c.text.strip() for c in out_doc.tables[0].rows[0].cells]
        assert cells[0] == "产品"
        assert cells[1] == "1月"
        assert cells[2] == "2月"
        assert cells[3] == "3月"
        assert cells[4] == "合计"

        row_a = [c.text.strip() for c in out_doc.tables[0].rows[1].cells]
        assert row_a[0] == "产品A"
        assert row_a[1] == "100.00"
        assert row_a[4] == "450.00"

    def test_for_block_loop(self, temp_dir):
        """测试 {% for %} 块级循环 — 多段落按列表迭代"""
        from docx import Document
        tmpl_path = temp_dir / "for_block_template.docx"
        out_path = temp_dir / "for_block_output.docx"

        doc = Document()
        doc.add_paragraph("项目基本情况报告")
        doc.add_paragraph("")
        doc.add_paragraph("{% for i in 项目基本情况.项目公司 %}")
        doc.add_paragraph(
            "{{ i.项目名称 }}位于{{ i.项目地址 }}，"
            "装机容量{{ i.项目装机量 }}，"
            "采用\u201c{{ i.项目模式 }}\u201d运行方式。"
        )
        doc.add_paragraph("{% endfor %}")
        doc.save(tmpl_path)

        gen = DocumentGenerator(str(tmpl_path))
        ctx = {
            "项目基本情况": {
                "项目公司": [
                    {"项目名称": "重庆电站", "项目地址": "重庆", "项目装机量": "50MW", "项目模式": "全额上网"},
                    {"项目名称": "安徽风场", "项目地址": "六安", "项目装机量": "100MW", "项目模式": "自发自用"},
                ]
            }
        }

        result = gen.render(ctx, str(out_path))
        assert result is True

        out_doc = Document(str(out_path))
        texts = [p.text.strip() for p in out_doc.paragraphs if p.text.strip()]
        assert "项目基本情况报告" in texts[0]
        assert any("重庆电站" in t for t in texts)
        assert any("安徽风场" in t for t in texts)
        # {% for %} 标签行不应残留
        assert not any("{%" in t for t in texts)
        # 2 个公司 → 至少 3 个段落 (标题 + 2 段公司)
        assert len(texts) >= 3
