"""pytest 报告模板渲染测试"""

import os
import sys
from pathlib import Path

import pytest
from docx import Document

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))

TEMPLATE_SRC = PROJECT_DIR / "templates" / "报告测试.docx"
TEMPLATE_OUT = PROJECT_DIR / "templates" / "报告模板.docx"
DATA_FILE = PROJECT_DIR / "data" / "报告数据.xlsx"
OUTPUT_FILE = PROJECT_DIR / "output" / "投资项目概况_测试输出.docx"


@pytest.fixture(scope="module")
def setup_files():
    """创建测试文件和执行渲染，返回输出文档"""
    os.makedirs(PROJECT_DIR / "output", exist_ok=True)

    _setup_template()
    _setup_data()

    from src.generator import generate
    generate(
        str(DATA_FILE),
        str(TEMPLATE_OUT),
        str(OUTPUT_FILE),
        strict=False,
        check_vars=False,
    )

    return OUTPUT_FILE


def _setup_template():
    """将原始报告测试.docx转换为含Jinja2变量的docxtpl模板"""
    doc = Document(str(TEMPLATE_SRC))

    basic_fields = [
        ("date.全局信息.公司名", None),
        ("date.基本情况.信用代码", None),
        ("date.基本情况.注册地址", None),
        ("date.基本情况.法定代表人", None),
        ("date.基本情况.公司类型", None),
        ("date.基本情况.注册资本", "num"),
        ("date.基本情况.经营范围", None),
        ("date.基本情况.成立日期", None),
        ("date.基本情况.经营期限", None),
    ]
    table0 = doc.tables[0]
    for ri, (field, flt) in enumerate(basic_fields):
        cell = table0.rows[ri].cells[1]
        para = cell.paragraphs[0]
        para.clear()
        if flt:
            para.add_run("{{ " + field + " | " + flt + " }}")
        else:
            para.add_run("{{ " + field + " }}")

    table1 = doc.tables[1]
    while len(table1.rows) > 1:
        table1.rows[-1]._tr.getparent().remove(table1.rows[-1]._tr)
    for_row1 = table1.add_row()
    for_row1.cells[0].paragraphs[0].add_run("{%tr for i in form.股东出资 %}")
    for_row1.cells[1].paragraphs[0].clear()
    for_row1.cells[2].paragraphs[0].clear()
    for_row1.cells[3].paragraphs[0].clear()
    for_row1.cells[4].paragraphs[0].clear()
    data_row1 = table1.add_row()
    data_row1.cells[0].paragraphs[0].add_run("{{ i.股东 }}")
    data_row1.cells[1].paragraphs[0].add_run("{{ i.认缴金额 | money }}")
    data_row1.cells[2].paragraphs[0].add_run("{{ i.认缴比例 | percent }}")
    data_row1.cells[3].paragraphs[0].add_run("{{ i.实缴金额 | money }}")
    data_row1.cells[4].paragraphs[0].add_run("{{ i.实缴比例 | percent }}")
    endfor_row1 = table1.add_row()
    endfor_row1.cells[0].paragraphs[0].add_run("{%tr endfor %}")

    table2 = doc.tables[2]
    table3 = doc.tables[3]
    if_para = doc.add_paragraph("{%p if date.基本情况.有职工信息 %}")
    end_para = doc.add_paragraph("{%p endif %}")
    body = doc.element.body
    body.insert(list(body).index(table2._element), if_para._p)
    body.insert(list(body).index(table3._element), end_para._p)

    while len(table2.rows) > 1:
        table2.rows[1]._tr.getparent().remove(table2.rows[1]._tr)
    for_row2 = table2.add_row()
    for_row2.cells[0].paragraphs[0].add_run("{%tr for i in form.员工信息 %}")
    data_row2 = table2.add_row()
    data_row2.cells[0].paragraphs[0].add_run("{{ i.序号 }}")
    data_row2.cells[1].paragraphs[0].add_run("{{ i.职务 }}")
    endfor_row2 = table2.add_row()
    endfor_row2.cells[0].paragraphs[0].add_run("{%tr endfor %}")

    table3 = doc.tables[3]
    while len(table3.rows) > 1:
        table3.rows[1]._tr.getparent().remove(table3.rows[1]._tr)
    for_row3 = table3.add_row()
    for_row3.cells[0].paragraphs[0].add_run("{%tr for i in form.投资明细 %}")
    data_row3 = table3.add_row()
    data_row3.cells[0].paragraphs[0].add_run("{{ i.项目 }}")
    data_row3.cells[1].paragraphs[0].add_run("{{ i.金额 | money }}")
    data_row3.cells[2].paragraphs[0].add_run("{{ i.持股比例 | percent }}")
    endfor_row3 = table3.add_row()
    endfor_row3.cells[0].paragraphs[0].add_run("{%tr endfor %}")

    doc.save(str(TEMPLATE_OUT))


def _setup_data():
    """创建测试数据Excel"""
    import openpyxl

    wb = openpyxl.Workbook()

    ws_global = wb.active
    ws_global.title = "date.全局信息"
    ws_global.append(["字段编码", "值"])
    ws_global.append(["公司名", "深圳市创新科技有限公司"])
    ws_global.column_dimensions["A"].width = 25
    ws_global.column_dimensions["B"].width = 50

    ws_basic = wb.create_sheet("date.基本情况")
    ws_basic.append(["字段编码", "值"])
    basic_data = [
        ["信用代码", "91440300MA5XXXXX00"],
        ["注册地址", "深圳市南山区粤海街道科技园南路88号"],
        ["法定代表人", "张伟"],
        ["公司类型", "有限责任公司"],
        ["注册资本", 1000.00],
        ["经营范围", "软件开发、技术咨询、技术服务"],
        ["成立日期", "2018-06-15"],
        ["经营期限", "2018-06-15 至 长期"],
        ["有职工信息", "FALSE"],
    ]
    for row_data in basic_data:
        ws_basic.append(row_data)
    ws_basic.column_dimensions["A"].width = 25
    ws_basic.column_dimensions["B"].width = 50

    ws_sh = wb.create_sheet("form.股东出资")
    ws_sh.append(["股东", "认缴金额", "认缴比例", "实缴金额", "实缴比例"])
    shareholders = [
        ["张伟", 600.00, 0.60, 600.00, 0.60],
        ["李芳", 250.00, 0.25, 250.00, 0.25],
        ["王强", 150.00, 0.15, 150.00, 0.15],
    ]
    for row_data in shareholders:
        ws_sh.append(row_data)
    ws_sh.column_dimensions["A"].width = 20
    for col in ["B", "C", "D", "E"]:
        ws_sh.column_dimensions[col].width = 18

    ws_info = wb.create_sheet("form.员工信息")
    ws_info.append(["序号", "职务"])
    for row_data in [[1, "张伟-执行董事"], [2, "李芳-监事"], [3, "王强-股东"]]:
        ws_info.append(row_data)

    ws_invest = wb.create_sheet("form.投资明细")
    ws_invest.append(["项目", "金额", "持股比例"])
    for row_data in [["项目A", 500.00, 0.50], ["项目B", 300.00, 0.30],
                     ["项目C", 200.00, 0.20], ["合计", 1000.00, 1.00]]:
        ws_invest.append(row_data)

    wb.save(str(DATA_FILE))


def _get_all_text(path):
    doc = Document(str(path))
    all_text = " ".join(p.text for p in doc.paragraphs)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                all_text += " " + c.text
    return all_text


class TestBasicFields:
    def test_company_name_replaced(self, setup_files):
        text = _get_all_text(setup_files)
        assert "深圳市创新科技有限公司" in text

    def test_credit_code_replaced(self, setup_files):
        text = _get_all_text(setup_files)
        assert "91440300MA5XXXXX00" in text

    def test_legal_person_replaced(self, setup_files):
        text = _get_all_text(setup_files)
        assert "张伟" in text

    def test_registered_capital_replaced(self, setup_files):
        text = _get_all_text(setup_files)
        assert "1,000" in text

    def test_establishment_date_replaced(self, setup_files):
        text = _get_all_text(setup_files)
        assert "2018-06-15" in text


class TestMoneyFilter:
    def test_shareholder_li_in_table(self, setup_files):
        text = _get_all_text(setup_files)
        assert "李芳" in text

    def test_subscription_amount_250(self, setup_files):
        text = _get_all_text(setup_files)
        assert "250" in text

    def test_money_format_subscription_250(self, setup_files):
        text = _get_all_text(setup_files)
        assert "250.00" in text

    def test_money_format_paid_600(self, setup_files):
        text = _get_all_text(setup_files)
        assert "600.00" in text


class TestPercentFilter:
    def test_percent_format_25(self, setup_files):
        text = _get_all_text(setup_files)
        assert "25.00%" in text

    def test_percent_format_60(self, setup_files):
        text = _get_all_text(setup_files)
        assert "60.00%" in text


class TestInvestmentTable:
    def test_project_a_in_table(self, setup_files):
        text = _get_all_text(setup_files)
        assert "项目A" in text

    def test_total_in_table(self, setup_files):
        text = _get_all_text(setup_files)
        assert "合计" in text

    def test_investment_amount_500_format(self, setup_files):
        text = _get_all_text(setup_files)
        assert "500.00" in text

    def test_share_ratio_50_format(self, setup_files):
        text = _get_all_text(setup_files)
        assert "50.00%" in text


class TestConditionalRendering:
    def test_employee_table_hidden_when_false(self, setup_files):
        text = _get_all_text(setup_files)
        assert "张伟-执行董事" not in text

    def test_no_jinja_tags_remaining(self, setup_files):
        text = _get_all_text(setup_files)
        assert "{{" not in text
        assert "{%" not in text
