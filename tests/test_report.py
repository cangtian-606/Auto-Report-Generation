import os
import sys
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))

from docx import Document

TEMPLATE_SRC = PROJECT_DIR / "templates" / "报告测试.docx"
TEMPLATE_OUT = PROJECT_DIR / "templates" / "报告模板.docx"
DATA_FILE = PROJECT_DIR / "data" / "报告数据.xlsx"
OUTPUT_FILE = PROJECT_DIR / "output" / "投资项目概况_测试输出.docx"
EXIT_SUCCESS = 0
EXIT_FAILURE = 1


def setup_template():
    """将原始报告测试.docx转换为含Jinja2变量的docxtpl模板"""
    from docx import Document

    doc = Document(str(TEMPLATE_SRC))

    # === 基本信息表 (Table 0): 9 rows x 2 cols ===
    basic_fields = [
        "g.company_name",
        "g.credit_code",
        "g.registered_address",
        "g.legal_representative",
        "g.company_type",
        "g.registered_capital",
        "g.business_scope",
        "g.established_date",
        "g.operating_period",
    ]
    table0 = doc.tables[0]
    for ri, field in enumerate(basic_fields):
        cell = table0.rows[ri].cells[1]
        para = cell.paragraphs[0]
        para.clear()
        para.add_run("{{ " + field + " }}")

    # === 股东出资表 (Table 1): 动态循环 ===
    # Row0: header, Row1: pure {%tr for %}, Row2: data cells, Row3: pure {%tr endfor %}
    table1 = doc.tables[1]
    header_fields = [
        "shareholder_name",
        "subscribed_amount",
        "subscribed_ratio",
        "paid_amount",
        "paid_ratio",
    ]
    # Row 1: clear all cells, put only {%tr for %} in cell 0
    for ci in range(len(table1.rows[1].cells)):
        table1.rows[1].cells[ci].paragraphs[0].clear()
    table1.rows[1].cells[0].paragraphs[0].add_run("{%tr for sh in data.shareholders %}")
    # Row 2: data row with {{ sh.xxx }}
    data_row = table1.add_row()
    for ci, field in enumerate(header_fields):
        data_row.cells[ci].paragraphs[0].clear()
        data_row.cells[ci].paragraphs[0].add_run("{{ sh." + field + " }}")
    # Row 3: endfor
    endfor_row = table1.add_row()
    endfor_row.cells[0].paragraphs[0].clear()
    endfor_row.cells[0].paragraphs[0].add_run("{%tr endfor %}")

    # === 股东信息表 (Table 2): 动态循环 ===
    # Row0: header, then for row, data row, endfor row
    table2 = doc.tables[2]
    # Set header
    table2.rows[0].cells[0].paragraphs[0].clear()
    table2.rows[0].cells[0].paragraphs[0].add_run("\u5e8f\u53f7")
    table2.rows[0].cells[1].paragraphs[0].clear()
    table2.rows[0].cells[1].paragraphs[0].add_run("\u80a1\u4e1c\u540d\u79f0")
    # Delete all rows except header
    while len(table2.rows) > 1:
        tbl_elem = table2.rows[1]._tr.getparent()
        tbl_elem.remove(table2.rows[1]._tr)
    # Row 1: pure {%tr for %}
    for_row = table2.add_row()
    for_row.cells[0].paragraphs[0].add_run("{%tr for si in data.shareholder_info %}")
    # Row 2: data row
    si_data_row = table2.add_row()
    si_data_row.cells[0].paragraphs[0].add_run("{{ si.no }}")
    si_data_row.cells[1].paragraphs[0].add_run("{{ si.detail }}")
    # Row 3: endfor
    si_endfor = table2.add_row()
    si_endfor.cells[0].paragraphs[0].add_run("{%tr endfor %}")

    doc.save(str(TEMPLATE_OUT))
    return True


def setup_data():
    """创建测试数据Excel"""
    import openpyxl
    from openpyxl.styles import Font

    wb = openpyxl.Workbook()

    # Sheet 1: 全局变量
    ws_global = wb.active
    ws_global.title = "00_全局信息"
    ws_global.append(["字段编码", "值"])
    global_data = [
        ["g.company_name", "深圳市创新科技有限公司"],
        ["g.credit_code", "91440300MA5XXXXX00"],
        ["g.registered_address", "深圳市南山区粤海街道科技园南路88号"],
        ["g.legal_representative", "张伟"],
        ["g.company_type", "有限责任公司"],
        ["g.registered_capital", 1000.00],
        ["g.business_scope", "软件开发、技术咨询、技术服务"],
        ["g.established_date", "2018-06-15"],
        ["g.operating_period", "2018-06-15 至 长期"],
    ]
    for row_data in global_data:
        ws_global.append(row_data)
    ws_global.column_dimensions["A"].width = 25
    ws_global.column_dimensions["B"].width = 50

    # Sheet 2: 股东出资 (data.shareholders)
    ws_sh = wb.create_sheet("data.shareholders")
    ws_sh.append([
        "shareholder_name",
        "subscribed_amount",
        "subscribed_ratio",
        "paid_amount",
        "paid_ratio",
    ])
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

    # Sheet 3: 股东信息 (data.shareholder_info)
    ws_info = wb.create_sheet("data.shareholder_info")
    ws_info.append(["no", "detail"])
    info_data = [
        [1, "张伟-执行董事"],
        [2, "李芳-监事"],
        [3, "王强-股东"],
    ]
    for row_data in info_data:
        ws_info.append(row_data)
    ws_info.column_dimensions["A"].width = 10
    ws_info.column_dimensions["B"].width = 25

    wb.save(str(DATA_FILE))
    return True


def run_render():
    """调用 template_renderer 执行渲染"""
    from src.renderer import render_single

    success = render_single(
        str(DATA_FILE),
        str(TEMPLATE_OUT),
        str(OUTPUT_FILE),
        strict=False,
        check_vars=True,
    )
    return success


def verify_output():
    """验证渲染输出文档内容正确"""
    doc = Document(str(OUTPUT_FILE))
    all_text = " ".join(p.text for p in doc.paragraphs)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                all_text += " " + c.text

    checks = {
        "公司名称已替换": "深圳市创新科技有限公司" in all_text,
        "信用代码已替换": "91440300MA5XXXXX00" in all_text,
        "法定代表人已替换": "张伟" in all_text,
        "注册资本已替换": "1000" in all_text,
        "成立日期已替换": "2018-06-15" in all_text,
        "股东李芳在表格中": "李芳" in all_text,
        "认缴比例已替换": "250" in all_text,
        "无残留Jinja2标签": "{{" not in all_text and "{%" not in all_text,
    }

    all_pass = True
    print("\n" + "=" * 50)
    print("输出验证:")
    for name, result in checks.items():
        status = "PASS" if result else "FAIL"
        print("  [{}] {}".format(status, name))
        if not result:
            all_pass = False
    print("=" * 50)

    return all_pass


def main():
    print("=" * 50)
    print("TDD: 报告模板渲染测试")
    print("=" * 50)

    # Step 1: 创建模板 (从原始报告测试.docx转换)
    print("\n[SETUP] 创建模板...")
    if not setup_template():
        print("ERROR: 模板创建失败")
        return EXIT_FAILURE
    print("  模板已生成: {}".format(TEMPLATE_OUT))

    # Step 2: 创建测试数据
    print("\n[SETUP] 创建测试数据...")
    if not setup_data():
        print("ERROR: 数据创建失败")
        return EXIT_FAILURE
    print("  数据已生成: {}".format(DATA_FILE))

    # Step 3: 执行渲染
    print("\n[RENDER] 执行渲染...")
    os.makedirs(OUTPUT_FILE.parent, exist_ok=True)
    success = run_render()
    if not success:
        print("ERROR: 渲染失败")
        return EXIT_FAILURE
    print("  输出已生成: {}".format(OUTPUT_FILE))

    # Step 4: 验证输出
    all_pass = verify_output()

    if all_pass:
        print("\n✓ 所有检查通过 - 测试通过!")
        return EXIT_SUCCESS
    else:
        print("\n✗ 存在失败检查 - 测试未通过")
        return EXIT_FAILURE


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)
