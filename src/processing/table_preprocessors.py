#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
表格预处理器 — {%tc for %} 列循环 + {%tr for %} 行循环继承

核心思想：首行声明，整表继承。
  TcInheritancePreprocessor: 首行声明 {%tc for %} → 后续行自动注入列循环标签
  TrInheritancePreprocessor: 首行声明 {%tr for %} → 剩余数据行自动继承行循环上下文
"""

import re
import zipfile
import tempfile
import shutil
import copy
import logging
from pathlib import Path
from typing import List, Optional, Tuple

from lxml import etree

logger = logging.getLogger(__name__)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# ---------------------------------------------------------------------------
# 共享正则
# ---------------------------------------------------------------------------
TC_FOR_RE = re.compile(r"\{%tc\s+for\s+(\w+)\s+in\s+(.+?)%\}")
TC_ENDFOR_RE = re.compile(r"\{%tc\s+endfor\s*%\}")
TR_FOR_RE = re.compile(r"\{%tr\s+for\s+(\w+)\s+in\s+(.+?)%\}")
TR_ENDFOR_RE = re.compile(r"\{%tr\s+endfor\s*%\}")
TR_TAG_RE = re.compile(r"\{%tr\s+(for|endfor)")


# ---------------------------------------------------------------------------
# 共享的 XML 操作基类
# ---------------------------------------------------------------------------
class _XmlHelper:
    """XML 文档操作的公共方法。"""

    @staticmethod
    def _cell_text(tc: etree._Element) -> str:
        texts = []
        for t_elem in tc.iter(f"{{{W}}}t"):
            if t_elem.text:
                texts.append(t_elem.text)
        return "".join(texts)

    @staticmethod
    def _clear_cell(tc: etree._Element) -> str:
        saved = []
        for t_elem in tc.iter(f"{{{W}}}t"):
            if t_elem.text:
                saved.append(t_elem.text)
                t_elem.text = ""
        return "".join(saved)

    @staticmethod
    def _prepend_text(tc: etree._Element, text: str) -> None:
        for t_elem in tc.iter(f"{{{W}}}t"):
            if t_elem.text is not None:
                t_elem.text = text + t_elem.text
                return
        _XmlHelper._create_text_element(tc, text)

    @staticmethod
    def _append_text(tc: etree._Element, text: str) -> None:
        last_t = None
        for t_elem in tc.iter(f"{{{W}}}t"):
            last_t = t_elem
        if last_t is not None and last_t.text is not None:
            last_t.text = last_t.text + text
            return
        _XmlHelper._create_text_element(tc, text)

    @staticmethod
    def _set_text(tc: etree._Element, text: str) -> None:
        found = False
        first = True
        for t_elem in tc.iter(f"{{{W}}}t"):
            found = True
            if first:
                t_elem.text = text
                first = False
            else:
                t_elem.text = ""
        if not found:
            _XmlHelper._create_text_element(tc, text)

    @staticmethod
    def _create_text_element(tc: etree._Element, text: str) -> None:
        p = tc.find(f"{{{W}}}p")
        if p is None:
            p = etree.SubElement(tc, f"{{{W}}}p")
        r = etree.SubElement(p, f"{{{W}}}r")
        t_elem = etree.SubElement(r, f"{{{W}}}t")
        t_elem.text = text

    @staticmethod
    def _has_flow_tags(row: etree._Element) -> bool:
        for tc in row.iter(f"{{{W}}}tc"):
            text = _XmlHelper._cell_text(tc)
            if TC_FOR_RE.search(text) or TC_ENDFOR_RE.search(text):
                return True
            if TR_TAG_RE.search(text):
                return True
        return False


# ---------------------------------------------------------------------------
# TcInheritancePreprocessor — 列循环继承
# ---------------------------------------------------------------------------
class TcInheritancePreprocessor:
    """检测表格首行的 {%tc for %} 声明，自动施加到后续行。"""

    def process(self, input_path: str, output_path: str) -> int:
        tmp_dir = tempfile.mkdtemp()
        try:
            with zipfile.ZipFile(input_path, "r") as zin:
                zin.extractall(tmp_dir)
            doc_xml = Path(tmp_dir) / "word" / "document.xml"
            tree = etree.parse(str(doc_xml))
            root = tree.getroot()

            modified_tables = 0
            for tbl in root.iter(f"{{{W}}}tbl"):
                if self._apply_tc_inheritance(tbl):
                    modified_tables += 1

            if modified_tables > 0:
                tree.write(str(doc_xml), xml_declaration=True, encoding="UTF-8", standalone=True)

            self._repack(tmp_dir, output_path)
            return modified_tables
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    # ---- tc 核心逻辑 ----
    def _apply_tc_inheritance(self, tbl: etree._Element) -> bool:
        row_elems = list(tbl.iter(f"{{{W}}}tr"))
        if len(row_elems) < 2:
            return False

        tc_info = self._detect_tc_for(row_elems[0])
        if tc_info is None:
            return False

        var_name, expr, start_idx, end_idx = tc_info
        modified = False

        for row in row_elems[1:]:
            cells = list(row.iter(f"{{{W}}}tc"))
            if start_idx >= len(cells) or end_idx >= len(cells):
                continue
            if _XmlHelper._has_flow_tags(row):
                continue

            has_template_cells = end_idx - start_idx > 1
            if has_template_cells:
                saved_text = _XmlHelper._clear_cell(cells[start_idx])
                _XmlHelper._clear_cell(cells[start_idx + 1])
                _XmlHelper._set_text(cells[start_idx + 1], saved_text)

            tag = f"{{%tc for {var_name} in {expr} %}}"
            _XmlHelper._prepend_text(cells[start_idx], tag)
            _XmlHelper._append_text(cells[end_idx], "{%tc endfor %}")
            modified = True

        return modified

    def _detect_tc_for(self, row: etree._Element) -> Optional[Tuple[str, str, int, int]]:
        cells = list(row.iter(f"{{{W}}}tc"))
        start_idx = None
        var_name = None
        expr = None

        for ci, tc in enumerate(cells):
            m = TC_FOR_RE.search(_XmlHelper._cell_text(tc))
            if m:
                start_idx, var_name, expr = ci, m.group(1), m.group(2)
                break
        if start_idx is None:
            return None

        for ci in range(start_idx + 1, len(cells)):
            if TC_ENDFOR_RE.search(_XmlHelper._cell_text(cells[ci])):
                return (var_name, expr, start_idx, ci)
        return None

    @staticmethod
    def _repack(tmp_dir: str, output_path: str) -> None:
        if Path(output_path).exists():
            Path(output_path).unlink()
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for dirpath, _, filenames in __import__("os").walk(tmp_dir):
                for fname in filenames:
                    fpath = Path(dirpath) / fname
                    arcname = str(fpath.relative_to(tmp_dir)).replace("\\", "/")
                    zout.write(str(fpath), arcname)


# ---------------------------------------------------------------------------
# TrInheritancePreprocessor — 行循环继承
# ---------------------------------------------------------------------------
class TrInheritancePreprocessor:
    """检测表格中的 {%tr for %} 声明，自动将未被包裹的数据行纳入循环。"""

    def process(self, input_path: str, output_path: str) -> int:
        tmp_dir = tempfile.mkdtemp()
        try:
            with zipfile.ZipFile(input_path, "r") as zin:
                zin.extractall(tmp_dir)
            doc_xml = Path(tmp_dir) / "word" / "document.xml"
            tree = etree.parse(str(doc_xml))
            root = tree.getroot()

            modified_tables = 0
            for tbl in root.iter(f"{{{W}}}tbl"):
                if self._apply_tr_inheritance(tbl):
                    modified_tables += 1

            if modified_tables > 0:
                tree.write(str(doc_xml), xml_declaration=True, encoding="UTF-8", standalone=True)

            TcInheritancePreprocessor._repack(tmp_dir, output_path)
            return modified_tables
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    # ---- tr 核心逻辑 ----
    def _apply_tr_inheritance(self, tbl: etree._Element) -> bool:
        """
        与 tc 对称的"自动补全"模式：
        - 有 {%tr for %} 但无 {%tr endfor %} → 在表末行追加 {%tr endfor %}
        - 同时将 for 和 endfor 之间的数据行纳入循环（不重复注入已有标签的行）
        """
        row_elems = list(tbl.iter(f"{{{W}}}tr"))
        if len(row_elems) < 2:
            return False

        for_idx, endfor_idx = self._find_tr_range(row_elems)
        if for_idx is None:
            return False

        modified = False
        if endfor_idx is None:
            self._append_tr_endfor(tbl, row_elems[-1])
            endfor_idx = len(row_elems) - 1
            modified = True

        return modified

    def _find_tr_range(self, row_elems: List[etree._Element]) -> Tuple[Optional[int], Optional[int]]:
        for_idx = None
        endfor_idx = None
        for ri, row in enumerate(row_elems):
            for tc in row.iter(f"{{{W}}}tc"):
                text = _XmlHelper._cell_text(tc)
                if for_idx is None and TR_FOR_RE.search(text):
                    for_idx = ri
                elif for_idx is not None and TR_ENDFOR_RE.search(text):
                    endfor_idx = ri
                    break
            if endfor_idx is not None:
                break
        return for_idx, endfor_idx

    def _append_tr_endfor(self, tbl: etree._Element, last_row: etree._Element) -> None:
        """在表格末尾创建新行, 写入 {%tr endfor %}。"""
        new_row = etree.SubElement(tbl, f"{{{W}}}tr")
        for tc in last_row.iter(f"{{{W}}}tc"):
            new_tc = etree.SubElement(new_row, f"{{{W}}}tc")
            prs = list(tc.iter(f"{{{W}}}tcPr"))
            if prs:
                new_tc.append(copy.deepcopy(prs[0]))
            new_p = etree.SubElement(new_tc, f"{{{W}}}p")
            new_r = etree.SubElement(new_p, f"{{{W}}}r")
            new_t = etree.SubElement(new_r, f"{{{W}}}t")
            new_t.text = ""
        cells = list(new_row.iter(f"{{{W}}}tc"))
        if cells:
            _XmlHelper._set_text(cells[0], "{%tr endfor %}")
