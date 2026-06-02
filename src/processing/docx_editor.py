#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""DocxEditor — docx 文件编辑上下文管理器

将 zip 解包、XML 解析、写回、重新打包的生命周期封装在 with 块中。
预处理器只需关心 XML 树变换，不再管理临时文件和 zip 操作。
"""

import os
import zipfile
import tempfile
import shutil
import logging
from pathlib import Path
from typing import Optional

from lxml import etree

logger = logging.getLogger(__name__)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


class DocxEditor:
    """docx 文件编辑器 — 上下文管理器

    Usage:
        with DocxEditor("template.docx") as editor:
            root = editor.root
            # 修改 XML 树...
            modified = do_something(root)
        editor.save_to("output.docx", modified=modified)
    """

    def __init__(self, input_path: str) -> None:
        self._input_path = input_path
        self._tmp_dir: Optional[str] = None
        self.root: Optional[etree._Element] = None
        self._tree: Optional[etree._ElementTree] = None

    def __enter__(self) -> "DocxEditor":
        self._tmp_dir = tempfile.mkdtemp()
        with zipfile.ZipFile(self._input_path, "r") as zin:
            zin.extractall(self._tmp_dir)

        doc_xml = Path(self._tmp_dir) / "word" / "document.xml"
        self._tree = etree.parse(str(doc_xml))
        self.root = self._tree.getroot()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        if self._tmp_dir:
            shutil.rmtree(self._tmp_dir, ignore_errors=True)
            self._tmp_dir = None
        return None

    def save_to(self, output_path: str, modified: bool = False) -> None:
        """将 XML 写回（如果有改动）并重新打包为 output_path。"""
        assert self._tmp_dir is not None, "save_to 只能在 with 块内调用"

        if modified and self._tree is not None:
            doc_xml = Path(self._tmp_dir) / "word" / "document.xml"
            self._tree.write(
                str(doc_xml), xml_declaration=True, encoding="UTF-8", standalone=True
            )

        if Path(output_path).exists():
            Path(output_path).unlink()
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for dirpath, _, filenames in os.walk(self._tmp_dir):
                for fname in filenames:
                    fpath = Path(dirpath) / fname
                    arcname = str(fpath.relative_to(self._tmp_dir)).replace("\\", "/")
                    zout.write(str(fpath), arcname)
