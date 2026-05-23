#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""YAML 数据读取器"""

import logging
from pathlib import Path
from typing import Any, Dict

try:
    import yaml
except ImportError:
    yaml = None

logger = logging.getLogger(__name__)


class YamlDataReader:
    """YAML 数据读取器 - 直接返回 context 结构，无需 mapper"""

    def __init__(self, file_path: str) -> None:
        if yaml is None:
            raise ImportError("请先安装 pyyaml: pip install pyyaml")
        self.file_path = file_path
        self.data: Dict[str, Any] = {}

    def read_all(self) -> Dict[str, Any]:
        logger.debug("读取YAML: %s", self.file_path)
        with open(self.file_path, 'r', encoding='utf-8') as f:
            self.data = yaml.safe_load(f) or {}
        return self.data

    def read_context(self) -> Dict[str, Any]:
        return self.read_all()
