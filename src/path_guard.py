#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""路径安全校验 — 防止路径遍历攻击"""

from pathlib import Path
from typing import List, Optional


def validate_path(path: str, allowed_dirs: Optional[List[str]] = None,
                  must_exist: bool = False) -> str:
    """
    校验和规范化路径，防止路径遍历攻击

    Args:
        path: 待校验的路径
        allowed_dirs: 允许的目录列表，若为 None 则允许所有路径
        must_exist: 是否要求路径必须已存在

    Returns:
        规范化后的绝对路径

    Raises:
        PermissionError: 路径超出允许范围
        FileNotFoundError: must_exist=True 且路径不存在
    """
    path_obj = Path(path)

    normalized = path_obj.resolve(strict=False)

    if allowed_dirs:
        allowed_abs = [Path(dir).resolve() for dir in allowed_dirs]
        in_allowed = False
        for dir_abs in allowed_abs:
            try:
                normalized.relative_to(dir_abs)
                in_allowed = True
                break
            except ValueError:
                continue
        if not in_allowed:
            raise PermissionError(f"路径 '{path}' 超出允许范围，仅允许访问: {allowed_dirs}")

    if must_exist and not normalized.exists():
        raise FileNotFoundError(f"文件不存在: {normalized}")
    return str(normalized)
