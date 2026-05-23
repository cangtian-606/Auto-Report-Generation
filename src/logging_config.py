#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""集中日志配置 — stdout/stderr 分离、文件持久化、级别独立控制"""

import sys
import logging
from pathlib import Path
from typing import Optional


def configure_logging(
    terminal_level: int = logging.INFO,
    log_file: Optional[str] = None,
) -> None:
    """
    配置全局日志系统

    - stdout: DEBUG ~ INFO（进度/状态/成功），terminal_level 控制
    - stderr: WARNING 及以上（问题/错误），始终启用
    - 文件: DEBUG 及以上，始终完整记录

    Args:
        terminal_level: 终端 stdout 最低级别，默认 INFO
        log_file: 文件日志路径，None 则仅终端输出
    """
    root = logging.getLogger()
    root.setLevel(logging.DEBUG)

    for h in list(root.handlers):
        root.removeHandler(h)

    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)7s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # stdout: only DEBUG ~ WARNING-1 (i.e. DEBUG, INFO)
    stdout_handler = logging.StreamHandler(sys.stdout)
    stdout_handler.setLevel(terminal_level)
    stdout_handler.addFilter(_LevelRangeFilter(min_level=terminal_level, max_level=logging.WARNING - 1))
    stdout_handler.setFormatter(fmt)
    root.addHandler(stdout_handler)

    # stderr: WARNING and above
    stderr_handler = logging.StreamHandler(sys.stderr)
    stderr_handler.setLevel(logging.WARNING)
    stderr_handler.setFormatter(fmt)
    root.addHandler(stderr_handler)

    # file: DEBUG and above, always full
    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(str(log_path), encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        file_handler.setFormatter(fmt)
        root.addHandler(file_handler)


class _LevelRangeFilter(logging.Filter):
    def __init__(self, min_level: int, max_level: int):
        super().__init__()
        self.min_level = min_level
        self.max_level = max_level

    def filter(self, record: logging.LogRecord) -> bool:
        return self.min_level <= record.levelno < self.max_level
