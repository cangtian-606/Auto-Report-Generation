#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""集中日志配置 — stdout/stderr 分离、文件持久化、级别独立控制、运行ID追踪"""

import sys
import uuid
import time
import logging
from contextlib import contextmanager
from pathlib import Path
from typing import Optional, Dict

_run_id: str = ""
_stage_times: Dict[str, float] = {}


def init_run_id() -> str:
    global _run_id
    _run_id = uuid.uuid4().hex[:8]
    return _run_id


def get_run_id() -> str:
    return _run_id


def get_stage_times() -> Dict[str, float]:
    return dict(_stage_times)


def reset_stage_times() -> None:
    _stage_times.clear()


@contextmanager
def timed_stage(name: str):
    start = time.perf_counter()
    try:
        yield
    finally:
        _stage_times[name] = time.perf_counter() - start


class _RunIdFilter(logging.Filter):
    def filter(self, record: logging.LogRecord) -> bool:
        record.run_id = _run_id or "-" * 8
        return True


class _LevelRangeFilter(logging.Filter):
    def __init__(self, min_level: int, max_level: int):
        super().__init__()
        self.min_level = min_level
        self.max_level = max_level

    def filter(self, record: logging.LogRecord) -> bool:
        return self.min_level <= record.levelno < self.max_level


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
        "%(asctime)s [%(run_id)s] [%(levelname)7s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    run_id_filter = _RunIdFilter()

    stdout_handler = logging.StreamHandler(sys.stdout)
    stdout_handler.setLevel(terminal_level)
    stdout_handler.addFilter(_LevelRangeFilter(min_level=terminal_level, max_level=logging.WARNING - 1))
    stdout_handler.addFilter(run_id_filter)
    stdout_handler.setFormatter(fmt)
    root.addHandler(stdout_handler)

    stderr_handler = logging.StreamHandler(sys.stderr)
    stderr_handler.setLevel(logging.WARNING)
    stderr_handler.addFilter(run_id_filter)
    stderr_handler.setFormatter(fmt)
    root.addHandler(stderr_handler)

    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(str(log_path), encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        file_handler.addFilter(run_id_filter)
        file_handler.setFormatter(fmt)
        root.addHandler(file_handler)
