#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""RenderPipeline — 渲染流水线

将 generate() 中的阶段编排、临时文件生命周期、汇总摘要
提炼为独立模块，提升局部性和可测试性。
"""

import os
import logging
from contextlib import contextmanager
from typing import Any, Callable, Dict, List, Optional, Tuple

from ..logging_config import timed_stage, get_run_id, init_run_id, reset_stage_times

logger = logging.getLogger(__name__)

SEP = "─" * 54

State = Dict[str, Any]
StageFn = Callable[[State, 'RenderPipeline'], Optional[State]]


class RenderPipeline:
    """渲染流水线 — 顺序执行命名阶段，管理临时文件，输出汇总摘要。

    Usage:
        pipeline = RenderPipeline()
        pipeline.add_stage("读数据", read_stage)
        pipeline.add_stage("构建", build_stage)
        pipeline.run(data_path="...", template_path="...")
    """

    def __init__(self) -> None:
        self._stages: List[Tuple[str, StageFn]] = []
        self._temp_files: List[str] = []

    def add_stage(self, name: str, fn: StageFn) -> "RenderPipeline":
        self._stages.append((name, fn))
        return self

    def run(self, **initial_state: Any) -> State:
        init_run_id()
        reset_stage_times()

        state: State = dict(initial_state)
        for name, fn in self._stages:
            with timed_stage(name):
                result = fn(state, self)
            if isinstance(result, dict):
                state.update(result)
        return state

    def track_temp(self, path: str) -> None:
        self._temp_files.append(path)

    def cleanup_temps(self, keep: str = "") -> None:
        for p in self._temp_files:
            if p != keep:
                try:
                    os.unlink(p)
                except OSError:
                    pass
        self._temp_files.clear()
