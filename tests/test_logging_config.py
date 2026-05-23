#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试日志配置: stdout/stderr 分离、级别控制、文件输出"""

import io
import sys
import logging
import tempfile
from pathlib import Path

import pytest

from src.logging_config import configure_logging, _LevelRangeFilter


class TestLevelRangeFilter:
    """测试 _LevelRangeFilter 区间过滤"""

    def test_inside_range_passes(self):
        f = _LevelRangeFilter(min_level=logging.DEBUG, max_level=logging.WARNING)
        record = logging.LogRecord("test", logging.INFO, "", 0, "msg", (), None)
        assert f.filter(record) is True

    def test_below_range_blocked(self):
        f = _LevelRangeFilter(min_level=logging.INFO, max_level=logging.WARNING)
        record = logging.LogRecord("test", logging.DEBUG, "", 0, "msg", (), None)
        assert f.filter(record) is False

    def test_at_or_above_max_blocked(self):
        f = _LevelRangeFilter(min_level=logging.DEBUG, max_level=logging.WARNING)
        record = logging.LogRecord("test", logging.WARNING, "", 0, "msg", (), None)
        assert f.filter(record) is False

    def test_well_above_max_blocked(self):
        f = _LevelRangeFilter(min_level=logging.DEBUG, max_level=logging.WARNING)
        record = logging.LogRecord("test", logging.ERROR, "", 0, "msg", (), None)
        assert f.filter(record) is False

    def test_at_minimum_passes(self):
        f = _LevelRangeFilter(min_level=logging.INFO, max_level=logging.ERROR)
        record = logging.LogRecord("test", logging.INFO, "", 0, "msg", (), None)
        assert f.filter(record) is True


class TestConfigureLogging:
    """测试 configure_logging 的 stdout/stderr 分离和级别控制"""

    def setup_method(self):
        root = logging.getLogger()
        for h in list(root.handlers):
            root.removeHandler(h)

    def test_stdout_gets_info_warning_goes_to_stderr(self):
        """info 去 stdout, warning 去 stderr"""
        stdout_capture = io.StringIO()
        stderr_capture = io.StringIO()
        old_stdout, old_stderr = sys.stdout, sys.stderr
        sys.stdout = stdout_capture
        sys.stderr = stderr_capture

        try:
            configure_logging(terminal_level=logging.DEBUG)
            logger = logging.getLogger("test_stdout_stderr")
            logger.info("progress message")
            logger.warning("warning message")

            stdout_output = stdout_capture.getvalue()
            stderr_output = stderr_capture.getvalue()

            assert "progress message" in stdout_output
            assert "progress message" not in stderr_output
            assert "warning message" in stderr_output
            assert "warning message" not in stdout_output
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr

    def test_quiet_mode_hides_info_from_terminal(self):
        """--quiet 模式: DEBUG/INFO 不出现在 stdout"""
        stdout_capture = io.StringIO()
        stderr_capture = io.StringIO()
        old_stdout, old_stderr = sys.stdout, sys.stderr
        sys.stdout = stdout_capture
        sys.stderr = stderr_capture

        try:
            configure_logging(terminal_level=logging.WARNING)
            logger = logging.getLogger("test_quiet")
            logger.info("should not appear")
            logger.warning("should appear")

            stdout_output = stdout_capture.getvalue()
            stderr_output = stderr_capture.getvalue()

            assert "should not appear" not in stdout_output
            assert "should appear" in stderr_output
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr

    def test_verbose_mode_shows_debug(self):
        """--verbose: DEBUG 出现在 stdout"""
        stdout_capture = io.StringIO()
        old_stdout = sys.stdout
        sys.stdout = stdout_capture

        try:
            configure_logging(terminal_level=logging.DEBUG)
            logger = logging.getLogger("test_verbose")
            logger.debug("debug detail")

            stdout_output = stdout_capture.getvalue()
            assert "debug detail" in stdout_output
        finally:
            sys.stdout = old_stdout

    def test_error_goes_to_stderr_only(self):
        """error 仅出现在 stderr"""
        stdout_capture = io.StringIO()
        stderr_capture = io.StringIO()
        old_stdout, old_stderr = sys.stdout, sys.stderr
        sys.stdout = stdout_capture
        sys.stderr = stderr_capture

        try:
            configure_logging(terminal_level=logging.INFO)
            logger = logging.getLogger("test_error")
            logger.error("critical failure")

            stdout_output = stdout_capture.getvalue()
            stderr_output = stderr_capture.getvalue()

            assert "critical failure" not in stdout_output
            assert "critical failure" in stderr_output
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr

    def test_file_handler_writes_debug(self):
        """文件日志始终记录 DEBUG"""
        with tempfile.TemporaryDirectory() as tmp:
            log_file = str(Path(tmp) / "app.log")
            configure_logging(terminal_level=logging.INFO, log_file=log_file)
            logger = logging.getLogger("test_file")
            logger.debug("file only debug")
            logger.info("file info")

            root = logging.getLogger()
            for h in root.handlers:
                h.flush()
                h.close()
            root.handlers.clear()

            content = Path(log_file).read_text(encoding='utf-8')
            assert "file only debug" in content
            assert "file info" in content

    def test_exception_includes_traceback_in_stderr(self):
        """logger.exception 附加堆栈到 stderr"""
        stderr_capture = io.StringIO()
        old_stderr = sys.stderr
        sys.stderr = stderr_capture

        try:
            configure_logging(terminal_level=logging.INFO)
            logger = logging.getLogger("test_exc")
            try:
                raise ValueError("simulated")
            except ValueError:
                logger.exception("caught error")

            stderr_output = stderr_capture.getvalue()
            assert "caught error" in stderr_output
            assert "ValueError" in stderr_output
            assert "Traceback" in stderr_output
        finally:
            sys.stderr = old_stderr
