# -*- coding: utf-8 -*-
"""测试 Config / 数据结构 / 日志"""

import pytest
import logging
from unittest.mock import patch
from link_checker_core import (
    Config, cfg,
    UrlEntry, CheckResult, ProcessResult,
    get_logger, log, _logger,
)


class TestConfig:
    """Config 数据类测试"""

    def test_defaults(self):
        c = Config()
        assert c.TIMEOUT == 10
        assert c.MAX_WORKERS == 10
        assert c.MAX_RETRIES == 2
        assert c.RETRY_BACKOFF == 1.5
        assert c.MAX_REDIRECTS == 5
        assert c.BATCH_SIZE == 200
        assert c.LARGE_BATCH_THRESHOLD == 500

    def test_frozen_immutable(self):
        c = Config()
        with pytest.raises(Exception):
            c.TIMEOUT = 5

    def test_custom_config(self):
        c = Config(TIMEOUT=5, MAX_WORKERS=20, MAX_RETRIES=1)
        assert c.TIMEOUT == 5
        assert c.MAX_WORKERS == 20
        assert c.MAX_RETRIES == 1
        assert c.BATCH_SIZE == 200

    def test_user_agent_is_string(self):
        assert isinstance(cfg.USER_AGENT, str)
        assert "Mozilla" in cfg.USER_AGENT

    def test_local_networks_cover_ipv4(self):
        import ipaddress
        local_ips = [
            "127.0.0.1", "10.0.0.1", "172.16.0.1",
            "192.168.1.1", "169.254.1.1",
        ]
        for ip_str in local_ips:
            ip = ipaddress.ip_address(ip_str)
            assert any(ip in net for net in cfg.LOCAL_NETWORKS), f"{ip_str} 应在内网范围"

    def test_local_networks_cover_ipv6(self):
        import ipaddress
        local_ips_v6 = ["::1", "fc00::1", "fe80::1"]
        for ip_str in local_ips_v6:
            ip = ipaddress.ip_address(ip_str)
            assert any(ip in net for net in cfg.LOCAL_NETWORKS), f"{ip_str} 应在内网范围"

    def test_public_ip_not_in_local(self):
        import ipaddress
        public_ips = ["8.8.8.8", "1.1.1.1", "93.184.216.34"]
        for ip_str in public_ips:
            ip = ipaddress.ip_address(ip_str)
            assert not any(ip in net for net in cfg.LOCAL_NETWORKS), f"{ip_str} 不应在内网范围"


class TestDataStructures:
    """数据结构测试"""

    def test_url_entry_creation(self):
        entry = UrlEntry(sheet="Sheet1", row=2, col_idx=0, col_name="链接",
                         url="https://example.com")
        assert entry.sheet == "Sheet1"
        assert entry.row == 2
        assert entry.col_idx == 0
        assert entry.col_name == "链接"
        assert entry.url == "https://example.com"

    def test_check_result_valid(self):
        cr = CheckResult(url="https://ok.com", valid=True, detail="HTTP 200", retries=0)
        assert cr.valid is True
        assert cr.detail == "HTTP 200"

    def test_check_result_invalid(self):
        cr = CheckResult(url="https://bad.com", valid=False, detail="连接超时", retries=2)
        assert cr.valid is False
        assert cr.retries == 2

    def test_process_result_defaults(self):
        pr = ProcessResult(
            input_path="a.xlsx", output_path="b.xlsx",
            total_urls=10, valid_count=7, invalid_count=3, skipped_count=0,
        )
        assert pr.results == []
        assert pr.errors == []

    def test_process_result_with_results(self):
        pr = ProcessResult(
            input_path="a.xlsx", output_path="b.xlsx",
            total_urls=3, valid_count=2, invalid_count=1, skipped_count=0,
            results=[{"url": "x"}, {"url": "y"}],
            errors=["err1"],
        )
        assert len(pr.results) == 2
        assert pr.errors == ["err1"]


class TestLogging:
    """日志系统测试"""

    def test_get_logger_returns_logger(self):
        logger = get_logger()
        assert isinstance(logger, logging.Logger)
        assert logger.name == "LinkChecker"

    def test_get_logger_singleton(self):
        a = get_logger()
        b = get_logger()
        assert a is b

    def test_logger_has_handlers(self):
        logger = get_logger()
        assert len(logger.handlers) >= 1

    def test_global_log_alias(self):
        assert log is get_logger()

    def test_permission_error_on_file_handler(self):
        """RotatingFileHandler 创建失败（PermissionError）时静默跳过"""
        import link_checker_core
        # 重置 _logger 强制重新初始化
        link_checker_core._logger = None
        with patch("link_checker_core.logging.handlers.RotatingFileHandler",
                   side_effect=PermissionError("denied")):
            logger = link_checker_core.get_logger()
            # 不应崩溃，至少有 StreamHandler
            assert logger is not None
            assert any(
                isinstance(h, logging.StreamHandler)
                for h in logger.handlers
            )
        # 恢复
        link_checker_core._logger = None
        link_checker_core.get_logger()
