# -*- coding: utf-8 -*-
"""测试 URL 安全检查 — is_url_safe()"""

import pytest
from link_checker_core import is_url_safe


class TestUrlSafePass:
    """正常 URL 应通过安全检查"""

    @pytest.mark.parametrize("url", [
        "https://www.google.com",
        "https://github.com/openai",
        "http://example.com/path?q=1",
        "https://sub.domain.co.uk/page",
        "https://192.0.2.1/page",
        "https://[2001:4860:4860::8888]",
        "https://example.com:8080/path",
        "http://localhost:3000",
    ])
    def test_safe_url_passes(self, url):
        ok, reason = is_url_safe(url)
        assert ok is True, f"{url} 应该安全，但被拦截: {reason}"


class TestUrlSafeBlock:
    """危险 URL 应被拦截"""

    @pytest.mark.parametrize("url,expected_reason", [
        # ---- 回环地址 ----
        ("http://127.0.0.1/admin",      "回环地址"),
        ("http://[::1]:8080",           "回环地址"),
        # ---- 内网 IPv4 ----
        ("https://10.0.0.5/api",        "内网地址"),
        ("http://172.16.0.1/login",     "内网地址"),
        ("https://192.168.1.100",       "内网地址"),
        # ---- 链路本地 ----
        ("http://169.254.1.1/config",   "链路本地"),
        # ---- 内网 IPv6 ----
        ("https://[fc00::1]/app",       "内网地址"),
        ("http://[fe80::1]",            "链路本地"),
        # ---- 危险协议 ----
        ("file:///etc/passwd",          "不允许的协议"),
        ("ftp://files.example.com",     "不允许的协议"),
        ("data:text/html,<script>",     "不允许的协议"),
        ("javascript:alert(1)",         "不允许的协议"),
        ("vbscript:msgbox(1)",          "不允许的协议"),
        # ---- 无 scheme / 无效 URL ----
        ("not-a-url",                   "不允许的协议: 无"),
        ("",                            "不允许的协议: 无"),
        ("   ",                         "不允许的协议: 无"),
        ("ht!tp://bad.com",             "不允许的协议: 无"),
        # ---- 内部域名后缀 ----
        ("https://server.local",        "内部域名后缀"),
        ("http://admin.internal/api",   "内部域名后缀"),
        ("https://portal.corp/login",   "内部域名后缀"),
        ("http://nas.home",             "内部域名后缀"),
        ("https://device.lan",          "内部域名后缀"),
    ])
    def test_unsafe_url_blocked(self, url, expected_reason):
        ok, reason = is_url_safe(url)
        assert ok is False, f"{url} 应被拦截"
        assert expected_reason in reason, f"原因应包含 '{expected_reason}'，实际: '{reason}'"


class TestUrlSafeEdgeCases:
    """边界情况"""

    def test_no_scheme(self):
        ok, reason = is_url_safe("www.example.com")
        assert ok is False
        assert "不允许的协议" in reason

    def test_none_input(self):
        ok, reason = is_url_safe(None)
        assert ok is False
        assert "不允许的协议" in reason

    def test_very_long_url(self):
        url = "https://example.com/" + "a" * 2000
        ok, reason = is_url_safe(url)
        assert isinstance(ok, bool)

    def test_url_with_special_chars(self):
        ok, reason = is_url_safe("https://example.com/path%20with%20spaces")
        assert ok is True

    def test_upper_case_ip_local(self):
        ok, reason = is_url_safe("HTTP://10.0.0.1/")
        assert ok is False
        assert "内网地址" in reason

    def test_upper_case_file_scheme(self):
        ok, reason = is_url_safe("FILE://host/path")
        assert ok is False
        assert "不允许的协议" in reason

    def test_domain_case_local(self):
        ok, reason = is_url_safe("https://SERVER.LOCAL")
        assert ok is False
        assert "内部域名后缀" in reason

    def test_url_without_hostname(self):
        """有 scheme 但没有 hostname"""
        ok, reason = is_url_safe("http:///path")
        assert ok is False
        assert "主机名" in reason

    def test_malformed_ipv6_url(self):
        """不完整的 IPv6 地址导致 URL 解析失败"""
        ok, reason = is_url_safe("https://[::1")
        assert ok is False
        assert "URL解析失败" in reason

    def test_invalid_ipv6_bracket(self):
        """无效的方括号 IPv6"""
        ok, reason = is_url_safe("https://[invalid]")
        assert ok is False
        assert "URL解析失败" in reason
