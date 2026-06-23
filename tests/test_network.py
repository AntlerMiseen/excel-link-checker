# -*- coding: utf-8 -*-
"""测试网络检测 — check_single_url / batch_check_urls（Mock 版）"""

import pytest
import requests as req
from unittest.mock import MagicMock, patch
from link_checker_core import (
    check_single_url, batch_check_urls,
    CheckResult, UrlEntry, cfg, _thread_local,
)


@pytest.mark.usefixtures("reset_thread_local")
class TestCheckSingleUrl:
    """单 URL 检测测试"""

    def test_head_200_valid(self, mock_requests_head_ok):
        result = check_single_url("https://example.com")
        assert result.valid is True
        assert "200" in result.detail
        assert result.retries == 0

    def test_head_404_get_404_invalid(self, mock_requests_head_404):
        result = check_single_url("https://example.com/404")
        assert result.valid is False
        assert "404" in result.detail

    def test_head_404_get_200_valid(self):
        """HEAD 返回 404，GET 回退返回 200 — 应判定有效"""
        if hasattr(_thread_local, "session"):
            del _thread_local.session
        with patch("link_checker_core.requests.Session.head") as mock_head, \
             patch("link_checker_core.requests.Session.get") as mock_get:
            mock_resp_404 = MagicMock()
            mock_resp_404.status_code = 404
            mock_head.return_value = mock_resp_404
            mock_resp_200 = MagicMock()
            mock_resp_200.status_code = 200
            mock_get.return_value = mock_resp_200
            result = check_single_url("https://example.com")
            assert result.valid is True
            assert "200" in result.detail

    def test_connection_refused_retries(self):
        """连接拒绝应重试，耗尽后返回无效"""
        if hasattr(_thread_local, "session"):
            del _thread_local.session
        with patch("link_checker_core.requests.Session.head") as mock_head:
            mock_head.side_effect = req.exceptions.ConnectionError("refused")
            result = check_single_url("https://example.com")
            assert result.valid is False
            assert result.retries == cfg.MAX_RETRIES
            assert "无法连接" in result.detail

    def test_timeout_retries_exhausted(self, mock_requests_timeout):
        result = check_single_url("https://example.com")
        assert result.valid is False
        assert result.retries == cfg.MAX_RETRIES
        assert "超时" in result.detail

    def test_ssl_error_no_retry(self):
        """SSL ?????????SSLError ? ConnectionError ?????"""
        if hasattr(_thread_local, "session"):
            del _thread_local.session
        with patch("link_checker_core.requests.Session.head") as mock_head:
            mock_head.side_effect = req.exceptions.SSLError()
            result = check_single_url("https://example.com")
            assert result.valid is False
            assert result.retries == 0
            assert "SSL" in result.detail

    def test_too_many_redirects_no_retry(self, mock_requests_too_many_redirects):
        result = check_single_url("https://example.com")
        assert result.valid is False
        assert result.retries == 0
        assert "重定向" in result.detail

    def test_dns_failure_no_retry(self):
        """DNS 解析失败不应重试"""
        if hasattr(_thread_local, "session"):
            del _thread_local.session
        with patch("link_checker_core.requests.Session.head") as mock_head:
            orig = req.exceptions.ConnectionError("getaddrinfo failed")
            mock_head.side_effect = orig
            result = check_single_url("https://nonexistent.invalid")
            assert result.valid is False
            assert result.retries == 0
            assert "DNS" in result.detail

    def test_generic_request_exception(self):
        """其他 RequestException 子类（非 Timeout/Connection/SSL/Redirect）应重试"""
        if hasattr(_thread_local, "session"):
            del _thread_local.session
        with patch("link_checker_core.requests.Session.head") as mock_head:
            mock_head.side_effect = req.exceptions.HTTPError("generic error")
            result = check_single_url("https://example.com")
            assert result.valid is False
            assert result.retries == cfg.MAX_RETRIES
            assert "请求异常" in result.detail


@pytest.mark.usefixtures("reset_thread_local")
class TestBatchCheckUrls:
    """批量检测测试"""

    def test_empty_entries(self):
        results = batch_check_urls([])
        assert results == {}

    def test_unique_dedup(self):
        entries = [
            UrlEntry("S1", 2, 0, "A", "https://dup.com"),
            UrlEntry("S1", 3, 0, "A", "https://dup.com"),
            UrlEntry("S2", 2, 1, "B", "https://dup.com"),
        ]
        with patch("link_checker_core.check_single_url") as mock_check:
            mock_check.return_value = CheckResult(
                url="https://dup.com", valid=True, detail="OK", retries=0
            )
            results = batch_check_urls(entries)
            assert mock_check.call_count == 1
            assert "https://dup.com" in results

    def test_progress_callback_called(self, mock_requests_head_ok):
        entries = [
            UrlEntry("S1", 2, 0, "A", "https://a.com"),
            UrlEntry("S1", 3, 0, "A", "https://b.com"),
        ]
        calls = []

        def cb(completed, total, result):
            calls.append((completed, total))

        batch_check_urls(entries, progress_callback=cb)
        assert len(calls) >= 1
        assert calls[-1][0] == calls[-1][1]

    def test_unsafe_urls_marked_invalid(self):
        entries = [
            UrlEntry("S1", 2, 0, "A", "http://10.0.0.1/admin"),
            UrlEntry("S1", 3, 0, "A", "ftp://bad.com"),
        ]
        results = batch_check_urls(entries)
        for u, cr in results.items():
            assert cr.valid is False
            assert "安全拦截" in cr.detail or "禁止" in cr.detail or "不允许" in cr.detail

    def test_mixed_safe_and_unsafe(self, mock_requests_head_ok):
        entries = [
            UrlEntry("S1", 2, 0, "A", "https://safe.com"),
            UrlEntry("S1", 3, 0, "A", "http://192.168.1.1"),
        ]
        results = batch_check_urls(entries)
        assert results["https://safe.com"].valid is True
        assert results["http://192.168.1.1"].valid is False
