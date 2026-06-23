# -*- coding: utf-8 -*-
"""pytest 共享夹具"""

import os
import sys
import tempfile
import pytest
from unittest.mock import MagicMock, patch
from openpyxl import Workbook

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


@pytest.fixture
def sample_workbook():
    """创建一个包含链接的测试 Workbook"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["名称", "链接", "备注"])
    ws.append(["谷歌", "https://www.google.com", "搜索引擎"])
    ws.append(["百度", "https://www.baidu.com", "中文搜索"])
    ws.append(["内网", "http://10.0.0.1/admin", "内部地址"])
    ws.append(["FTP", "ftp://files.example.com", "危险协议"])
    ws.append(["纯文本", "这不是链接", "无URL"])
    ws.append(["多链接", "https://a.com 和 https://b.com", "两个"])
    ws.append(["空值", None, None])
    ws.append(["数字", 12345, 3.14])
    return wb


@pytest.fixture
def temp_xlsx(sample_workbook):
    """将测试 Workbook 保存为临时 .xlsx 文件"""
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    sample_workbook.save(path)
    yield path
    try:
        os.unlink(path)
    except OSError:
        pass


@pytest.fixture
def mock_requests_head_ok():
    with patch("link_checker_core.requests.Session.head") as mock_head:
        mock_resp = MagicMock()
        mock_resp.status_code = 200
        mock_head.return_value = mock_resp
        yield mock_head


@pytest.fixture
def mock_requests_head_404():
    with patch("link_checker_core.requests.Session.head") as mock_head, \
         patch("link_checker_core.requests.Session.get") as mock_get:
        mock_resp_404 = MagicMock()
        mock_resp_404.status_code = 404
        mock_head.return_value = mock_resp_404
        mock_get.return_value = mock_resp_404
        yield mock_head, mock_get


@pytest.fixture
def mock_requests_head_fail_get_ok():
    with patch("link_checker_core.requests.Session.head") as mock_head, \
         patch("link_checker_core.requests.Session.get") as mock_get:
        import requests as req
        mock_head.side_effect = req.exceptions.ConnectionError("refused")
        mock_resp_200 = MagicMock()
        mock_resp_200.status_code = 200
        mock_get.return_value = mock_resp_200
        yield mock_head, mock_get


@pytest.fixture
def mock_requests_timeout():
    with patch("link_checker_core.requests.Session.head") as mock_head, \
         patch("link_checker_core.requests.Session.get") as mock_get:
        import requests as req
        mock_head.side_effect = req.exceptions.Timeout()
        mock_get.side_effect = req.exceptions.Timeout()
        yield mock_head, mock_get


@pytest.fixture
def mock_requests_ssl_error():
    with patch("link_checker_core.requests.Session.head") as mock_head:
        import requests as req
        mock_head.side_effect = req.exceptions.SSLError()
        yield mock_head


@pytest.fixture
def mock_requests_too_many_redirects():
    with patch("link_checker_core.requests.Session.head") as mock_head:
        import requests as req
        mock_head.side_effect = req.exceptions.TooManyRedirects()
        yield mock_head


@pytest.fixture
def reset_thread_local():
    from link_checker_core import _thread_local
    if hasattr(_thread_local, "session"):
        del _thread_local.session
    yield
    if hasattr(_thread_local, "session"):
        del _thread_local.session
