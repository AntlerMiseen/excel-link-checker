@echo off
chcp 65001 >nul
title 链接检测工具 - 安装

echo ========================================
echo   链接检测工具 - 一键安装
echo ========================================
echo.

:: 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python
    echo 下载地址：https://www.python.org/downloads/
    echo 安装时请勾选 "Add Python to PATH"
    pause
    exit /b 1
)
echo [√] Python 已就绪

:: 创建虚拟环境
if not exist "venv" (
    echo [*] 正在创建虚拟环境...
    python -m venv venv
    echo [√] 虚拟环境创建完成
) else (
    echo [√] 虚拟环境已存在
)

:: 激活虚拟环境并安装依赖
echo [*] 正在安装依赖...
call venv\Scriptsctivate.bat
pip install requests openpyxl pandas -q
if errorlevel 1 (
    echo [错误] 依赖安装失败，请检查网络后重试
    pause
    exit /b 1
)
echo [√] 依赖安装完成

echo.
echo ========================================
echo   安装完成，即将启动程序...
echo ========================================
python link_checker_gui.py
pause
