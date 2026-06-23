# Excel 链接有效性检测工具

读取 Excel 文件中的链接，自动检测每个链接是否可访问，无效链接所在行标黄，输出结果文件。


## AI 参与说明

本项目完全由 AI 辅助生成，包括但不限于：

- 全部代码（核心逻辑、CLI、GUI）
- 单元测试（116 个）
- 项目文档（README）
- 项目结构与工程配置

无任何代码由人工手写。
## 项目结构

```
├── link_checker_core.py    # 核心模块（日志、配置、检测、Excel处理）
├── link_checker.py         # 命令行入口
├── link_checker_gui.py     # 图形界面入口（tkinter，推荐）
├── 示例链接.xlsx            # 测试用示例文件
├── tests/                  # 单元测试（116 个）
├── setup.bat               # Windows 一键安装脚本（推荐零基础用户使用）
├── requirements.txt        # 运行依赖
├── requirements-dev.txt    # 开发依赖（含 pytest）
├── .gitignore
└── README.md
```

## 环境要求

| 依赖 | 版本要求 |
|------|---------|
| Python | ≥ 3.9 |
| requests | ≥ 2.28 |
| openpyxl | ≥ 3.1 |
| pandas | ≥ 1.5 |
| tkinter | 通常 Python 自带 |

安装依赖：

```bash
pip install requests openpyxl pandas
```

## 快速开始（零基础）

### 1. 安装 Python

前往 [python.org](https://www.python.org/downloads/) 下载安装，**安装时务必勾选 "Add Python to PATH"**。

### 2. 运行安装脚本

双击项目文件夹中的 **`setup.bat`**，脚本会自动：

1. 创建独立的 Python 虚拟环境（不影响系统）
2. 安装所需依赖
3. 启动程序

### 3. 使用

1. 点击 **📁 浏览** 选择 Excel 文件
2. 点击 **▶ 开始检测**
3. 在底部点击「保存完整版」或「保存仅标黄版」
4. 按钮变为「打开」后可点击直接查看结果

> 结果文件保存在输入文件同目录下的 **result/** 文件夹中。

### 以后怎么启动

双击 `setup.bat` 即可，虚拟环境已存在时会跳过安装直接启动。

### 高级用户：手动运行

```powershell
# GUI 版（推荐）
python link_checker_gui.py

# 命令行版
python link_checker.py "你的文件.xlsx"
```

## 功能特性

### 链接检测
- 自动扫描 Excel 所有 Sheet 中的 `http://` / `https://` 链接
- 并发检测（默认 10 线程），HEAD 优先，失败后 GET 重试
- 自动重试（默认 2 次），指数退避
- 区分超时、DNS失败、SSL错误等，提供中文详情

### 安全防护
- 拦截内网地址（127.0.0.0/8、10.0.0.0/8、172.16.0.0/12、192.168.0.0/16）
- 拦截 file://、ftp:// 等危险协议
- 拦截 .local、.internal 等内部域名

### 输出文件

每次检测输出 **两个文件**（可在 CLI/GUI 中选择性输出）：

| 文件 | 命名格式 | 说明 |
|------|----------|------|
| **完整版** | result/原文件名_链接检测结果_时间戳.xlsx | 无效链接整行标黄 + 新增"链接检测结果"汇总 Sheet |
| **仅标黄版** | result/原文件名_链接检测_仅标黄_时间戳.xlsx | 与原文件基本一致，仅无效链接整行标黄 |

> 两个文件均输出到输入文件所在目录下的 **result/** 子文件夹中，自动创建。

#### CLI 参数

```powershell
# 默认输出两个文件
python link_checker.py "文件.xlsx"

# 仅输出完整版
python link_checker.py "文件.xlsx" --no-light

# 仅输出仅标黄版
python link_checker.py "文件.xlsx" --no-full

# 都不输出（只看终端结果）
python link_checker.py "文件.xlsx" --no-full --no-light
```

#### GUI 选项

检测完成后，底部显示两个按钮：

- **保存完整版** — 输出汇总+标黄的完整结果文件
- **保存仅标黄版** — 输出仅标黄的结果文件

点击保存后按钮变为 **打开完整版/打开仅标黄版**，可直接打开对应文件。可只保存其中一个，或两个都保存。

### 输出特性
- 文件名带时间戳，**永不覆盖原文件**
- 无效链接所在 **整行** 标黄底色
- 有效链接不做任何改动
- 完整版新增「链接检测结果」Sheet：检测时间、统计汇总、逐条详情、重试次数
- 原格式、列宽、字体全部保留

### 日志
- 控制台输出 INFO 级别（时间 + 摘要）
- 文件记录 DEBUG 级别（`link_checker.log`，自动轮转 5MB×3）

## 配置

所有可调参数集中在 `link_checker_core.py` 的 `Config` 数据类：

| 参数 | 默认值 | 说明 |
|------|--------|------|
| TIMEOUT | 10 | 请求超时（秒） |
| MAX_WORKERS | 10 | 并发线程数 |
| MAX_RETRIES | 2 | 失败重试次数 |
| RETRY_BACKOFF | 1.5 | 重试退避系数 |
| MAX_REDIRECTS | 5 | 最大重定向次数 |
| BATCH_SIZE | 200 | 大批量时分批大小 |

## 常见问题

**Q: 为什么所有链接都显示"无效"？**

A: 当前运行环境可能有网络限制（如公司防火墙、沙箱）。请在正常网络环境的机器上运行。

**Q: 检测到一半卡住不动？**

A: 部分网站连接后不响应，会在超时后才继续。可适当调小 `TIMEOUT` 值（如 5 秒）。
