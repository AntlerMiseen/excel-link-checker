#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel链接有效性检测工具 - 命令行版 (企业级)
基于 link_checker_core 核心模块。
"""

import sys
import io
import os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

from link_checker_core import (
    Config, ProcessResult,
    process_excel, generate_output_path, hash_file,
    log, get_logger,
)


def _progress(completed: int, total: int, result):
    """控制台进度回调"""
    if result is None:
        return
    pct = int(completed / total * 100)
    icon = "✅" if result.valid else "❌"
    print(f"\r  [{completed}/{total}] {pct:3d}%  {icon} {result.url[:70]}", end="", flush=True)


def main():
    output_full = True
    output_light = True
    args = sys.argv[1:]
    input_path = None
    for a in args:
        if a in ('--no-full',):
            output_full = False
        elif a in ('--no-light',):
            output_light = False
        elif not a.startswith('--'):
            input_path = a.strip().strip("'\"")
    if input_path is None:
        input_path = input("请输入Excel文件路径: ").strip().strip("'\"")

    if not os.path.exists(input_path):
        log.error(f"文件不存在: {input_path}")
        sys.exit(1)

    # 计算原文件哈希（防篡改审计）
    try:
        original_hash = hash_file(input_path)
        log.debug(f"原文件SHA256: {original_hash}")
    except Exception:
        original_hash = ""

    output_path = generate_output_path(input_path)

    try:
        result = process_excel(input_path, output_path, _progress, output_full=output_full, output_light=output_light)
    except FileNotFoundError as e:
        log.error(str(e))
        sys.exit(1)
    except PermissionError as e:
        log.error(f"文件权限不足: {e}")
        sys.exit(1)
    except Exception as e:
        log.exception(f"处理失败: {e}")
        sys.exit(1)

    print()  # 换行
    print(f"\n{'='*50}")
    print(f"📂 输入: {result.input_path}")
    print(f"完整版: {result.output_path}")
    print(f"仅标黄: {result.light_output_path}")
    print(f"🔗 总链接: {result.total_urls}")
    print(f"✅ 有效:   {result.valid_count}")
    print(f"❌ 无效:   {result.invalid_count}")
    if result.skipped_count:
        print(f"⚠️  跳过:   {result.skipped_count}")
    print(f"{'='*50}")

    if result.errors:
        print(f"\n⚠️  检测到 {len(result.errors)} 个错误:")
        for err in result.errors[:5]:
            print(f"   - {err}")


if __name__ == "__main__":
    main()
