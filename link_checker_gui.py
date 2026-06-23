#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel链接有效性检测工具 - GUI版 (企业级)
基于 link_checker_core 核心模块。
"""

import os
import sys
import io
import threading

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from link_checker_core import (
    Config, ProcessResult,
    process_excel, generate_output_path,
    log,
)


class LinkCheckerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Excel 链接有效性检测工具 — 企业版")
        self.root.geometry("820x640")
        self.root.minsize(720, 500)

        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - 820) // 2
        y = (sh - 640) // 2
        self.root.geometry(f"820x640+{x}+{y}")

        self.input_path = tk.StringVar()
        self.is_running = False
        self.result: ProcessResult | None = None

        self._build_ui()

    # ---------- UI ----------
    def _build_ui(self):
        # 顶部 — 文件选择
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill=tk.X)

        ttk.Label(top, text="选择 Excel 文件：", font=("Microsoft YaHei", 10)).pack(side=tk.LEFT)

        self.file_entry = ttk.Entry(top, textvariable=self.input_path, font=("Consolas", 10))
        self.file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)

        ttk.Button(top, text="📁 浏览", command=self._browse).pack(side=tk.LEFT)
        self.start_btn = ttk.Button(top, text="▶ 开始检测", command=self._start)
        self.start_btn.pack(side=tk.LEFT, padx=(8, 0))

        # 进度
        # 进度
        prog = ttk.LabelFrame(self.root, text="检测进度", padding=10)
        prog.pack(fill=tk.X, padx=10, pady=(0, 10))

        self.progress_bar = ttk.Progressbar(prog, mode="determinate")
        self.progress_bar.pack(fill=tk.X)

        self.progress_label = ttk.Label(prog, text="等待选择文件...", font=("Microsoft YaHei", 9))
        self.progress_label.pack(fill=tk.X, pady=(5, 0))

        self.detail_label = ttk.Label(prog, text="", font=("Consolas", 8), foreground="gray")
        self.detail_label.pack(fill=tk.X)

        # 结果
        res = ttk.LabelFrame(self.root, text="检测结果", padding=10)
        res.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        cols = ("sheet", "row", "col", "url", "status", "detail", "retries")
        self.tree = ttk.Treeview(res, columns=cols, show="headings", height=14)
        self.tree.heading("sheet", text="Sheet")
        self.tree.heading("row", text="行")
        self.tree.heading("col", text="列名")
        self.tree.heading("url", text="链接")
        self.tree.heading("status", text="状态")
        self.tree.heading("detail", text="详情")
        self.tree.heading("retries", text="重试")

        self.tree.column("sheet", width=90)
        self.tree.column("row", width=40, anchor=tk.CENTER)
        self.tree.column("col", width=70)
        self.tree.column("url", width=300)
        self.tree.column("status", width=60, anchor=tk.CENTER)
        self.tree.column("detail", width=150)
        self.tree.column("retries", width=45, anchor=tk.CENTER)

        sb = ttk.Scrollbar(res, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree.tag_configure("valid", foreground="#006100", background="#C6EFCE")
        self.tree.tag_configure("invalid", foreground="#9C0006", background="#FFC7CE")

        # 底部
        bot = ttk.Frame(self.root, padding=10)
        bot.pack(fill=tk.X)

        self.stats_label = ttk.Label(bot, text="", font=("Microsoft YaHei", 10))
        self.stats_label.pack(side=tk.LEFT)

        self.save_full_btn = ttk.Button(
            bot, text="保存完整版", command=self._save_full, state=tk.DISABLED
        )
        self.save_full_btn.pack(side=tk.RIGHT, padx=4)
        self.save_light_btn = ttk.Button(
            bot, text="保存仅标黄版", command=self._save_light, state=tk.DISABLED
        )
        self.save_light_btn.pack(side=tk.RIGHT)

    # ---------- Actions ----------
    def _browse(self):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if path:
            self.input_path.set(path)

    def _start(self):
        path = self.input_path.get().strip()
        if not path:
            messagebox.showwarning("提示", "请先选择一个 Excel 文件。")
            return
        if not os.path.exists(path):
            messagebox.showerror("错误", f"文件不存在：\n{path}")
            return
        if self.is_running:
            return

        self.is_running = True
        self._clear()
        self.progress_label.config(text="正在分析文件...")
        self.progress_bar["value"] = 0

        t = threading.Thread(target=self._run, args=(path,), daemon=True)
        t.start()

    def _run(self, path: str):
        try:
            output = generate_output_path(path)
            self.result = process_excel(path, output, self._on_progress,
                                       output_full=False, output_light=False)
            self.root.after(0, self._on_finish, self.result)
        except Exception as e:
            self.root.after(0, self._on_error, str(e))

    def _on_progress(self, completed: int, total: int, result):
        def update():
            self.progress_bar["maximum"] = total
            self.progress_bar["value"] = completed
            if result is not None:
                pct = int(completed / total * 100) if total else 0
                icon = "✅" if result.valid else "❌"
                self.progress_label.config(
                    text=f"[{completed}/{total}] {pct}%  {icon} {result.url[:80]}"
                )
                self.detail_label.config(text=f"详情: {result.detail}  |  重试: {result.retries}次")
        self.root.after(0, update)

    def _on_finish(self, result: ProcessResult):
        self.is_running = False
        for r in result.results:
            tag = "valid" if r["valid"] else "invalid"
            self.tree.insert("", tk.END, values=(
                r["sheet"], r["row"], r["col"], r["url"],
                "有效" if r["valid"] else "无效",
                r["detail"], r["retries"],
            ), tags=(tag,))

        self.stats_label.config(
            text=f"总链接: {result.total_urls}  |  有效: {result.valid_count}  "
                 f"|  无效: {result.invalid_count}  "
                 + (f"|  跳过: {result.skipped_count}" if result.skipped_count else "")
        )
        self.progress_label.config(text="检测完成，请选择要保存的输出格式")
        self.progress_bar["value"] = self.progress_bar["maximum"]
        self.detail_label.config(text="")
        self.save_full_btn.config(state=tk.NORMAL)
        self.save_light_btn.config(state=tk.NORMAL)
    def _on_error(self, msg: str):
        self.is_running = False
        messagebox.showerror("检测失败", f"发生错误：\n{msg}")
        self.progress_label.config(text="检测失败")
        self.progress_bar["value"] = 0
        self.detail_label.config(text="")

    def _clear(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.stats_label.config(text="")
        self.progress_label.config(text="")
        self.detail_label.config(text="")
        self.save_full_btn.config(text="保存完整版", command=self._save_full, state=tk.DISABLED)
        self.save_light_btn.config(text="保存仅标黄版", command=self._save_light, state=tk.DISABLED)
        self.result = None
    def _save_full(self):
        if not self.result:
            return
        from link_checker_core import save_full_result
        path = save_full_result(self.input_path.get(), self.result)
        self._saved_full_path = path
        self.save_full_btn.config(text="打开完整版", command=self._open_full, state=tk.NORMAL)

    def _open_full(self):
        if hasattr(self, '_saved_full_path') and os.path.exists(self._saved_full_path):
            os.startfile(self._saved_full_path)

    def _save_light(self):
        if not self.result:
            return
        from link_checker_core import save_light_result
        path = save_light_result(self.input_path.get(), self.result)
        self._saved_light_path = path
        self.save_light_btn.config(text="打开仅标黄版", command=self._open_light, state=tk.NORMAL)

    def _open_light(self):
        if hasattr(self, '_saved_light_path') and os.path.exists(self._saved_light_path):
            os.startfile(self._saved_light_path)


def main():
    root = tk.Tk()
    LinkCheckerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
