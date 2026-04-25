# -*- coding: utf-8 -*-
import os, sys, io, traceback, json, re, webbrowser, threading
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as mb
import tkinter.filedialog as fd
import tkinter.simpledialog as sd
from tkinter.scrolledtext import ScrolledText
from tkinter import colorchooser
import sqlite3
from PIL import Image, ImageTk

# 确保项目根目录在 sys.path 中（PyInstaller 兼容）
_THIS_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _THIS_DIR not in sys.path:
    sys.path.insert(0, _THIS_DIR)

from config import *
import tender_lib_db2 as db
import extractors
import docx_inserter as inserter

class CompareDlg(tk.Toplevel):
    def __init__(self, parent, current, version_id):
        super().__init__(parent)
        self.title(f"版本对比 - v{version_id}")
        self.geometry("900x600")
        self.grab_set()

        ver = None
        for v in db.get_versions(current["id"]):
            if str(v["id"]) == str(version_id):
                ver = v
                break
        if not ver:
            mb.showerror("错误", "版本记录不存在")
            self.destroy()
            return

        old_text = ver.get("raw_text", "") or ""
        new_text = current.get("raw_text", "") or ""

        # 计算简单 diff
        old_lines = old_text.splitlines()
        new_lines = new_text.splitlines()

        f = tk.Frame(self, padx=10, pady=8)
        f.pack(fill=tk.BOTH, expand=True)

        info = tk.Label(f, text=f"当前版本 vs 历史版本 v{ver['version_no']}（{ver['created_at'][:16]}）",
                       font=FONT_HEAD, anchor="w", bg=C_PANEL)
        info.pack(fill=tk.X, pady=0)

        pan = tk.Frame(f)
        pan.pack(fill=tk.BOTH, expand=True)

        # 左侧：历史版本
        tk.Label(pan, text=f"历史版本 v{ver['version_no']}（删除内容标红）",
                 font=FONT_SMALL, anchor="w", bg="#FFF0F0").pack(fill=tk.X)
        left_fr = tk.Frame(pan, bd=1, relief=tk.SOLID, bg="#FFF0F0")
        left_fr.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=0)
        lt = ScrolledText(left_fr, wrap=tk.WORD, font=FONT_MONO, bg="#FFF8F8",
                           bd=0, padx=8, pady=6)
        lt.pack(fill=tk.BOTH, expand=True)
        lt.tag_configure("del", foreground=C_DANGER, background="#FFE0E0")
        for line in old_lines:
            if line.strip() and line.strip() not in new_lines:
                lt.insert(tk.END, line + "\n", "del")
            else:
                lt.insert(tk.END, line + "\n")
        lt.config(state="disabled")

        # 右侧：当前版本
        tk.Label(pan, text="当前版本（新增内容标绿）",
                 font=FONT_SMALL, anchor="w", bg="#F0FFF0").pack(fill=tk.X)
        right_fr = tk.Frame(pan, bd=1, relief=tk.SOLID, bg="#F0FFF0")
        right_fr.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)
        rt = ScrolledText(right_fr, wrap=tk.WORD, font=FONT_MONO, bg="#F8FFF8",
                           bd=0, padx=8, pady=6)
        rt.pack(fill=tk.BOTH, expand=True)
        rt.tag_configure("add", foreground="#1A7A1A", background="#E0FFE0")
        for line in new_lines:
            if line.strip() and line.strip() not in old_lines:
                rt.insert(tk.END, line + "\n", "add")
            else:
                rt.insert(tk.END, line + "\n")
        rt.config(state="disabled")

        bf = tk.Frame(f)
        bf.pack(fill=tk.X, pady=6)
        tk.Button(bf, text="恢复到此版本", font=FONT_BODY,
                 bg=C_ACCENT2, fg="white", cursor="hand2",
                 command=lambda: self._restore(current["id"], ver)).pack(side=tk.LEFT, padx=4)
        tk.Button(bf, text="关闭", font=FONT_BODY, command=self.destroy).pack(side=tk.LEFT)

    def _restore(self, eid, ver):
        if mb.askyesno("确认", f"确定恢复到 v{ver['version_no']}？\n当前内容将被覆盖。"):
            rid = db.restore_version(ver["id"])
            if rid:
                mb.showinfo("完成", f"已恢复到 v{ver['version_no']}")
                self.destroy()


# ══════════════════════════════════════════════════════════════════
#  对话框：模板管理
# ══════════════════════════════════════════════════════════════════


# ══════════════════════════════════════════════════════════════════
# 导入预览校对对话框（核心新功能）
# ══════════════════════════════════════════════════════════════════


