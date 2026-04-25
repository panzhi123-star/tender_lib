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

class EntryDlg(tk.Toplevel):
    def __init__(self, parent, title, entry):
        super().__init__(parent)
        self.title(title)
        self.result = None
        self.entry = entry
        self.grab_set()
        self._build()

    def _build(self):
        cats = db.list_categories()
        cat_map = {c["id"]: c["name"] for c in cats}
        all_cats = [{"id": None, "name": "（未分类）"}] + cats

        f = tk.Frame(self, padx=18, pady=14)
        f.pack(fill=tk.BOTH, expand=True)

        # 标题
        tk.Label(f, text="标题：", font=FONT_BODY, anchor="w").grid(
            row=0, column=0, sticky="nw", pady=6)
        self.title_var = tk.StringVar(value=self.entry.get("title", ""))
        tk.Entry(f, textvariable=self.title_var, font=FONT_BODY, width=46).grid(
            row=0, column=1, columnspan=2, sticky="ew", pady=6)

        # 分类
        tk.Label(f, text="分类：", font=FONT_BODY, anchor="w").grid(
            row=1, column=0, sticky="nw", pady=6)
        self.cat_var = tk.StringVar(
            value=cat_map.get(self.entry.get("category_id"), "（未分类）"))
        self.cat_combo = ttk.Combobox(f, textvariable=self.cat_var,
                     values=[c["name"] for c in all_cats],
                     font=FONT_BODY, width=28, state="readonly")
        self.cat_combo.grid(row=1, column=1, columnspan=2, sticky="ew", pady=6)

        # 标签
        tk.Label(f, text="标签：", font=FONT_BODY, anchor="w").grid(
            row=2, column=0, sticky="nw", pady=6)
        self.tags_var = tk.StringVar(value=self.entry.get("tags", ""))
        tk.Entry(f, textvariable=self.tags_var, font=FONT_BODY, width=46).grid(
            row=2, column=1, columnspan=2, sticky="ew", pady=6)

        # 内容
        tk.Label(f, text="内容：", font=FONT_BODY, anchor="w").grid(
            row=3, column=0, sticky="nw", pady=6)
        self.txt_widget = ScrolledText(f, wrap=tk.WORD, font=FONT_BODY, width=60, height=14)
        self.txt_widget.insert("1.0", self.entry.get("raw_text", "") or "")
        self.txt_widget.grid(row=3, column=1, columnspan=2, sticky="nsew", pady=6)

        bf = tk.Frame(f)
        bf.grid(row=4, column=0, columnspan=3, pady=10)
        tk.Button(bf, text="取消", command=self.destroy,
                 font=FONT_BODY, width=10).pack(side=tk.RIGHT, padx=6)
        tk.Button(bf, text="保存", command=self._save,
                 font=FONT_BODY, width=10, bg=C_ACCENT, fg="white",
                 cursor="hand2").pack(side=tk.RIGHT)
        f.columnconfigure(1, weight=1)
        f.rowconfigure(3, weight=1)

    def _save(self):
        title = self.title_var.get().strip()
        if not title:
            mb.showwarning("提示", "请输入标题", parent=self)
            return
        cats_rev = {c["name"]: c["id"] for c in db.list_categories()}
        # Get edited text content
        raw_text = self.txt_widget.get("1.0", tk.END).strip() if hasattr(self, 'txt_widget') else None
        self.result = {
            "title": title,
            "category_id": cats_rev.get(self.cat_var.get()),
            "tags": self.tags_var.get().strip(),
        }
        if raw_text is not None:
            self.result["raw_text"] = raw_text
        self.destroy()


# ══════════════════════════════════════════════════════════════════
#  对话框：选择条目插入
# ══════════════════════════════════════════════════════════════════


