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

_THIS_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _THIS_DIR not in sys.path:
    sys.path.insert(0, _THIS_DIR)

from config import *
import tender_lib_db2 as db
import extractors
import docx_inserter as inserter


class InsertDlg(tk.Toplevel):
    def __init__(self, parent, entries):
        super().__init__(parent)
        self.title("选择要插入的条目")
        self.entries = entries
        self.selected_ids = []
        self._checked = set()   # track checked iids
        self.grab_set()
        self._build()

    def _build(self):
        f = tk.Frame(self, padx=14, pady=12)
        f.pack(fill=tk.BOTH, expand=True)
        tk.Label(f, text="勾选要插入的条目（双击预览）：",
                 font=FONT_BODY, anchor="w").pack(anchor="w")
        tf = tk.Frame(f)
        tf.pack(fill=tk.BOTH, expand=True, pady=8)
        # Use #0 (text col) for checkbox — shown with show="tree headings"
        self.tree = ttk.Treeview(tf, columns=["标题", "分类", "类型"],
                                show="tree headings", height=16)
        self.tree.heading("#0", text="✓")
        self.tree.column("#0", width=40, anchor="center")
        self.tree.heading("标题", text="标题")
        self.tree.column("标题", width=280)
        self.tree.heading("分类", text="分类")
        self.tree.column("分类", width=100)
        self.tree.heading("类型", text="类型")
        self.tree.column("类型", width=70)
        vsb = ttk.Scrollbar(tf, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<Button-1>", self._on_click)
        self.tree.bind("<Double-Button-1>", self._on_preview)
        cats = {c["id"]: c["name"] for c in db.list_categories()}
        for e in self.entries:
            self.tree.insert("", tk.END, iid=str(e["id"]),
                            text="[ ]",
                            values=(e["title"][:35],
                                    cats.get(e["category_id"], ""),
                                    e["content_type"]))

        bf = tk.Frame(f)
        bf.pack()
        tk.Button(bf, text="全选", font=FONT_BODY, width=8,
                 command=self._select_all).pack(side=tk.LEFT, padx=4)
        tk.Button(bf, text="取消全选", font=FONT_BODY, width=8,
                 command=self._deselect_all).pack(side=tk.LEFT, padx=4)
        tk.Button(bf, text="确定插入", font=FONT_BODY, width=12,
                 bg=C_ACCENT, fg="white", cursor="hand2",
                 command=self._ok).pack(side=tk.LEFT, padx=12)
        tk.Button(bf, text="取消", font=FONT_BODY, width=8,
                 command=self.destroy).pack(side=tk.LEFT)

    def _on_click(self, event=None):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        if col != "#0":
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        cur = self.tree.item(iid, "text")
        new = "[x]" if cur != "[x]" else "[ ]"
        self.tree.item(iid, text=new)
        eid = int(iid)
        if new == "[x]":
            self._checked.add(eid)
        else:
            self._checked.discard(eid)

    def _on_preview(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        e = next((e for e in self.entries if str(e["id"]) == sel[0]), None)
        if not e:
            return
        win = tk.Toplevel(self)
        win.title(e["title"][:60])
        win.geometry("600x400")
        t = ScrolledText(win, wrap=tk.WORD, font=FONT_BODY)
        t.insert("1.0", e.get("raw_text", "")[:2000] or "（无内容）")
        t.config(state="disabled")
        t.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

    def _select_all(self):
        for iid in self.tree.get_children():
            self.tree.item(iid, text="[x]")
        self._checked = {int(iid) for iid in self.tree.get_children()}

    def _deselect_all(self):
        for iid in self.tree.get_children():
            self.tree.item(iid, text="[ ]")
        self._checked.clear()

    def _ok(self):
        self.selected_ids = sorted(self._checked)
        self.destroy()


# ══════════════════════════════════════════════════════════════════
#  对话框：版本对比
# ══════════════════════════════════════════════════════════════════


class CompareDlg(tk.Toplevel):
    """版本对比对话框"""
    def __init__(self, parent, versions):
        super().__init__(parent)
        self.title("版本对比")
        self.geometry("900x600")
        self.grab_set()
        self.versions = versions
        self._build()

    def _build(self):
        f = tk.Frame(self, padx=10, pady=10)
        f.pack(fill=tk.BOTH, expand=True)
        tk.Label(f, text="← 旧版本          新版本 →",
                 font=FONT_HEAD).pack(pady=4)
        paned = tk.PanedWindow(f, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)
        for side, idx in [("left", 0), ("right", 1)]:
            fr = tk.Frame(paned, bg=C_PANEL)
            paned.add(fr)
            lv = tk.Listbox(fr, font=FONT_BODY, bg=C_PANEL)
            lv.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
            tk.Scrollbar(fr, command=lv.yview).pack(side=tk.RIGHT, fill=tk.Y)
            lv.configure(yscrollcommand=lambda s, l=lv: s.set(*l.yview))
            for v in self.versions:
                label = f"v{v['version_no']}  {v.get('raw_text','')[:30]}"
                lv.insert(tk.END, label)
            setattr(self, f"lv_{side}", lv)

        bf = tk.Frame(f)
        bf.pack(pady=6)
        tk.Button(bf, text="对比", command=self._do_compare,
                 font=FONT_BODY, bg=C_ACCENT, fg="white",
                 cursor="hand2").pack(side=tk.LEFT, padx=8)
        tk.Button(bf, text="关闭", command=self.destroy,
                 font=FONT_BODY).pack(side=tk.LEFT)

    def _do_compare(self):
        lidx = self.lv_left.curselection()
        ridx = self.lv_right.curselection()
        if not lidx or not ridx:
            mb.showinfo("提示", "请在两侧各选一个版本", parent=self)
            return
        vl = self.versions[lidx[0]]
        vr = self.versions[ridx[0]]
        old = vl.get("raw_text", "") or "（空）"
        new = vr.get("raw_text", "") or "（空）"
        w = tk.Toplevel(self)
        w.title(f"v{vl['version_no']} vs v{vr['version_no']}")
        w.geometry("800x500")
        paned = tk.PanedWindow(w, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)
        for txt in [old, new]:
            fr = tk.Frame(paned)
            paned.add(fr)
            t = ScrolledText(fr, wrap=tk.WORD, font=FONT_BODY)
            t.insert("1.0", txt)
            t.config(state="disabled")
            t.pack(fill=tk.BOTH, expand=True)
