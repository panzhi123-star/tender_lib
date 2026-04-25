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

class TemplateDlg(tk.Toplevel):
    def __init__(self, parent, app, mode=None):
        super().__init__(parent)
        self.app = app
        self.mode = mode
        self.title("模板管理")
        self.geometry("700x520")
        self.grab_set()
        self._build()

    def _build(self):
        f = tk.Frame(self, padx=14, pady=12)
        f.pack(fill=tk.BOTH, expand=True)

        tk.Label(f, text="标书模板库", font=FONT_HEAD, anchor="w").pack(anchor="w")
        tf = tk.Frame(f)
        tf.pack(fill=tk.BOTH, expand=True, pady=8)

        self.tpl_tree = ttk.Treeview(tf, columns=["name","cat","desc","use"],
                                       show="headings", height=12)
        for col, w, h in [("name",200,"模板名称"),("cat",100,"分类"),
                           ("desc",240,"描述"),("use",60,"使用次数")]:
            self.tpl_tree.heading(col, text=h)
            self.tpl_tree.column(col, width=w, anchor="w")
        vsb = ttk.Scrollbar(tf, orient=tk.VERTICAL, command=self.tpl_tree.yview)
        self.tpl_tree.configure(yscrollcommand=vsb.set)
        self.tpl_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self._fill()

        bf = tk.Frame(f)
        bf.pack(fill=tk.X, pady=6)
        for txt, cmd in [
            ("+ 新建模板", self._new),
            ("编辑模板",   self._edit),
            ("删除",       self._delete),
            ("刷新",       self._fill),
        ]:
            tk.Button(bf, text=txt, command=cmd, font=FONT_BODY,
                     bd=1, cursor="hand2", padx=10, pady=3).pack(side=tk.LEFT, padx=4)
        if self.mode == "new":
            self._new()

    def _fill(self):
        for row in self.tpl_tree.get_children():
            self.tpl_tree.delete(row)
        for t in db.list_templates():
            self.tpl_tree.insert("", tk.END, iid=str(t["id"]), values=(
                t["name"], t["category"], t["description"], str(t["use_count"])))

    def _new(self):
        dlg = TemplateEditDlg(self)
        if dlg.result:
            db.create_template(**dlg.result)
            self._fill()
            self.app._update_stats()

    def _edit(self):
        sel = self.tpl_tree.selection()
        if not sel:
            mb.showinfo("提示", "请先选择要编辑的模板")
            return
        tid = int(sel[0])
        t = db.get_template(tid)
        dlg = TemplateEditDlg(self, t)
        if dlg.result:
            db.update_template(tid, **dlg.result)
            self._fill()

    def _delete(self):
        sel = self.tpl_tree.selection()
        if not sel:
            return
        tid = int(sel[0])
        if mb.askyesno("确认", "删除此模板？"):
            db.delete_template(tid)
            self._fill()
            self.app._update_stats()


class TemplateEditDlg(tk.Toplevel):
    def __init__(self, parent, template=None):
        super().__init__(parent)
        self.result = None
        self.template = template
        self.title("编辑模板")
        self.geometry("500x360")
        self.grab_set()
        self._build()

    def _build(self):
        f = tk.Frame(self, padx=16, pady=12)
        f.pack(fill=tk.BOTH, expand=True)
        self.name_var = tk.StringVar(value=self.template.get("name", "") if self.template else "")
        self.category_var = tk.StringVar(value=self.template.get("category", "") if self.template else "")
        self.desc_var = tk.StringVar(value=self.template.get("description", "") if self.template else "")

        for i, (lbl, var, w) in enumerate([
            ("模板名称：", self.name_var, 36), ("分类：", self.category_var, 20),
            ("描述：", self.desc_var, 40),
        ]):
            tk.Label(f, text=lbl, font=FONT_BODY, anchor="w").grid(
                row=i, column=0, sticky="nw", pady=6)
            e = tk.Entry(f, textvariable=var, font=FONT_BODY, width=w)
            e.grid(row=i, column=1, sticky="ew", pady=6)

        tk.Label(f, text="内容：", font=FONT_BODY, anchor="w").grid(
            row=3, column=0, sticky="nw", pady=6)
        self.content_text = ScrolledText(f, wrap=tk.WORD, font=FONT_BODY, width=55, height=10)
        self.content_text.insert("1.0", self.template.get("content", "") if self.template else "")
        self.content_text.grid(row=3, column=1, sticky="nsew", pady=6)

        bf = tk.Frame(f)
        bf.grid(row=4, column=0, columnspan=2, pady=8)
        tk.Button(bf, text="取消", command=self.destroy, font=FONT_BODY,
                 width=10).pack(side=tk.RIGHT, padx=6)
        tk.Button(bf, text="保存", command=self._save, font=FONT_BODY,
                 width=10, bg=C_ACCENT, fg="white", cursor="hand2").pack(side=tk.RIGHT)
        f.columnconfigure(1, weight=1)
        f.rowconfigure(3, weight=1)

    def _save(self):
        name = self.name_var.get().strip()
        category = self.category_var.get().strip()
        description = self.desc_var.get().strip()
        content = self.content_text.get("1.0", tk.END).strip()
        if not name:
            mb.showwarning("提示", "请输入模板名称", parent=self)
            return
        self.result = {
            "name": name,
            "category": category,
            "description": description,
            "content": content,
        }
        self.destroy()


class TemplateInsertDlg(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("从模板插入条目")
        self.result = None
        self.grab_set()
        self._build()

    def _build(self):
        f = tk.Frame(self, padx=14, pady=12)
        f.pack(fill=tk.BOTH, expand=True)
        tk.Label(f, text="选择模板，勾选要插入的条目片段：",
                 font=FONT_BODY, anchor="w").pack(anchor="w")

        tf = tk.Frame(f)
        tf.pack(fill=tk.BOTH, expand=True, pady=8)
        cols = ["", "模板", "分类", "条目", "类型"]
        self.tree = ttk.Treeview(tf, columns=cols, show="headings", height=14)
        for col, w in [("", 30), ("模板", 150), ("分类", 80), ("条目", 250), ("类型", 60)]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="w")
        vsb = ttk.Scrollbar(tf, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<Button-1>", self._on_click)

        for t in db.list_templates():
            self.tree.insert("", tk.END, iid=f"tpl_{t['id']}",
                             values=("[ ]", t["name"], t["category"], "", ""),
                             tags=("tpl", str(t["id"])))

        bf = tk.Frame(f)
        bf.pack()
        tk.Button(bf, text="全选", font=FONT_BODY, width=8,
                 command=self._select_all).pack(side=tk.LEFT, padx=4)
        tk.Button(bf, text="确定插入", font=FONT_BODY, width=12,
                 bg=C_ACCENT, fg="white", cursor="hand2", command=self._ok).pack(side=tk.LEFT, padx=12)
        tk.Button(bf, text="取消", font=FONT_BODY, width=8, command=self.destroy).pack(side=tk.LEFT)

    def _on_click(self, event=None):
        """Toggle checkbox on click in first column."""
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        if col != "#1":
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        vals = self.tree.item(iid)["values"]
        self.tree.item(iid, values=("[x]" if vals[0] == "[ ]" else "[ ]",) + vals[1:])

    def _select_all(self):
        for iid in self.tree.get_children():
            self.tree.item(iid, values=("[x]",) + self.tree.item(iid)["values"][1:])

    def _ok(self):
        sel_ids = []
        for i in self.tree.get_children():
            if self.tree.item(i)["values"][0] != "[x]":
                continue
            tags = self.tree.item(i)["tags"]
            if not tags or tags[0] != "tpl":
                continue
            sel_ids.append(int(tags[1]))
        if not sel_ids:
            mb.showinfo("提示", "请至少选择一个条目")
            return
        self.result = []
        for tid in sel_ids:
            tpl = db.get_template(tid)
            if tpl:
                self.result.append({
                    "title": tpl["name"],
                    "raw_text": tpl.get("content", ""),
                    "attachments": [],
                })
        self.destroy()


# ══════════════════════════════════════════════════════════════════
#  对话框：批量管理
# ══════════════════════════════════════════════════════════════════


