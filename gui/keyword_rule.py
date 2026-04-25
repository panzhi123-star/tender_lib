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

class KeywordRuleDlg(tk.Toplevel):
    """关键词 → 分类/标签 映射规则管理"""

    def __init__(self, parent, current_rules=None, on_updated=None):
        super().__init__(parent)
        self.title("关键词规则管理")
        self.geometry("750x520")
        self.grab_set()
        self.on_updated = on_updated
        self.current_rules = current_rules or db.list_keyword_rules()
        self._build_ui()
        self._refresh()

    def _build_ui(self):
        f = tk.Frame(self, padx=14, pady=12)
        f.pack(fill=tk.BOTH, expand=True)

        tk.Label(f, text="关键词规则：关键词匹配后自动打标签或分类",
                 font=FONT_HEAD, anchor="w").pack(anchor="w", pady=(0, 8))

        # 列表
        tf = tk.Frame(f)
        tf.pack(fill=tk.BOTH, expand=True, pady=4)
        cols = ["关键词", "操作", "目标值", "优先级", "启用", ""]
        self.kr_tree = ttk.Treeview(tf, columns=cols, show="headings", height=16)
        for col, w in [("关键词", 150), ("操作", 80),
                         ("目标值", 120), ("优先级", 70), ("启用", 50), ("", 50)]:
            self.kr_tree.heading(col, text=col)
            self.kr_tree.column(col, width=w, anchor="w")
        kr_v = ttk.Scrollbar(tf, orient=tk.VERTICAL, command=self.kr_tree.yview)
        self.kr_tree.configure(yscrollcommand=kr_v.set)
        self.kr_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        kr_v.pack(side=tk.RIGHT, fill=tk.Y)
        self.kr_tree.bind("<Double-Button-1>", lambda _: self._edit_rule())

        # 按钮行
        bf = tk.Frame(f)
        bf.pack(fill=tk.X, pady=(8, 0))
        for txt, cmd in [
            ("+ 新建规则", self._new_rule),
            ("编辑选中", self._edit_rule),
            ("删除选中", self._del_rule),
            ("上移", lambda: self._move_sel(-1)),
            ("下移", lambda: self._move_sel(1)),
            ("确定保存", self._save),
        ]:
            b = tk.Button(bf, text=txt, command=cmd, font=FONT_BODY,
                         bd=1, cursor="hand2", padx=10, pady=3)
            b.pack(side=tk.LEFT, padx=3)

    def _refresh(self):
        for row in self.kr_tree.get_children():
            self.kr_tree.delete(row)
        self.current_rules = db.list_keyword_rules()
        for r in self.current_rules:
            enabled = "✔" if r.get("enabled", 1) else "✖"
            action_map = {"category": "分类", "tag": "标签", "label": "标签"}
            at = action_map.get(r.get("action_type", "tag"), r.get("action_type", "tag"))
            self.kr_tree.insert("", tk.END, iid=str(r["id"]), values=(
                r["keyword"], at, r.get("action_value", ""),
                str(r.get("priority", 10)), enabled, ""
            ))

    def _new_rule(self):
        dlg = KeywordRuleEditDlg(self)
        if dlg.result:
            db.create_keyword_rule(**dlg.result)
            self._refresh()
            self._notify()

    def _edit_rule(self):
        sel = self.kr_tree.selection()
        if not sel:
            mb.showinfo("提示", "请选择一条规则")
            return
        rid = int(sel[0])
        rule = next((r for r in self.current_rules if r["id"] == rid), None)
        if not rule:
            return
        dlg = KeywordRuleEditDlg(self, rule)
        if dlg.result:
            db.update_keyword_rule(rid, **dlg.result)
            self._refresh()
            self._notify()

    def _del_rule(self):
        sel = self.kr_tree.selection()
        if not sel:
            return
        rid = int(sel[0])
        if mb.askyesno("确认", "删除此规则？"):
            db.delete_keyword_rule(rid)
            self._refresh()
            self._notify()

    def _move_sel(self, direction):
        sel = self.kr_tree.selection()
        if not sel:
            return
        rid = int(sel[0])
        rule = next((r for r in self.current_rules if r["id"] == rid), None)
        if not rule:
            return
        new_priority = rule.get("priority", 10) + direction * 5
        db.update_keyword_rule(rid, priority=new_priority)
        self._refresh()
        self._notify()

    def _save(self):
        self._notify()
        mb.showinfo("完成", "规则已保存，导入时将自动匹配。")
        self.destroy()

    def _notify(self):
        if self.on_updated:
            self.on_updated(db.list_keyword_rules())


class KeywordRuleEditDlg(tk.Toplevel):
    """新建/编辑单条关键词规则"""

    def __init__(self, parent, rule=None):
        super().__init__(parent)
        self.result = None
        self.rule = rule
        self.title("编辑规则" if rule else "新建规则")
        self.geometry("500x340")
        self.grab_set()
        self._build()

    def _build(self):
        f = tk.Frame(self, padx=20, pady=16)
        f.pack(fill=tk.BOTH, expand=True)

        for i, (lbl, key, w, v0) in enumerate([
            ("关键词（英文或中文）：", "keyword", 30, ""),
            ("操作类型：", "action_type", 15, "tag"),
            ("目标值（分类ID或标签名）：", "action_value", 30, ""),
            ("优先级（1-99）：", "priority", 8, "10"),
        ]):
            tk.Label(f, text=lbl, font=FONT_BODY, anchor="w").grid(row=i, column=0, sticky="nw", pady=8)
            if key == "action_type":
                self.act_var = tk.StringVar(value=self.rule.get("action_type", "tag") if self.rule else "tag")
                for j, (val, lbl2) in enumerate([("tag", "打标签"), ("category", "分类"), ("label", "标签")]):
                    tk.Radiobutton(f, text=lbl2, variable=self.act_var, value=val,
                                  font=FONT_BODY).grid(row=i, column=1, sticky="w", padx=(0, j*80))
            else:
                v = tk.StringVar(value=str(self.rule.get(key, v0) if self.rule else v0))
                e = tk.Entry(f, textvariable=v, font=FONT_BODY, width=w)
                e.grid(row=i, column=1, sticky="w", pady=8)

        # 颜色
        tk.Label(f, text="标签颜色：", font=FONT_BODY, anchor="w").grid(
            row=4, column=0, sticky="nw", pady=8)
        self.color_var = tk.StringVar(value=self.rule.get("label_color", "#3498DB") if self.rule else "#3498DB")
        color_fr = tk.Frame(f)
        color_fr.grid(row=4, column=1, sticky="w", pady=8)
        for col in ["#3498DB", "#27AE60", "#E67E22", "#E74C3C", "#9B59B6", "#1ABC9C", "#F39C12", "#95A5A6"]:
            b = tk.Button(color_fr, bg=col, width=3, bd=1, cursor="hand2",
                         command=lambda c=col: self.color_var.set(c))
            b.pack(side=tk.LEFT, padx=2)
        tk.Label(color_fr, textvariable=self.color_var, font=FONT_SMALL, fg=C_SUBTXT).pack(side=tk.LEFT, padx=6)

        bf = tk.Frame(f)
        bf.grid(row=5, column=0, columnspan=2, pady=(12, 0))
        tk.Button(bf, text="取消", command=self.destroy, font=FONT_BODY,
                 width=10).pack(side=tk.RIGHT, padx=6)
        tk.Button(bf, text="保存", command=self._save, font=FONT_BODY,
                 width=10, bg=C_ACCENT, fg="white", cursor="hand2").pack(side=tk.RIGHT)

    def _save(self):
        # 收集表单数据
        keyword_val = None
        action_type_val = self.act_var.get()
        action_value_val = None
        priority_val = 10
        for w in self.children.values():
            if isinstance(w, tk.Frame):
                for child in w.winfo_children():
                    if isinstance(child, tk.Entry) and child.winfo_viewable():
                        v = child.get().strip()
                        if not keyword_val:
                            keyword_val = v
                        elif not action_value_val:
                            action_value_val = v
                        else:
                            try:
                                priority_val = int(v)
                            except ValueError:
                                pass
        if not keyword_val:
            mb.showwarning("提示", "请输入关键词", parent=self)
            return
        self.result = {
            "keyword": keyword_val,
            "action_type": action_type_val,
            "action_value": action_value_val or "",
            "priority": priority_val,
            "label_color": self.color_var.get(),
        }
        self.destroy()


