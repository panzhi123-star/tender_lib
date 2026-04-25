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

class BatchDlg(tk.Toplevel):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.title("批量管理")
        self.geometry("700x520")
        self.grab_set()
        nb = ttk.Notebook(self)
        nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self._build_import(nb)
        self._build_merge(nb)
        self._build_history(nb)
        self._build_keyword(nb)

    def log_insert(self, msg):
        self.app.log_insert(msg)

    def _build_import(self, nb):
        tab = tk.Frame(nb, padx=16, pady=14)
        nb.add(tab, text="  批量导入  ")
        tk.Label(tab, text="批量导入标书文件（.docx/.doc/.pdf/.txt）",
                 font=FONT_BODY, anchor="w", wraplength=600).pack(anchor="w")
        self.batch_var = tk.StringVar()
        tk.Entry(tab, textvariable=self.batch_var, font=FONT_BODY,
                 width=72, state="readonly").pack(fill=tk.X, pady=6)
        tk.Button(tab, text="选择文件", font=FONT_BODY,
                 command=self._sel_batch).pack(anchor="w", pady=2)
        self.split_var = tk.BooleanVar(value=False)
        tk.Checkbutton(tab, text="按段落拆分成多个条目（精细插入）",
                       variable=self.split_var, font=FONT_BODY).pack(anchor="w", pady=6)
        tk.Button(tab, text="开始导入", font=FONT_BODY,
                 bg=C_ACCENT2, fg="white", cursor="hand2",
                 padx=16, pady=6, command=self._do_import).pack(pady=16, anchor="center")
        tk.Label(tab, text="提示：支持批量选择多个文件（Ctrl+多选）",
                 font=FONT_SMALL, fg=C_SUBTXT, anchor="w").pack(anchor="w")

    def _build_merge(self, nb):
        tab = tk.Frame(nb, padx=16, pady=14)
        nb.add(tab, text="  合并条目  ")
        tk.Label(tab, text="将多个条目合并成一个（双击预览）：",
                 font=FONT_BODY, anchor="w").pack(anchor="w")
        tf = tk.Frame(tab)
        tf.pack(fill=tk.BOTH, expand=True, pady=8)
        self.merge_tree = ttk.Treeview(tf, columns=["标题", "分类"], show="headings", height=12)
        self.merge_tree.heading("标题", text="标题")
        self.merge_tree.column("标题", width=400)
        self.merge_tree.heading("分类", text="分类")
        self.merge_tree.column("分类", width=150)
        mv = ttk.Scrollbar(tf, orient=tk.VERTICAL, command=self.merge_tree.yview)
        self.merge_tree.configure(yscrollcommand=mv.set)
        self.merge_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        mv.pack(side=tk.RIGHT, fill=tk.Y)
        cats = {c["id"]: c["name"] for c in db.list_categories()}
        for e in db.list_entries(limit=200):
            self.merge_tree.insert("", tk.END, values=(e["title"][:50], cats.get(e.get("category_id"), "")))
        bf = tk.Frame(tab)
        bf.pack()
        tk.Button(bf, text="合并选中", font=FONT_BODY, bg=C_ACCENT2, fg="white",
                 cursor="hand2", padx=14, command=self._do_merge).pack(side=tk.LEFT, padx=6)
        tk.Button(bf, text="刷新", font=FONT_BODY, padx=10, command=self._build_merge).pack(side=tk.LEFT)

    def _build_history(self, nb):
        tab = tk.Frame(nb, padx=16, pady=14)
        nb.add(tab, text="  操作历史  ")
        tk.Label(tab, text="最近操作记录：",
                 font=FONT_BODY, anchor="w").pack(anchor="w")
        tf = tk.Frame(tab)
        tf.pack(fill=tk.BOTH, expand=True, pady=8)
        self.hist_tree = ttk.Treeview(tf, columns=["时间", "操作", "详情"], show="headings", height=14)
        for col, w in [("时间", 140), ("操作", 80), ("详情", 400)]:
            self.hist_tree.heading(col, text=col)
            self.hist_tree.column(col, width=w, anchor="w")
        hv = ttk.Scrollbar(tf, orient=tk.VERTICAL, command=self.hist_tree.yview)
        self.hist_tree.configure(yscrollcommand=hv.set)
        self.hist_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        hv.pack(side=tk.RIGHT, fill=tk.Y)
        for h in db.get_history(limit=100):
            self.hist_tree.insert("", tk.END, values=(
                h["happened_at"][:16], h["action"], h.get("detail", "")))

    def _build_keyword(self, nb):
        tab = tk.Frame(nb, padx=16, pady=14)
        nb.add(tab, text="  关键词提取  ")
        tk.Label(tab, text="输入关键词，批量为条目打标签：",
                 font=FONT_BODY, anchor="w").pack(anchor="w")
        self.kw_var = tk.StringVar()
        tk.Entry(tab, textvariable=self.kw_var, font=FONT_BODY,
                 width=50).pack(anchor="w", pady=6)
        tk.Label(tab, text="匹配标签的关键词（逗号分隔）：",
                 font=FONT_SMALL, fg=C_SUBTXT, anchor="w").pack(anchor="w")
        self.kw_tags_var = tk.StringVar()
        tk.Entry(tab, textvariable=self.kw_tags_var, font=FONT_BODY,
                 width=50).pack(anchor="w", pady=6)
        tk.Button(tab, text="批量打标签", font=FONT_BODY,
                 bg=C_ACCENT2, fg="white", cursor="hand2",
                 padx=14, command=self._do_tag).pack(anchor="w", pady=10)
        tk.Label(tab, text="说明：自动为包含关键词的条目添加标签，便于分类管理。",
                 font=FONT_SMALL, fg=C_SUBTXT, anchor="w").pack(anchor="w")

    def _sel_batch(self):
        files = fd.askopenfilenames(title="批量导入",
            filetypes=[("标书文件", "*.docx *.doc *.pdf *.txt"), ("全部", "*.*")])
        if files:
            self.batch_var.set("; ".join(files))

    def _do_import(self):
        paths = [p.strip() for p in self.batch_var.get().split(";") if p.strip()]
        if not paths:
            mb.showwarning("提示", "请先选择文件", parent=self)
            return
        cnt = 0
        img_total = 0
        self.log_insert(f"[批量管理] 开始导入 {len(paths)} 个文件...")
        for path in paths:
            try:
                paras, imgs = extractors.extract_from_file(path)
                fname = os.path.basename(path)
                text = "\n".join(p["text"] for p in paras)
                if paras:
                    eid = db.create_entry(title=fname, raw_text=text,
                                   content_type="composite" if imgs else "text",
                                   source_file=str(path))
                    for img in imgs:
                        db.add_attachment(eid, img["name"], "."+img["ext"], img["data"])
                    img_total += len(imgs)
                    cnt += 1
                    self.log_insert(f"  {fname}: {len(paras)}段落，{len(imgs)}图片")
            except Exception as ex:
                self.log_insert(f"  [错误] {os.path.basename(path)}: {ex}")
        self.batch_var.set("")
        self.app._refresh_list()
        self.app._update_stats()
        self.log_insert(f"[完成] 共 {cnt} 个条目，{img_total} 张图片")
        mb.showinfo("导入完成", f"共创建 {cnt} 个条目，{img_total} 张图片", parent=self)

    def _do_merge(self):
        sel = self.merge_tree.selection()
        if len(sel) < 2:
            mb.showwarning("提示", "请至少选择2个条目", parent=self)
            return
        entries = [db.get_entry(int(self.merge_tree.item(s)["values"][0])) for s in sel]
        combined = "\n\n---\n\n".join(
            f"【{e['title']}】\n{e.get('raw_text', '')}" for e in entries)
        title = sd.askstring("合并", "输入合并后条目名称：",
                             initialvalue=entries[0]["title"], parent=self)
        if title:
            db.create_entry(title=title, raw_text=combined,
                          content_type="composite",
                          category_id=entries[0].get("category_id"))
            self.app._refresh_list()
            self.app._update_stats()
            mb.showinfo("完成", "新条目已创建", parent=self)

    def _do_tag(self):
        kw = self.kw_var.get().strip()
        tags = self.kw_tags_var.get().strip()
        if not kw or not tags:
            mb.showwarning("提示", "请输入关键词和标签", parent=self)
            return
        rows = db.search_entries(kw, limit=500)
        n = 0
        for r in rows:
            cur_tags = r.get("tags", "") or ""
            tag_list = [t.strip() for t in cur_tags.split(",") if t.strip()]
            for t in tags.split(","):
                t = t.strip()
                if t and t not in tag_list:
                    tag_list.append(t)
            new_tags = ",".join(tag_list)
            db.update_entry(r["id"], tags=new_tags)
            n += 1
        self.app._refresh_list()
        mb.showinfo("完成", f"已为 {n} 个条目添加标签「{tags}」", parent=self)


# ══════════════════════════════════════════════════════════════════
#  入口
# ══════════════════════════════════════════════════════════════════

def main():
    root = tk.Tk()
    root.update_idletasks()
    App(root)
    root.mainloop()

if __name__ == "__main__":
    main()


