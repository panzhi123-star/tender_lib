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

class ImportPreviewDlg(tk.Toplevel):
    """解析预览 + 关键词高亮 + 逐段编辑 + 入库确认"""

    def __init__(self, parent, paragraphs, images, filename, keyword_rules):
        super().__init__(parent)
        self.title(f"导入预览校对 - {filename}")
        self.geometry("1200x780")
        self.minsize(800, 500)
        self.grab_set()
        self._edit_timer = None

        # Center on screen
        self.update_idletasks()
        try:
            pw = self.winfo_reqwidth()
            ph = self.winfo_reqheight()
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            x = max(80, (sw - pw) // 2)
            y = max(40, (sh - ph) // 2)
            self.geometry(f"{pw}x{ph}+{x}+{y}")
        except Exception:
            pass  # Fallback: let tkinter position

        self.paragraphs = paragraphs          # [{text, style, level, is_heading, checked, ...}]
        self.images = images                  # [{name, ext, data}]
        self.filename = filename
        self.keyword_rules = keyword_rules    # 关键词规则列表
        self.result_entries = []             # 最终要入库的条目

        # 预处理：给每段打勾 + 匹配关键词
        self._preprocess()

        self._build_ui()

    def _preprocess(self):
        """预处理：检查段落，匹配关键词"""
        cats = {c["id"]: c["name"] for c in db.list_categories()}
        for p in self.paragraphs:
            p["checked"] = True   # 默认全选
            p["edit_text"] = p["text"]
            # 匹配关键词
            p["matched_rules"] = []
            text_lower = p["text"].lower()
            for rule in self.keyword_rules:
                if not rule.get("enabled", 1):
                    continue
                kw = rule["keyword"].lower()
                if kw and kw in text_lower:
                    p["matched_rules"].append(rule)
            # 自动分类建议
            p["suggested_cat"] = None
            for rule in p["matched_rules"]:
                if rule["action_type"] == "category":
                    try:
                        p["suggested_cat"] = cats.get(int(rule["action_value"]), "")
                    except (ValueError, TypeError):
                        pass
            # 收集标签
            tags = set()
            for rule in p["matched_rules"]:
                if rule["action_type"] in ("tag", "label"):
                    tags.add(rule["action_value"])
            p["matched_tags"] = list(tags)

    def _build_ui(self):
        """Modern card-based import preview UI"""
        self.configure(bg="#F8FAFC")
        
        # ========== Header Bar ==========
        header = tk.Frame(self, bg="#1E293B", height=56)
        header.pack(side=tk.TOP, fill=tk.X)
        header.pack_propagate(False)
        
        # File icon + name
        icon_lbl = tk.Label(header, text="", font=("Segoe UI", 20),
                           bg="#1E293B", fg="#60A5FA")
        icon_lbl.pack(side=tk.LEFT, padx=(20, 8))
        tk.Label(header, text=self.filename, font=("Microsoft YaHei", 12, "bold"),
                bg="#1E293B", fg="white").pack(side=tk.LEFT, pady=12)
        
        # Stats pills
        stats_fr = tk.Frame(header, bg="#1E293B")
        stats_fr.pack(side=tk.RIGHT, padx=20)
        for val, lbl, clr in [(str(len(self.paragraphs)), "段落", "#60A5FA"),
                              (str(len(self.images)), "图片", "#34D399")]:
            pill = tk.Frame(stats_fr, bg="#334155", padx=12, pady=4)
            pill.pack(side=tk.LEFT, padx=(0, 10))
            tk.Label(pill, text=val, font=("Microsoft YaHei", 14, "bold"),
                    bg="#334155", fg=clr).pack(side=tk.LEFT)
            tk.Label(pill, text=lbl, font=("Microsoft YaHei", 9),
                    bg="#334155", fg="#94A3B8").pack(side=tk.LEFT, padx=(4, 0))
        
        # ========== Toolbar ==========
        toolbar = tk.Frame(self, bg="white", height=52)
        toolbar.pack(side=tk.TOP, fill=tk.X)
        toolbar.pack_propagate(False)
        
        # Left: Selection buttons
        left_grp = tk.Frame(toolbar, bg="white")
        left_grp.pack(side=tk.LEFT, padx=16, pady=10)
        
        for txt, cmd, bg, fg in [
            ("✓ 全选", self._select_all, "#DBEAFE", "#1E40AF"),
            ("✕ 取消", self._deselect_all, "#FEF3C7", "#92400E"),
        ]:
            btn = tk.Button(left_grp, text=txt, command=cmd,
                          font=("Microsoft YaHei", 9), bg=bg, fg=fg,
                          bd=0, cursor="hand2", padx=14, pady=5,
                          activebackground=bg, activeforeground=fg)
            btn.pack(side=tk.LEFT, padx=(0, 8))
        
        # Divider
        tk.Frame(toolbar, bg="#E2E8F0", width=1).pack(side=tk.LEFT, fill=tk.Y, padx=4, pady=12)
        
        # Filter dropdown
        filt_fr = tk.Frame(toolbar, bg="white")
        filt_fr.pack(side=tk.LEFT, padx=12, pady=10)
        tk.Label(filt_fr, text="筛选:", font=("Microsoft YaHei", 9),
                bg="white", fg="#64748B").pack(side=tk.LEFT)
        self._filter_var = tk.StringVar(value="全部")
        filt_cb = ttk.Combobox(filt_fr, textvariable=self._filter_var,
                              values=["全部", "标题", "正文", "表格", "图片说明"],
                              state="readonly", width=10, font=("Microsoft YaHei", 9))
        filt_cb.pack(side=tk.LEFT, padx=(8, 0))
        filt_cb.bind("<<ComboboxSelected>>", lambda e: self._on_filter_changed())
        
        # Parse mode
        mode_fr = tk.Frame(toolbar, bg="white")
        mode_fr.pack(side=tk.LEFT, padx=12, pady=10)
        tk.Label(mode_fr, text="解析:", font=("Microsoft YaHei", 9),
                bg="white", fg="#64748B").pack(side=tk.LEFT)
        self.parse_mode = tk.StringVar(value="auto")
        mode_cb = ttk.Combobox(mode_fr, textvariable=self.parse_mode,
                              values=["auto", "page", "heading"],
                              state="readonly", width=10, font=("Microsoft YaHei", 9))
        mode_cb.pack(side=tk.LEFT, padx=(8, 0))
        mode_cb.bind("<<ComboboxSelected>>", lambda e: self._apply_parse_mode())
        
        # Right: Action buttons
        right_grp = tk.Frame(toolbar, bg="white")
        right_grp.pack(side=tk.RIGHT, padx=16, pady=10)
        
        self.sel_count_lbl = tk.Label(right_grp, text="已选: 0", font=("Microsoft YaHei", 10, "bold"),
                                     bg="white", fg="#3B82F6")
        self.sel_count_lbl.pack(side=tk.RIGHT, padx=(16, 0))
        
        for txt, cmd, bg, fg in [
            ("⚡ 自动分类", self._auto_categorize, "#D1FAE5", "#065F46"),
            ("📁 批量分类", self._open_cat_assign, "#F3E8FF", "#6B21A8"),
        ]:
            btn = tk.Button(right_grp, text=txt, command=cmd,
                          font=("Microsoft YaHei", 9), bg=bg, fg=fg,
                          bd=0, cursor="hand2", padx=14, pady=5,
                          activebackground=bg, activeforeground=fg)
            btn.pack(side=tk.RIGHT, padx=(0, 10))
        
        # ========== Main Content ==========
        content = tk.Frame(self, bg="#F8FAFC")
        content.pack(fill=tk.BOTH, expand=True, padx=16, pady=(12, 0))
        
        pw = tk.PanedWindow(content, orient=tk.HORIZONTAL, bg="#E2E8F0",
                           sashwidth=4, sashrelief=tk.FLAT)
        pw.pack(fill=tk.BOTH, expand=True)
        
        # ---- Left: Paragraph list ----
        left_card = tk.Frame(pw, bg="white", highlightbackground="#E2E8F0",
                            highlightthickness=1)
        pw.add(left_card, width=580, minsize=300)
        
        # List header
        list_hdr = tk.Frame(left_card, bg="#F1F5F9", height=40)
        list_hdr.pack(fill=tk.X)
        list_hdr.pack_propagate(False)
        tk.Label(list_hdr, text="☰ 解析结果", font=("Microsoft YaHei", 11, "bold"),
                bg="#F1F5F9", fg="#1E293B").pack(side=tk.LEFT, padx=16, pady=8)
        tk.Label(list_hdr, text="勾选入库 · 双击编辑", font=("Microsoft YaHei", 9),
                bg="#F1F5F9", fg="#94A3B8").pack(side=tk.RIGHT, padx=16, pady=8)
        
        # Treeview with modern styling
        style = ttk.Style()
        style.configure("Modern.Treeview",
                       background="white",
                       foreground="#334155",
                       fieldbackground="white",
                       rowheight=36,
                       font=("Microsoft YaHei", 10),
                       borderwidth=0)
        style.configure("Modern.Treeview.Heading",
                       background="#F1F5F9",
                       foreground="#475569",
                       font=("Microsoft YaHei", 9, "bold"),
                       borderwidth=0)
        style.map("Modern.Treeview",
                 background=[("selected", "#DBEAFE")],
                 foreground=[("selected", "#1E40AF")])
        
        tree_frame = tk.Frame(left_card, bg="white")
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)
        
        cols = ("#0", "类型", "预览", "关键词")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="tree headings",
                                style="Modern.Treeview", height=16)
        self.tree.heading("#0", text="✓")
        self.tree.column("#0", width=40, anchor="center")
        self.tree.heading("类型", text="类型")
        self.tree.column("类型", width=70, anchor="center")
        self.tree.heading("预览", text="内容预览")
        self.tree.column("预览", width=320)
        self.tree.heading("关键词", text="匹配关键词")
        self.tree.column("关键词", width=120)
        
        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree.bind("<<TreeviewSelect>>", self._on_para_select)
        self.tree.bind("<Button-1>", self._on_checkbox_click)
        self.tree.bind("<Double-Button-1>", self._edit_current_para)
        
        # ---- Right: Editor panel ----
        right_card = tk.Frame(pw, bg="white", highlightbackground="#E2E8F0",
                             highlightthickness=1)
        pw.add(right_card, width=420, minsize=280)
        
        # Editor header
        edit_hdr = tk.Frame(right_card, bg="#F1F5F9", height=40)
        edit_hdr.pack(fill=tk.X)
        edit_hdr.pack_propagate(False)
        tk.Label(edit_hdr, text="✏️ 编辑段落", font=("Microsoft YaHei", 11, "bold"),
                bg="#F1F5F9", fg="#1E293B").pack(side=tk.LEFT, padx=16, pady=8)
        
        # Editor content
        edit_body = tk.Frame(right_card, bg="white")
        edit_body.pack(fill=tk.BOTH, expand=True, padx=16, pady=12)
        
        tk.Label(edit_body, text="段落内容:", font=("Microsoft YaHei", 9),
                bg="white", fg="#64748B").pack(anchor="w", pady=(0, 6))
        
        self.edit_text = tk.Text(edit_body, wrap=tk.WORD, font=("Microsoft YaHei", 10),
                                bg="#FAFBFC", fg="#1E293B", insertbackground="#3B82F6",
                                height=12, bd=1, relief=tk.SOLID,
                                highlightbackground="#E2E8F0", highlightthickness=1)
        self.edit_text.pack(fill=tk.BOTH, expand=True)
        self.edit_text.bind("<KeyRelease>", self._on_text_edit)
        
        # Keywords section
        kw_frame = tk.Frame(edit_body, bg="white")
        kw_frame.pack(fill=tk.X, pady=(16, 0))
        tk.Label(kw_frame, text="匹配关键词:", font=("Microsoft YaHei", 9),
                bg="white", fg="#64748B").pack(anchor="w")
        self.kw_lbl = tk.Label(kw_frame, text="无", font=("Microsoft YaHei", 9),
                              bg="white", fg="#94A3B8", wraplength=360)
        self.kw_lbl.pack(anchor="w", pady=(4, 0))
        
        # Suggested category
        cat_frame = tk.Frame(edit_body, bg="white")
        cat_frame.pack(fill=tk.X, pady=(12, 0))
        tk.Label(cat_frame, text="建议分类:", font=("Microsoft YaHei", 9),
                bg="white", fg="#64748B").pack(anchor="w")
        self.suggest_lbl = tk.Label(cat_frame, text="未分类", font=("Microsoft YaHei", 10, "bold"),
                                   bg="#FEF3C7", fg="#92400E", padx=10, pady=4)
        self.suggest_lbl.pack(anchor="w", pady=(4, 0))
        
        # Images section
        img_hdr = tk.Frame(right_card, bg="#F1F5F9", height=36)
        img_hdr.pack(fill=tk.X, pady=(12, 0))
        img_hdr.pack_propagate(False)
        tk.Label(img_hdr, text="图片附件", font=("Microsoft YaHei", 10, "bold"),
                bg="#F1F5F9", fg="#1E293B").pack(side=tk.LEFT, padx=16, pady=6)
        
        self.img_canvas = tk.Canvas(right_card, bg="#F8FAFC", highlightthickness=0, height=100)
        self.img_frame = tk.Frame(self.img_canvas, bg="#F8FAFC")
        self.img_canvas.create_window((0, 0), window=self.img_frame, anchor="nw")
        self.img_frame.bind("<Configure>", lambda e: self.img_canvas.configure(
            scrollregion=self.img_canvas.bbox("all")))
        
        img_hsb = ttk.Scrollbar(right_card, orient=tk.HORIZONTAL, command=self.img_canvas.xview)
        self.img_canvas.configure(xscrollcommand=img_hsb.set)
        
        self.img_canvas.pack(fill=tk.X, padx=1, pady=1)
        img_hsb.pack(fill=tk.X)
        self.img_canvas.pack_forget()
        img_hsb.pack_forget()
        
        # 图片计数标签
        self.img_count_lbl = tk.Label(right_card, text="0张", font=("Microsoft YaHei", 8),
                                     bg="#F1F5F9", fg="#64748B")
        self.img_count_lbl.pack(fill=tk.X)
        
        # ========== Bottom Action Bar ==========
        bottom = tk.Frame(self, bg="white", height=64)
        bottom.pack(side=tk.BOTTOM, fill=tk.X)
        bottom.pack_propagate(False)
        
        # Left: Rules button
        tk.Button(bottom, text="⚙️ 关键词规则", command=self._open_kw_rules,
                 font=("Microsoft YaHei", 10), bg="#F1F5F9", fg="#475569",
                 bd=0, cursor="hand2", padx=20, pady=8).pack(side=tk.LEFT, padx=20, pady=12)
        
        # Right: Import buttons
        btn_fr = tk.Frame(bottom, bg="white")
        btn_fr.pack(side=tk.RIGHT, padx=20)
        
        tk.Button(btn_fr, text="取消", command=self.destroy,
                 font=("Microsoft YaHei", 10), bg="#F1F5F9", fg="#64748B",
                 bd=0, cursor="hand2", padx=24, pady=10).pack(side=tk.LEFT, padx=(0, 12))
        
        tk.Button(btn_fr, text="导入选中项", command=self._import_selected,
                 font=("Microsoft YaHei", 10, "bold"), bg="#DBEAFE", fg="#1E40AF",
                 bd=0, cursor="hand2", padx=24, pady=10).pack(side=tk.LEFT, padx=(0, 12))
        
        tk.Button(btn_fr, text="全部导入", command=self._import_all,
                 font=("Microsoft YaHei", 10, "bold"), bg="#3B82F6", fg="white",
                 bd=0, cursor="hand2", padx=24, pady=10).pack(side=tk.LEFT)
        
        self._refresh_list()

    def _refresh_list(self):
        """刷新段落列表"""
        try:
            self.tree.tag_configure("heading",
                                            background="#E8F5E9",
                                            foreground="#1B5E20",
                                            font=("Microsoft YaHei", 9, "bold"))
            self.tree.tag_configure("body",
                                            background="white",
                                            foreground="#37474F")
            self.tree.tag_configure("unchecked",
                                            background="#FFF3E0",
                                            foreground="#BF360C")
        except tk.TclError:
            pass

        for row in self.tree.get_children():
            self.tree.delete(row)
        for i, p in enumerate(self.paragraphs):
            checked = p.get("checked", True)
            icon = "\u2611" if checked else "\u2610"
            is_heading = p.get("is_heading")
            level = p.get("level", 0)
            indent = "    " * level if is_heading else ""
            summary = indent + p["edit_text"][:60].replace("\n", " ")
            if len(p["edit_text"]) > 60:
                summary += "..."
            tags_str = ", ".join(p.get("matched_tags", [])[:4])
            if is_heading:
                tags_str = "\u3010\u6807\u9898\u3011" + tags_str
            edited_mark = "\u270E" if p.get("edit_text", "") != p.get("text", "") else ""
            if checked:
                tags = ("heading",) if is_heading else ("body",)
            else:
                tags = ("unchecked",)
            self.tree.insert("", tk.END, iid=str(i),
                                     text=edited_mark + icon,
                                     values=("", summary[:50], tags_str[:20]),
                                     tags=tags)
    def _on_checkbox_click(self, event):
        """Handle click on checkbox column (#0). Toggle checked state."""
        try:
            col_id = self.tree.identify_column(event.x)
            row_id = self.tree.identify_row(event.y)
            if col_id != "#0" or not row_id:
                return
            idx = int(row_id)
            p = self.paragraphs[idx]
            p["checked"] = not p.get("checked", True)
            # 更新单行显示
            self._update_single_row(idx)
            # 更新底部计数
            cnt = sum(1 for p2 in self.paragraphs if p2.get("checked", True))
        except Exception:
            pass
    def _on_para_select(self, event=None):
        try:
            sel = self.tree.selection()
            if not sel:
                return
            idx = int(sel[0])
            p = self.paragraphs[idx]
            self.edit_text.config(state="normal")
            self.edit_text.delete("1.0", tk.END)
            self.edit_text.insert("1.0", p.get("edit_text", ""))
            self._update_kw_info(idx)
            self._show_images(idx)
        except Exception:
            pass
    def _on_text_edit(self, event=None):
        """编辑框内容变化时更新元信息（带防抖，避免频繁刷新）"""
        # 防抖：取消上次定时器，500ms后执行
        if self._edit_timer:
            self.after_cancel(self._edit_timer)
        self._edit_timer = self.after(500, self._do_text_update)

    def _do_text_update(self):
        """实际执行编辑更新（防抖回调）"""
        self._edit_timer = None
        text = self.edit_text.get("1.0", tk.END).rstrip("\n")
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        self.paragraphs[idx]["edit_text"] = text
        # 更新字数
        # 重新匹配关键词
        p = self.paragraphs[idx]
        p["matched_rules"] = []
        text_lower = text.lower()
        for rule in self.keyword_rules:
            if not rule.get("enabled", 1):
                continue
            kw = rule["keyword"].lower()
            if kw and kw in text_lower:
                p["matched_rules"].append(rule)
        tags = set()
        for rule in p["matched_rules"]:
            if rule["action_type"] in ("tag", "label"):
                tags.add(rule["action_value"])
        p["matched_tags"] = list(tags)
        # 更新标签显示
        self.kw_lbl.config(text=", ".join(p["matched_tags"]) if p["matched_tags"] else "无")
        # 只更新当前行，不重建整个列表
        self._update_single_row(idx)

    def _update_single_row(self, idx):
        """更新 Treeview 中单行的显示"""
        iid = str(idx)
        if not self.tree.exists(iid):
            return
        p = self.paragraphs[idx]
        checked = p.get("checked", True)
        icon = "\u2611" if checked else "\u2610"
        is_heading = p.get("is_heading")
        level = p.get("level", 0)
        indent = "    " * level if is_heading else ""
        summary = indent + p["edit_text"][:60].replace("\n", " ")
        if len(p["edit_text"]) > 60:
            summary += "..."
        tags_str = ", ".join(p.get("matched_tags", [])[:4])
        if is_heading:
            tags_str = "\u3010\u6807\u9898\u3011" + tags_str
        edited_mark = "\u270E" if p.get("edit_text", "") != p.get("text", "") else ""
        chars = len(p.get("edit_text", ""))
        if checked:
            tags = ("heading",) if is_heading else ("body",)
        else:
            tags = ("unchecked",)
        self.tree.item(iid, text=edited_mark + icon,
                               values=(summary, tags_str[:30], str(chars)),
                               tags=tags)
    def _show_images(self, idx):
        """显示与当前段落关联的图片缩略图（PDF按页/DOCX全文档）"""
        # 销毁旧缩略图（释放 PhotoImage 引用）
        for w in self.img_frame.winfo_children():
            w.destroy()

        # 缩略图缓存：{img_idx: (PhotoImage, pil_image)}
        if not hasattr(self, "_img_cache"):
            self._img_cache = {}
            self._img_order = []   # LRU 顺序

        p = self.paragraphs[idx]
        page_num = p.get("page", 0)

        # ─ 图片来源：优先使用 per-entry images 字段 ─
        entry_imgs = p.get("images", [])
        if entry_imgs:
            # 新格式：每条条目有自己的图片列表
            pdf_imgs = [img for img in entry_imgs if img.get("source") == "pdf"]
            docx_imgs = [img for img in entry_imgs if img.get("source") == "docx"]
        else:
            # 旧格式（fallback）：PDF按页码过滤，DOCX不显示（无法精确关联）
            pdf_imgs = [img for img in self.images
                        if img.get("source") == "pdf" and img.get("page") == page_num]
            docx_imgs = []  # DOCX图片必须通过 entry["images"] 关联

        total = len(pdf_imgs) + len(docx_imgs)
        self.img_count_lbl.config(text=f"{total}张")
        
        # Show canvas only when there's content
        if total > 0:
            self.img_canvas.pack(fill=tk.X, padx=1, pady=1)
        else:
            self.img_canvas.pack_forget()

        def _thumb(img_data, key):
            """获取或创建缓存缩略图"""
            if key in self._img_cache:
                # 移到LRU末尾
                self._img_order.remove(key)
                self._img_order.append(key)
                return self._img_cache[key][0]
            from PIL import Image as PILImage
            from PIL import ImageTk as PILImageTk
            pil_img = PILImage.open(io.BytesIO(img_data))
            pil_img.thumbnail((95, 90))
            photo = PILImageTk.PhotoImage(pil_img)
            # LRU：最多缓存30张
            if len(self._img_cache) >= 30:
                old_key = self._img_order.pop(0)
                old_photo, _ = self._img_cache.pop(old_key)
                del old_photo
            self._img_cache[key] = (photo, pil_img)
            self._img_order.append(key)
            return photo

        def _make_thumb(parent, img, img_key):
            """创建单个缩略图卡片（美化版）"""
            fr = tk.Frame(parent, bg="white", bd=0,
                         highlightbackground="#E0E0E0", highlightthickness=1,
                         cursor="hand2")
            photo = _thumb(img["data"], img_key)
            lbl_img = tk.Label(fr, image=photo, bg="white")
            lbl_img.pack(padx=4, pady=(4, 2))
            lbl_img.bind("<Button-1>",
                        lambda e, d=img["data"], n=img["name"]: self._zoom_img(d, n))
            tk.Label(fr, text=img["name"][:14], font=("Consolas", 7),
                     bg="white", fg="#9E9E9E").pack(pady=(0, 4))
            fr.pack(side=tk.LEFT, padx=4, pady=4)

        if not pdf_imgs and not docx_imgs:
            self.img_canvas.pack_forget()
            tk.Label(self.img_frame, text="✖ 无关联图片",
                     font=("Microsoft YaHei", 8), bg="#F8FAFC", fg="#BDBDBD").pack(anchor="w", padx=12, pady=12)
            return
        
        self.img_canvas.pack(fill=tk.X, padx=1, pady=1)

        # PDF 页内图片
        for img in pdf_imgs:
            img_key = id(img)  # 稳定key
            _make_thumb(self.img_frame, img, img_key)

        # DOCX 全局图片（用特殊key）
        if docx_imgs:
            # 分隔标签
            sep = tk.Frame(self.img_frame, bg=C_BORDER, width=1)
            sep.pack(side=tk.LEFT, fill=tk.Y, padx=2)
            hint = tk.Label(self.img_frame, text=f"📎 文档图片 ({len(docx_imgs)}张)",
                            font=FONT_SMALL, bg=C_PANEL, fg=C_SUBTXT)
            hint.pack(side=tk.LEFT, padx=6)
            for img in docx_imgs:
                _make_thumb(self.img_frame, img, ("docx", img["name"]))

    def _zoom_img(self, img_data, name):
        """点击放大图片"""
        from PIL import Image as PILImage
        from PIL import ImageTk as PILImageTk
        pil = PILImage.open(io.BytesIO(img_data))
        win = tk.Toplevel(self.root)
        win.title(name)
        win.configure(bg="#222")
        # 限制最大尺寸
        max_w, max_h = 900, 700
        w, h = pil.size
        scale = min(max_w / w, max_h / h, 1.0)
        new_w, new_h = int(w * scale), int(h * scale)
        pil_display = pil.resize((new_w, new_h), PILImage.LANCZOS)
        photo = PILImageTk.PhotoImage(pil_display)
        # 保持引用
        win.photo = photo
        canvas = tk.Canvas(win, width=new_w, height=new_h, bg="#222", highlightthickness=0)
        canvas.create_image(0, 0, anchor=tk.NW, image=photo)
        canvas.pack()
        # 居中
        win.update_idletasks()
        sx = win.winfo_screenwidth() // 2 - new_w // 2
        sy = win.winfo_screenheight() // 2 - new_h // 2
        win.geometry(f"{new_w}x{new_h}+{sx}+{sy}")
        win.grab_set()

    def _update_kw_info(self, idx):
        """更新关键词匹配信息"""
        p = self.paragraphs[idx]
        rules = p.get("matched_rules", [])
        if rules:
            kw_parts = []
            for r in rules[:6]:
                label = r.get("action_value") or r.get("keyword")
                kw_parts.append(f"\u25CF {label}")
            self.kw_lbl.config(text="\u5339\u914D: " + "  ".join(kw_parts),
                               fg="#F57F17")
        else:
            self.kw_lbl.config(text="\u5339\u914D\u5173\u952E\u8BCD: \u65E0",
                               fg="#BDBDBD")
        # 分类建议
        cat_name = p.get("suggested_cat") or "\u2014"
        self.cat_lbl.config(text=f"\u5206\u7C7B: {cat_name}")
        # 标签
        tag_text = ", ".join(p.get("matched_tags", [])) if p.get("matched_tags") else "\u2014"
        self.kw_lbl.config(text=f"\u6807\u7B7E: {tag_text}")
        # 字数
    def _edit_current_para(self):
        """双击编辑当前段落"""
        self.edit_text.focus_set()

    def _select_all(self):
        for p in self.paragraphs:
            p["checked"] = True
        self._refresh_list()
    def _deselect_all(self):
        for p in self.paragraphs:
            p["checked"] = False
        self._refresh_list()
    def _filter_type(self, dtype):
        if dtype == "Normal":
            # 仅保留正文（取消标题勾选，保留正文）
            for p in self.paragraphs:
                p["checked"] = not p.get("is_heading", False)
        elif dtype == "heading":
            # 仅保留标题
            for p in self.paragraphs:
                p["checked"] = p.get("is_heading", False)
        self._refresh_list()
        checked_cnt = sum(1 for p2 in self.paragraphs if p2.get("checked", True))

    def _auto_categorize(self):
        cats = {c["id"]: c["name"] for c in db.list_categories()}
        total = len(self.paragraphs)
        categorized = 0
        for p in self.paragraphs:
            for rule in p.get("matched_rules", []):
                if rule["action_type"] == "category":
                    try:
                        cid = int(rule["action_value"])
                        p["suggested_cat"] = cats.get(cid, "")
                        categorized += 1
                    except (ValueError, TypeError):
                        pass
        self._refresh_list()
        mb.showinfo("自动分类",
                   f"已根据关键词进行自动分类，共{categorized}段落已标记分类。")


    def _on_filter_changed(self):
        filt = self._filter_var.get()
        checked = 0
        for iid in self.tree.get_children():
            p = self.paragraphs[int(iid)]
            if filt == "正文" and p.get("is_heading"):
                self.tree.detach(iid); continue
            if filt == "标题" and not p.get("is_heading"):
                self.tree.detach(iid); continue
            if filt == "表格" and p.get("is_table", False):
                self.tree.detach(iid); continue
            self.tree.move(iid, "", "end")
            checked += 1 if p.get("checked", True) else 0
        self.sel_count_lbl.config(text=f"已选: {checked}/{len(self.paragraphs)}")

    def _on_parse_mode_change(self):
        """解析模式改变时的回调"""
        pass

    def _apply_parse_mode(self):
        """应用当前选择的解析模式"""
        mode = self.parse_mode.get()
        if mode == "auto":
            self._auto_categorize()
        elif mode == "page":
            self._categorize_by_page()
        elif mode == "heading":
            self._categorize_by_heading()

    def _categorize_by_page(self):
        """按页码分类：将同一页的段落归入同一分类"""
        pages = {}
        for p in self.paragraphs:
            page = p.get("page", 0)
            if page not in pages:
                pages[page] = []
            pages[page].append(p)

        cats = {c["name"]: c["id"] for c in db.list_categories()}
        for page_num, paras in sorted(pages.items()):
            cat_name = f"第{page_num}页"
            if cat_name not in cats:
                try:
                    cat_id = db.create_category(cat_name)
                    cats[cat_name] = cat_id
                except Exception as e:
                    print(f"Create category failed: {e}")
                    continue
            else:
                cat_id = cats[cat_name]
            for p in paras:
                p["suggested_cat"] = cat_name

        self._refresh_list()
        mb.showinfo("按页分类", f"已将文档按 {len(pages)} 个页面分类")

    def _categorize_by_heading(self):
        """按标题等级分类：根据标题层级创建分类结构"""
        headings = []
        for p in self.paragraphs:
            if p.get("is_heading") and p.get("level", 0) > 0:
                headings.append({
                    "text": p["text"][:30],
                    "level": p["level"],
                    "full_text": p["text"]
                })

        if not headings:
            mb.showwarning("按标题分类", "未检测到标题，请确认文档格式")
            return

        cats = {c["name"]: c["id"] for c in db.list_categories()}
        parent_stack = [None]

        for h in headings[:20]:
            level = h["level"]
            cat_name = h["text"]

            while len(parent_stack) > level:
                parent_stack.pop()

            parent_id = parent_stack[-1] if parent_stack else None

            if cat_name not in cats:
                try:
                    cat_id = db.create_category(cat_name, parent_id=parent_id)
                    cats[cat_name] = cat_id
                except Exception as e:
                    print(f"Create category failed: {e}")
                    continue
            else:
                cat_id = cats[cat_name]

            parent_stack.append(cat_id)

        current_cat = None
        for p in self.paragraphs:
            if p.get("is_heading") and p.get("level", 0) > 0:
                cat_name = p["text"][:30]
                current_cat = cat_name if cat_name in cats else None
            if current_cat:
                p["suggested_cat"] = current_cat

        self._refresh_list()
        mb.showinfo("按标题分类", f"已根据 {len(headings)} 个标题创建分类结构")

    def _import_all(self):
        self._do_import(check_selected=False)

    def _import_selected(self):
        self._do_import(check_selected=True)

    def _open_cat_assign(self):
        """Open category assignment dialog for selected paragraphs."""
        cats = db.list_categories()
        if not cats:
            mb.showinfo("提示", "请先创建分类", parent=self)
            return
        sel = self.tree.selection()
        if not sel:
            mb.showinfo("提示", "请先在列表中选择要分配的段落", parent=self)
            return

        win = tk.Toplevel(self)
        win.title("分配分类")
        win.geometry("400x300")
        win.transient(self)
        win.grab_set()

        tk.Label(win, text="选择分类：",
                 font=FONT_BODY).pack(anchor="w", padx=16, pady=(16, 4))

        cat_var = tk.StringVar()
        for cat in cats:
            tk.Radiobutton(win, text=cat["name"], variable=cat_var,
                          value=str(cat["id"]),
                          font=FONT_BODY, anchor="w").pack(anchor="w", padx=32, pady=2)

        tk.Label(win, text="或输入新分类名称：",
                 font=FONT_SMALL, fg=C_SUBTXT).pack(anchor="w", padx=16, pady=(12, 2))
        new_cat_entry = tk.Entry(win, font=FONT_BODY, width=30)
        new_cat_entry.pack(padx=16, pady=4)

        def do_assign():
            new_name = new_cat_entry.get().strip()
            if new_name:
                cid = db.create_category(new_name)
                if cid:
                    self.parent._refresh_cats()
            else:
                cid = int(cat_var.get()) if cat_var.get() else None
            if cid is None and not new_name:
                mb.showwarning("提示", "请选择分类或输入新名称", parent=win)
                return

            # Apply to selected paragraphs
            sel_idx = int(sel[0])
            if 0 <= sel_idx < len(self.paragraphs):
                self.paragraphs[sel_idx]["suggested_cat"] = new_name if new_name else cats[[c["id"] for c in cats].index(cid)]["name"]
                self._refresh_list()
                self._update_kw_info(sel_idx)
            win.destroy()

        bf = tk.Frame(win)
        bf.pack(pady=16)
        tk.Button(bf, text="确认分配", command=do_assign,
                 font=FONT_BODY, bg=C_ACCENT, fg="white",
                 cursor="hand2", padx=16, pady=4).pack(side=tk.LEFT, padx=8)
        tk.Button(bf, text="取消", command=win.destroy,
                 font=FONT_BODY, padx=16, pady=4).pack(side=tk.LEFT)


    def _do_import(self, check_selected=False):
        """执行入库"""
        selected = [p for p in self.paragraphs if not check_selected or p.get("checked", True)]
        if not selected:
            mb.showwarning("提示", "请至少选择一段落导入", parent=self)
            return

        # 预处理：将短/空段落的图片重新分配到最近的非空段落
        # 避免跳过短段落时丢失其关联图片
        non_empty_indices = []
        for i, p in enumerate(selected):
            text = p.get("edit_text", "").strip()
            if text and len(text) >= 5:
                non_empty_indices.append(i)

        for i, p in enumerate(selected):
            text = p.get("edit_text", "").strip()
            if (not text or len(text) < 5) and p.get("images"):
                # 找最近的非空段落（先向下，再向上）
                target_idx = None
                for offset in range(1, len(selected)):
                    if i + offset < len(selected) and (i + offset) in non_empty_indices:
                        target_idx = i + offset
                        break
                    if i - offset >= 0 and (i - offset) in non_empty_indices:
                        target_idx = i - offset
                        break
                if target_idx is not None:
                    selected[target_idx]["images"].extend(p["images"])
                    # 记录调试信息
                    try:
                        import os as _os
                        _log_dir = _os.path.join(_os.environ.get('APPDATA', _os.expanduser('~')), 'tender_lib')
                        _os.makedirs(_log_dir, exist_ok=True)
                        with open(_os.path.join(_log_dir, 'debug_img_assoc.log'), 'a', encoding='utf-8') as _f:
                            _f.write(f"[导入重分配] 段落{i}('{text[:20]}')的{len(p['images'])}张图片→段落{target_idx}\n")
                    except Exception:
                        pass
                    p["images"] = []  # 清空原段落的图片，避免重复

        cats = {c["id"]: c["name"] for c in db.list_categories()}
        entry_titles = set()
        cnt = 0
        img_total = 0

        for p in selected:
            text = p.get("edit_text", "").strip()
            if not text or len(text) < 5:
                continue

            # 标题取前50字
            title = text[:50]
            if len(text) > 50:
                title += "..."
            if title in entry_titles and not check_selected:
                title = f"{title} (2)"
            entry_titles.add(title)

            # 关键词标签
            tags = list(p.get("matched_tags", []))
            if p.get("is_heading"):
                tags.insert(0, "标题")
            if p.get("style") == "Table":
                tags.insert(0, "表格")

            # 分类
            cat_id = None
            cat_name = p.get("suggested_cat", "")
            if cat_name:
                cat_id_map = {v: k for k, v in cats.items()}
                cat_id = cat_id_map.get(cat_name)

            # 入库
            eid = db.create_entry(
                title=title,
                raw_text=text,
                content_type="text",
                category_id=cat_id,
                tags=",".join(tags),
                source_file=self.filename
            )
            cnt += 1

            # 图片附件关联到段落
            for img in p.get("images", []):
                db.add_attachment(eid, img["name"], "."+img["ext"], img["data"])
                img_total += 1

        self.result_entries = selected
        mb.showinfo("导入完成",
                   f"共导入 {cnt} 段落，{img_total} 张图片。",
                   parent=self)
        # 通知主窗口刷新
        try:
            if hasattr(self, "parent") and self.parent:
                if hasattr(self.parent, "_refresh_list"):
                    self.parent._refresh_list()
                if hasattr(self.parent, "_update_stats"):
                    self.parent._update_stats()
        except Exception:
            pass
        self.destroy()

    def _open_kw_rules(self):
        KeywordRuleDlg(self, self.keyword_rules, self._on_rules_updated)

    def _on_rules_updated(self, rules):
        self.keyword_rules = rules
        for p in self.paragraphs:
            p["matched_rules"] = []
            text_lower = p.get("edit_text", "").lower()
            for rule in self.keyword_rules:
                if not rule.get("enabled", 1):
                    continue
                kw = rule["keyword"].lower()
                if kw and kw in text_lower:
                    p["matched_rules"].append(rule)
            tags = set()
            for rule in p["matched_rules"]:
                if rule["action_type"] in ("tag", "label"):
                    tags.add(rule["action_value"])
            p["matched_tags"] = list(tags)
        self._refresh_list()
        if self.tree.selection():
            self._update_kw_info(int(self.tree.selection()[0]))


# ══════════════════════════════════════════════════════════════════
# 关键词规则管理对话框
# ══════════════════════════════════════════════════════════════════


