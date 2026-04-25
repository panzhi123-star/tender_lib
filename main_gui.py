# -*- coding: utf-8 -*-
"""标书资料库管理工具 v2.7.1 — 入口文件"""

import os
import sys

# ── PyInstaller 运行时路径修复 ──────────────────────────────
if getattr(sys, "frozen", False):
    base = getattr(sys, "_MEIPASS",
                   os.path.dirname(os.path.abspath(sys.executable)))
    if base not in sys.path:
        sys.path.insert(0, base)
    appdata = os.environ.get("APPDATA",
                              os.path.join(os.path.expanduser("~"),
                                           "AppData", "Roaming"))
    db_dir = os.path.join(appdata, "tender_lib")
    os.makedirs(db_dir, exist_ok=True)
    os.environ["TENDER_DB_DIR"] = db_dir

# 确保项目根目录在 sys.path 中
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
if _THIS_DIR not in sys.path:
    sys.path.insert(0, _THIS_DIR)

# ── 依赖检查 ───────────────────────────────────────────────
MISSING = []
for pkg, name in [("docx", "python-docx"), ("fitz", "PyMuPDF"), ("PIL", "Pillow")]:
    try:
        __import__(pkg)
    except ImportError:
        MISSING.append(name)
if MISSING:
    import tkinter as tk
    import tkinter.messagebox as mb
    tk.Tk().withdraw()
    mb.showerror("缺少依赖", "请先安装：pip install " + " ".join(MISSING))
    sys.exit(1)

import tkinter as tk
from gui.app import App


def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
