# -*- coding: utf-8 -*-
"""
标书资料库 - SQLite 数据模型 v2
支持：条目 / 附件 / 版本历史 / 模板 / 共享模式
"""
import sqlite3, os, datetime, json

# ── 路径（与 main_gui.py 保持一致）───────────────────────────────
_db_dir = os.environ.get("TENDER_DB_DIR",
    os.path.join(os.environ.get("APPDATA",
    os.path.join(os.path.expanduser("~"), "AppData", "Roaming")),
    "tender_lib"))
os.makedirs(_db_dir, exist_ok=True)
DB_PATH = os.path.join(_db_dir, "tender_library.db")

def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn

# ── 数据库初始化 ────────────────────────────────────────────────
def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.executescript("""
    CREATE TABLE IF NOT EXISTS categories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL UNIQUE,
        parent_id INTEGER REFERENCES categories(id),
        sort_order INTEGER DEFAULT 0
    );
    CREATE TABLE IF NOT EXISTS entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        uuid TEXT NOT NULL UNIQUE DEFAULT (lower(hex(randomblob(16)))),
        category_id INTEGER REFERENCES categories(id) ON DELETE SET NULL,
        title TEXT NOT NULL,
        content_type TEXT NOT NULL DEFAULT 'text',
        raw_text TEXT,
        source_file TEXT,
        tags TEXT DEFAULT '',
        created_at TEXT NOT NULL DEFAULT (datetime('now')),
        updated_at TEXT NOT NULL DEFAULT (datetime('now')),
        use_count INTEGER DEFAULT 0
    );
    CREATE TABLE IF NOT EXISTS entry_attachments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        entry_id INTEGER NOT NULL REFERENCES entries(id) ON DELETE CASCADE,
        file_name TEXT NOT NULL,
        file_ext TEXT NOT NULL,
        file_data BLOB NOT NULL,
        mime_type TEXT DEFAULT '',
        sort_order INTEGER DEFAULT 0,
        created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS entry_versions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        entry_id INTEGER NOT NULL REFERENCES entries(id) ON DELETE CASCADE,
        version_no INTEGER NOT NULL DEFAULT 1,
        title TEXT,
        raw_text TEXT,
        content_type TEXT,
        tags TEXT,
        comment TEXT DEFAULT '',
        created_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS templates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        category TEXT DEFAULT '',
        description TEXT DEFAULT '',
        content TEXT DEFAULT '',
        structure TEXT DEFAULT '[]',
        use_count INTEGER DEFAULT 0,
        created_at TEXT NOT NULL DEFAULT (datetime('now')),
        updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS settings (
        key TEXT PRIMARY KEY,
        value TEXT
    );
    CREATE TABLE IF NOT EXISTS history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        entry_id INTEGER REFERENCES entries(id) ON DELETE SET NULL,
        action TEXT NOT NULL,
        detail TEXT,
        happened_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE TABLE IF NOT EXISTS keyword_rules (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        keyword       TEXT NOT NULL,
        action_type   TEXT NOT NULL DEFAULT 'tag',
        action_value  TEXT DEFAULT '',
        priority      INTEGER DEFAULT 10,
        label_color   TEXT DEFAULT '#3498DB',
        enabled       INTEGER DEFAULT 1,
        created_at    TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_entries_title ON entries(title);
    CREATE INDEX IF NOT EXISTS idx_entries_cat ON entries(category_id);
    CREATE INDEX IF NOT EXISTS idx_attach_entry ON entry_attachments(entry_id);
    """)
    conn.commit()

    # 默认分类
    defaults = ["公司资质","工程业绩","施工组织设计","技术方案","安全管理","人员配置","机械设备","政府采购"]
    for name in defaults:
        cur.execute("INSERT OR IGNORE INTO categories(name) VALUES(?)", (name,))

    # 默认模板
    d_tpls = [
        ("技术标模板","技术标","通用技术标章节结构","","[]"),
        ("商务标模板","商务标","商务标章节结构，含公司资质、业绩等","",""),
        ("投标函模板","通用","投标函标准格式","",""),
    ]
    for n,c,d,ct,s in d_tpls:
        cur.execute("INSERT OR IGNORE INTO templates(name) VALUES(?)",(n,))

    # 默认关键词规则
    kw_defaults = [
        ("施工组织设计", "category", "4", 20, "#27AE60"),
        ("技术方案", "category", "5", 20, "#E67E22"),
        ("安全管理", "category", "6", 20, "#E74C3C"),
        ("项目经理", "tag", "项目经理", 15, "#9B59B6"),
        ("资质证书", "tag", "资质", 15, "#3498DB"),
        ("业绩", "tag", "业绩", 15, "#1ABC9C"),
        ("投标报价", "tag", "报价", 15, "#F39C12"),
        ("工期", "tag", "工期", 10, "#95A5A6"),
        ("生产安全", "tag", "安全", 10, "#E74C3C"),
    ]
    for kw, at, av, pri, col in kw_defaults:
        cur.execute(
            "INSERT OR IGNORE INTO keyword_rules(keyword,action_type,action_value,priority,label_color)"
            " VALUES(?,?,?,?,?)", (kw, at, av, pri, col))

    conn.close()

# ── 分类 ───────────────────────────────────────────────────────
def list_categories(parent_id=None):
    """List categories, optionally filtered by parent_id."""
    conn = get_conn()
    if parent_id is None:
        rows = conn.execute(
            "SELECT * FROM categories WHERE parent_id IS NULL ORDER BY sort_order, name").fetchall()
    else:
        rows = conn.execute(
            "SELECT * FROM categories WHERE parent_id=? ORDER BY sort_order, name",
            (parent_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def list_all_categories():
    """Get all categories with parent info for tree display."""
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM categories ORDER BY sort_order, name").fetchall()
    conn.close()
    return [dict(r) for r in rows]

def create_category(name, parent_id=None):
    """Create a category. Optionally specify parent_id for subcategory."""
    conn = get_conn()
    try:
        cur = conn.execute(
            "INSERT INTO categories(name, parent_id) VALUES(?,?)",
            (name, parent_id))
        conn.commit()
        rid = cur.lastrowid
    except sqlite3.IntegrityError:
        rid = None
    conn.close()
    return rid

def delete_category(cid):
    conn = get_conn()
    conn.execute("DELETE FROM categories WHERE id=?",(cid,))
    conn.commit()
    conn.close()

def rename_category(cid, new_name):
    """Rename a category. Returns True on success, False if name exists."""
    conn = get_conn()
    try:
        conn.execute("UPDATE categories SET name=? WHERE id=?", (new_name, cid))
        conn.commit()
        success = True
    except sqlite3.IntegrityError:
        success = False
    conn.close()
    return success


# ── 条目 ───────────────────────────────────────────────────────
def create_entry(title, content_type="text", raw_text="",
                 category_id=None, tags="", source_file=""):
    conn = get_conn()
    cur = conn.execute(
        "INSERT INTO entries(title,content_type,raw_text,category_id,tags,source_file)"
        " VALUES(?,?,?,?,?,?)",
        (title,content_type,raw_text,category_id,tags,source_file))
    eid = cur.lastrowid
    conn.execute("INSERT INTO history(entry_id,action,detail) VALUES(?,?,?)",
                 (eid,"created",title))
    conn.commit()
    conn.close()
    return eid

def update_entry(eid, **kw):
    conn = get_conn()
    # 保存旧版本
    old = conn.execute("SELECT * FROM entries WHERE id=?",(eid,)).fetchone()
    if old:
        max_v = conn.execute(
            "SELECT MAX(version_no) FROM entry_versions WHERE entry_id=?",
            (eid,)).fetchone()[0] or 0
        conn.execute(
            "INSERT INTO entry_versions(entry_id,version_no,title,raw_text,content_type,tags)"
            " VALUES(?,?,?,?,?,?)",
            (eid, max_v+1, old["title"], old["raw_text"],
             old["content_type"], old["tags"]))
    # 更新
    # Handle category_id separately so None (=uncategorised) is preserved
    cat_id = kw.pop("category_id", ...)
    fields = {k:v for k,v in kw.items() if v is not None}
    if fields:
        set_sql = ",".join(f"{k}=?" for k in fields)
        conn.execute(
            f"UPDATE entries SET {set_sql},updated_at=datetime('now') WHERE id=?",
            list(fields.values())+[eid])
    if cat_id is not ...:
        conn.execute("UPDATE entries SET category_id=?,updated_at=datetime('now') WHERE id=?",
                      (cat_id, eid))
    conn.execute("INSERT INTO history(entry_id,action,detail) VALUES(?,?,?)",
                 (eid,"updated",""))
    conn.commit()
    conn.close()

def delete_entry(eid):
    conn = get_conn()
    conn.execute("DELETE FROM entries WHERE id=?",(eid,))
    conn.commit()
    conn.close()

def get_entry(eid):
    conn = get_conn()
    row = conn.execute("SELECT * FROM entries WHERE id=?",(eid,)).fetchone()
    conn.close()
    return dict(row) if row else None

def search_entries(keyword, cat_id=None, limit=50):
    conn = get_conn()
    kw = f"%{keyword}%"
    sql = ("SELECT * FROM entries WHERE "
           "(title LIKE ? OR raw_text LIKE ? OR tags LIKE ?) ")
    params = [kw, kw, kw]
    if cat_id is not None:
        sql += " AND category_id=? "
        params.append(cat_id)
    sql += "ORDER BY use_count DESC,updated_at DESC LIMIT ?"
    params.append(limit)
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def list_entries(cat_id=None, limit=200):
    conn = get_conn()
    if cat_id is not None:
        rows = conn.execute(
            "SELECT * FROM entries WHERE category_id=? "
            "ORDER BY updated_at DESC LIMIT ?",(cat_id,limit)).fetchall()
    else:
        rows = conn.execute(
            "SELECT * FROM entries ORDER BY updated_at DESC LIMIT ?",
            (limit,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def increment_use(eids):
    if not eids: return
    conn = get_conn()
    for eid in eids:
        conn.execute("UPDATE entries SET use_count=use_count+1 WHERE id=?",(eid,))
    conn.commit()
    conn.close()

# ── 附件 ───────────────────────────────────────────────────────
def add_attachment(eid, file_name, file_ext, file_data, mime_type=""):
    conn = get_conn()
    cur = conn.execute(
        "INSERT INTO entry_attachments(entry_id,file_name,file_ext,file_data,mime_type)"
        " VALUES(?,?,?,?,?)",(eid,file_name,file_ext,file_data,mime_type))
    conn.commit()
    conn.close()
    return cur.lastrowid

def list_attachments(eid):
    conn = get_conn()
    rows = conn.execute(
        "SELECT id,file_name,file_ext,mime_type,created_at "
        "FROM entry_attachments WHERE entry_id=? ORDER BY sort_order",(eid,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def get_attachment_data(aid):
    conn = get_conn()
    row = conn.execute(
        "SELECT file_data,file_ext,file_name FROM entry_attachments WHERE id=?",(aid,)).fetchone()
    conn.close()
    return row["file_data"] if row else None

def delete_attachment(aid):
    conn = get_conn()
    conn.execute("DELETE FROM entry_attachments WHERE id=?",(aid,))
    conn.commit()
    conn.close()

# ── 版本历史 ───────────────────────────────────────────────────

def swap_attachment_order(aid1, aid2):
    """Swap the sort_order of two attachments."""
    conn = get_conn()
    a1 = conn.execute("SELECT sort_order FROM entry_attachments WHERE id=?", (aid1,)).fetchone()
    a2 = conn.execute("SELECT sort_order FROM entry_attachments WHERE id=?", (aid2,)).fetchone()
    if a1 and a2:
        conn.execute("UPDATE entry_attachments SET sort_order=? WHERE id=?", (a2[0], aid1))
        conn.execute("UPDATE entry_attachments SET sort_order=? WHERE id=?", (a1[0], aid2))
        conn.commit()
    conn.close()


def get_versions(eid):
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM entry_versions WHERE entry_id=? ORDER BY version_no DESC",
        (eid,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]

def restore_version(vid):
    conn = get_conn()
    ver = conn.execute(
        "SELECT * FROM entry_versions WHERE id=?",(vid,)).fetchone()
    if ver:
        conn.execute(
            "UPDATE entries SET title=?,raw_text=?,content_type=?,tags=?,"
            "updated_at=datetime('now') WHERE id=?",
            (ver["title"],ver["raw_text"],ver["content_type"],
             ver["tags"],ver["entry_id"]))
        conn.commit()
        result = ver["entry_id"]
    else:
        result = None
    conn.close()
    return result

# ── 模板 ───────────────────────────────────────────────────────
def list_templates():
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM templates ORDER BY use_count DESC, name").fetchall()
    conn.close()
    return [dict(r) for r in rows]

def create_template(name, category="", description="", content="", structure="[]"):
    conn = get_conn()
    cur = conn.execute(
        "INSERT INTO templates(name,category,description,content,structure)"
        " VALUES(?,?,?,?,?)",(name,category,description,content,structure))
    conn.commit()
    conn.close()
    return cur.lastrowid

def update_template(tid, **kw):
    conn = get_conn()
    fields = {k:v for k,v in kw.items() if v is not None}
    if fields:
        set_sql = ",".join(f"{k}=?" for k in fields)
        conn.execute(
            f"UPDATE templates SET {set_sql},updated_at=datetime('now') WHERE id=?",
            list(fields.values())+[tid])
        conn.commit()
    conn.close()

def delete_template(tid):
    conn = get_conn()
    conn.execute("DELETE FROM templates WHERE id=?",(tid,))
    conn.commit()
    conn.close()

def get_template(tid):
    conn = get_conn()
    row = conn.execute("SELECT * FROM templates WHERE id=?",(tid,)).fetchone()
    conn.close()
    return dict(row) if row else None

def increment_template_use(tid):
    conn = get_conn()
    conn.execute("UPDATE templates SET use_count=use_count+1 WHERE id=?",(tid,))
    conn.commit()
    conn.close()

# ── 设置 / 共享模式 ─────────────────────────────────────────────
def get_setting(key, default=""):
    conn = get_conn()
    row = conn.execute("SELECT value FROM settings WHERE key=?",(key,)).fetchone()
    conn.close()
    return row["value"] if row else default

def set_setting(key, value):
    conn = get_conn()
    conn.execute(
        "INSERT OR REPLACE INTO settings(key,value) VALUES(?,?)",(key,value))
    conn.commit()
    conn.close()

def find_entry_by_source(source_file):
    """Find existing entry by source file path"""
    conn = get_conn()
    row = conn.execute(
        "SELECT id FROM entries WHERE source_file=? LIMIT 1",
        (str(source_file),)).fetchone()
    conn.close()
    return row["id"] if row else None

def get_stats():
    conn = get_conn()
    s = {
        "total_entries": conn.execute("SELECT COUNT(*) FROM entries").fetchone()[0],
        "total_cats": conn.execute("SELECT COUNT(*) FROM categories").fetchone()[0],
        "total_attachs": conn.execute("SELECT COUNT(*) FROM entry_attachments").fetchone()[0],
        "total_templates": conn.execute("SELECT COUNT(*) FROM templates").fetchone()[0],
        "total_versions": conn.execute("SELECT COUNT(*) FROM entry_versions").fetchone()[0],
    }
    conn.close()
    return s

def get_history(eid=None, limit=50):
    conn = get_conn()
    if eid:
        rows = conn.execute(
            "SELECT * FROM history WHERE entry_id=? ORDER BY happened_at DESC LIMIT ?",
            (eid,limit)).fetchall()
    else:
        rows = conn.execute(
            "SELECT * FROM history ORDER BY happened_at DESC LIMIT ?",
            (limit,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def list_keyword_rules():
    conn = get_conn()
    rows = conn.execute("SELECT * FROM keyword_rules ORDER BY priority DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]

def create_keyword_rule(keyword, action_type, action_value="", priority=10, label_color="#3498DB"):
    conn = get_conn()
    cur = conn.execute(
        "INSERT INTO keyword_rules(keyword,action_type,action_value,priority,label_color)"
        " VALUES(?,?,?,?,?)",
        (keyword.strip(), action_type, action_value, priority, label_color))
    conn.commit()
    conn.close()
    return cur.lastrowid

def update_keyword_rule(rid, **kw):
    conn = get_conn()
    fields = {k: v for k, v in kw.items() if v is not None}
    if fields:
        set_sql = ",".join(f"{k}=?" for k in fields)
        conn.execute(f"UPDATE keyword_rules SET {set_sql} WHERE id=?", list(fields.values()) + [rid])
        conn.commit()
    conn.close()

def delete_keyword_rule(rid):
    conn = get_conn()
    conn.execute("DELETE FROM keyword_rules WHERE id=?", (rid,))
    conn.commit()
    conn.close()

def match_keywords(text, rules):
    matched_cats = []
    matched_tags = []
    text_lower = text.lower()
    for rule in rules:
        kw = rule["keyword"].strip().lower()
        if not kw:
            continue
        all_match = True
        for part in kw.split():
            if part and part not in text_lower:
                all_match = False
                break
        if all_match:
            if rule["action_type"] == "category":
                matched_cats.append(int(rule["action_value"]))
            elif rule["action_type"] in ("tag", "label"):
                if rule["action_value"]:
                    matched_tags.append(rule["action_value"])
    return matched_cats, matched_tags

