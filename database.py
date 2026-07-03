"""project-cost-sheet — 数据库层"""
import sqlite3, logging
from decimal import Decimal
from pathlib import Path
from contextlib import contextmanager

logger = logging.getLogger(__name__)

sqlite3.register_adapter(Decimal, str)
sqlite3.register_converter("decimal", lambda s: Decimal(s.decode() if isinstance(s, bytes) else s))

SCHEMA = """
CREATE TABLE IF NOT EXISTS projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    project_types TEXT NOT NULL DEFAULT 'print',
    currency TEXT DEFAULT 'CNY',
    tax_rate DECIMAL DEFAULT 0,
    sort_order INTEGER DEFAULT 0,
    notes TEXT DEFAULT '',
    created_at TEXT DEFAULT (datetime('now','localtime')),
    updated_at TEXT DEFAULT (datetime('now','localtime'))
);

CREATE TABLE IF NOT EXISTS budget_categories (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
    name TEXT NOT NULL,
    sort_order INTEGER DEFAULT 0
);

CREATE TABLE IF NOT EXISTS line_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    category_id INTEGER NOT NULL REFERENCES budget_categories(id) ON DELETE CASCADE,
    description TEXT NOT NULL DEFAULT '',
    unit_price DECIMAL DEFAULT 0,
    quantity DECIMAL DEFAULT 1,
    unit TEXT DEFAULT '',
    total DECIMAL DEFAULT 0,
    notes TEXT DEFAULT '',
    sort_order INTEGER DEFAULT 0
);

CREATE INDEX IF NOT EXISTS idx_cat_project ON budget_categories(project_id);
CREATE INDEX IF NOT EXISTS idx_item_category ON line_items(category_id);
"""


class Database:
    def __init__(self, path="budget.db"):
        self.path = Path(path)

    def connect(self):
        c = sqlite3.connect(str(self.path),
            detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
        c.execute("PRAGMA journal_mode=WAL")
        c.execute("PRAGMA foreign_keys=ON")
        c.row_factory = sqlite3.Row
        return c

    @contextmanager
    def tx(self):
        c = self.connect()
        try:
            yield c
            c.commit()
        except Exception:
            c.rollback()
            raise
        finally:
            c.close()

    def init(self):
        with self.tx() as c:
            for s in SCHEMA.strip().split(";"):
                if s.strip():
                    c.execute(s.strip())
            # migration: add sort_order if missing (v1.1)
            cols = [r[1] for r in c.execute("PRAGMA table_info(projects)").fetchall()]
            if "sort_order" not in cols:
                c.execute("ALTER TABLE projects ADD COLUMN sort_order INTEGER DEFAULT 0")

    def fetch(self, sql, params=()):
        with self.connect() as c:
            return c.execute(sql, params).fetchall()

    def fetch_one(self, sql, params=()):
        with self.connect() as c:
            return c.execute(sql, params).fetchone()

    def exec(self, sql, params=()):
        with self.tx() as c:
            return c.execute(sql, params)

    def exec_insert(self, sql, params=()):
        with self.tx() as c:
            cur = c.execute(sql, params)
            return cur.lastrowid
