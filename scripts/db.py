import sqlite3, pathlib
from datetime import datetime

from pandas.io import sql
from sqlalchemy.sql.functions import now

BASE = pathlib.Path(__file__).resolve().parents[1]
DB_PATH = BASE / "storage" / "lab.db"
DB_PATH.parent.mkdir(parents=True, exist_ok=True)


def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        sql = (BASE / "sql" / "schema.sql").read_text(encoding="utf-8")
        conn.executescript(sql)

def upsert_samples(df, source_path:str):
    now = datetime.now().isoformat(timespec="seconds")
    with sqlite3.connect(DB_PATH) as conn:
        for _, row in df.iterrows():
            conn.execute("""
                INSERT INTO samples (sample_no, site_name, collected_at, kind, item, status, uniq_key, raw_path, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(uniq_key) DO UPDATE SET
                site_name=excluded.site_name,
                collected_at=excluded.collected_at,
                kind=excluded.kind,
                item=excluded.item,
                status=excluded.status,
                raw_path=excluded.raw_path,
                created_at=excluded.created_at
            """, (row.get("sample_no"), row.get("site_name"), row.get("collected_at"), row.get("kind"), row.get("item"), row.get("status"), row.get("uniq_key"), source_path, now))


def today_loads():
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute("""
            SELECT researcher, COUNT(*)
            FROM assignments
            WHERE date(assigned_at) = date('now', 'localtime')
            GROUP BY researcher
        """)
        return dict(cur.fetchall())

def save_assignments(rows):
    now = datetime.now().isoformat(timespec="seconds")
    with sqlite3.connect(DB_PATH) as conn:
        for r in rows:
            conn.execute("""
            INSERT INTO assignments (sample_no, item, researcher, assigned_at, method)
            VALUES (?, ?, ?, ?, ?)
            """, (r.get("sample_no"), r.get("item"), r.get("researcher"), now, r.get("method", "rule+rr")))
    