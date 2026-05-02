"""SQLite-Helper für V2-Projekt-State.
Ein Projekt = ein generiertes Exposé mit editierbaren Folien-Specs.
"""
from __future__ import annotations
import sqlite3
import json
import time
import uuid
from pathlib import Path
from typing import Any

DB_PATH = "/tmp/interpres_v2.db"


def _conn():
    c = sqlite3.connect(DB_PATH, timeout=10)
    c.row_factory = sqlite3.Row
    return c


def init():
    c = _conn()
    c.executescript("""
        CREATE TABLE IF NOT EXISTS projects (
            id          TEXT PRIMARY KEY,
            created     REAL NOT NULL,
            updated     REAL NOT NULL,
            name        TEXT,
            specs_json  TEXT NOT NULL,
            expose_json TEXT
        );
        CREATE TABLE IF NOT EXISTS uploads (
            project_id  TEXT NOT NULL,
            slot        TEXT NOT NULL,
            path        TEXT NOT NULL,
            uploaded    REAL NOT NULL,
            PRIMARY KEY (project_id, slot)
        );
    """)
    c.commit()
    c.close()


def create_project(specs: list[dict], expose: dict | None = None,
                   name: str = "") -> str:
    init()
    pid = str(uuid.uuid4())
    now = time.time()
    c = _conn()
    c.execute(
        "INSERT INTO projects (id, created, updated, name, specs_json, expose_json) "
        "VALUES (?, ?, ?, ?, ?, ?)",
        (pid, now, now, name, json.dumps(specs, ensure_ascii=False),
         json.dumps(expose or {}, ensure_ascii=False))
    )
    c.commit()
    c.close()
    return pid


def get_project(pid: str) -> dict | None:
    init()
    c = _conn()
    row = c.execute("SELECT * FROM projects WHERE id = ?", (pid,)).fetchone()
    c.close()
    if not row:
        return None
    return {
        "id":      row["id"],
        "created": row["created"],
        "updated": row["updated"],
        "name":    row["name"] or "",
        "specs":   json.loads(row["specs_json"]),
        "expose":  json.loads(row["expose_json"] or "{}"),
    }


def update_specs(pid: str, specs: list[dict]) -> bool:
    init()
    c = _conn()
    cur = c.execute(
        "UPDATE projects SET specs_json = ?, updated = ? WHERE id = ?",
        (json.dumps(specs, ensure_ascii=False), time.time(), pid),
    )
    c.commit()
    c.close()
    return cur.rowcount > 0


def update_slide(pid: str, idx: int, data: dict) -> bool:
    proj = get_project(pid)
    if not proj or idx < 0 or idx >= len(proj["specs"]):
        return False
    proj["specs"][idx]["data"] = {**proj["specs"][idx].get("data", {}), **data}
    return update_specs(pid, proj["specs"])


def list_uploads(pid: str) -> dict[str, str]:
    init()
    c = _conn()
    rows = c.execute(
        "SELECT slot, path FROM uploads WHERE project_id = ?", (pid,)
    ).fetchall()
    c.close()
    return {r["slot"]: r["path"] for r in rows}


def save_upload(pid: str, slot: str, path: str):
    init()
    c = _conn()
    c.execute(
        "INSERT OR REPLACE INTO uploads (project_id, slot, path, uploaded) "
        "VALUES (?, ?, ?, ?)",
        (pid, slot, path, time.time())
    )
    c.commit()
    c.close()


def remove_upload(pid: str, slot: str):
    init()
    c = _conn()
    c.execute("DELETE FROM uploads WHERE project_id = ? AND slot = ?", (pid, slot))
    c.commit()
    c.close()
