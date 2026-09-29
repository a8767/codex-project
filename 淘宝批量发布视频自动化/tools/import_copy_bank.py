"""Import the original copy bank (带货种草向 / 氛围感向) into the reference database.

Source of truth is the JSON file next to the workbook, so the copy can be edited
without touching Python code. The import is idempotent: rows are upserted by
copy_id, and rows removed from the JSON are deleted from the table.

Run from the project root: python tools/import_copy_bank.py
"""

from __future__ import annotations

import json
import sqlite3
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs" / "video_reference_db"
DB_PATH = OUT / "视频发布参考数据库.sqlite3"
JSON_PATH = OUT / "原创文案库.json"
TITLE_LIMIT = 30

CREATE_TABLE = """
CREATE TABLE IF NOT EXISTS original_copy_bank (
    copy_id TEXT PRIMARY KEY,
    direction TEXT NOT NULL,
    video_theme TEXT NOT NULL,
    title TEXT NOT NULL,
    body TEXT NOT NULL,
    angle TEXT NOT NULL,
    keywords TEXT,
    source TEXT NOT NULL DEFAULT '原创草稿',
    created_at TEXT NOT NULL
)
"""


def load_items(path: Path = JSON_PATH) -> tuple[dict, list[dict]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    return payload.get("meta", {}), payload.get("items", [])


def validate(items: list[dict]) -> None:
    seen: set[str] = set()
    for item in items:
        for field in ("id", "direction", "theme", "title", "body", "angle"):
            if not str(item.get(field, "")).strip():
                raise ValueError(f"文案 {item.get('id', '?')} 缺少字段：{field}")
        if item["id"] in seen:
            raise ValueError(f"文案编号重复：{item['id']}")
        seen.add(item["id"])
        length = len(item["title"])
        if length > TITLE_LIMIT:
            raise ValueError(f"标题超过 {TITLE_LIMIT} 字（{length}）：{item['title']}")


def import_copy_bank(db: sqlite3.Connection, path: Path = JSON_PATH) -> int:
    meta, items = load_items(path)
    validate(items)
    source = meta.get("source", "原创草稿")
    created_at = meta.get("created_at", "")
    db.execute(CREATE_TABLE)
    db.execute("CREATE INDEX IF NOT EXISTS idx_copy_bank_direction ON original_copy_bank(direction, video_theme)")
    db.executemany(
        """INSERT INTO original_copy_bank
           (copy_id, direction, video_theme, title, body, angle, keywords, source, created_at)
           VALUES(?,?,?,?,?,?,?,?,?)
           ON CONFLICT(copy_id) DO UPDATE SET
             direction=excluded.direction,
             video_theme=excluded.video_theme,
             title=excluded.title,
             body=excluded.body,
             angle=excluded.angle,
             keywords=excluded.keywords,
             source=excluded.source,
             created_at=excluded.created_at""",
        (
            (item["id"], item["direction"], item["theme"], item["title"], item["body"],
             item["angle"], item.get("keywords", ""), source, created_at)
            for item in items
        ),
    )
    keep = [item["id"] for item in items]
    placeholders = ",".join("?" for _ in keep) or "''"
    db.execute(f"DELETE FROM original_copy_bank WHERE copy_id NOT IN ({placeholders})", keep)
    db.commit()
    return db.execute("SELECT count(*) FROM original_copy_bank").fetchone()[0]


def main() -> None:
    if not JSON_PATH.exists():
        raise FileNotFoundError(f"缺少文案源文件：{JSON_PATH}")
    db = sqlite3.connect(DB_PATH)
    try:
        total = import_copy_bank(db)
        by_direction = db.execute(
            "SELECT direction, count(*) FROM original_copy_bank GROUP BY direction ORDER BY direction"
        ).fetchall()
    finally:
        db.close()
    print(json.dumps({"original_copy_bank": total,
                      "by_direction": {k: v for k, v in by_direction}}, ensure_ascii=False))


if __name__ == "__main__":
    main()
