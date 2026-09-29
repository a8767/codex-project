"""Incrementally add captured 2026 Taobao Guanghe quality videos to the reference DB.

The import is idempotent by platform video ID: a video already present in `videos`
is left unchanged, including its copy and topics, as requested.
"""

from __future__ import annotations

import json
import re
import sqlite3
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / ".test_run" / "quality_2026_increment_20260930.json"
DATABASE = ROOT / "outputs" / "video_reference_db" / "视频发布参考数据库.sqlite3"
HASHTAG_RE = re.compile(r"[#＃]\s*([^\s#＃，,。！？!?；;、：:\[\]（）()]{1,60})")


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def main() -> None:
    source_data = json.loads(SOURCE.read_text(encoding="utf-8"))
    rows = source_data.get("records", []) if isinstance(source_data, dict) else source_data
    if not isinstance(rows, list):
        raise ValueError(f"采集数据应为列表：{SOURCE}")

    added = duplicate = topic_count = 0
    with sqlite3.connect(DATABASE) as db:
        db.execute("PRAGMA foreign_keys=ON")
        for source_row, row in enumerate(rows, 1):
            video_id = clean(row.get("id") or row.get("video_id"))
            title = clean(row.get("title"))
            published_at = clean(row.get("published_at"))
            if (not video_id.isdigit() or not title or not published_at.startswith("2026-")
                    or row.get("quality", True) is not True):
                raise ValueError(f"第 {source_row} 条缺少有效作品 ID、标题或 2026 年发布时间")
            video_key = "taobao:" + video_id
            if db.execute("SELECT 1 FROM videos WHERE video_key=?", (video_key,)).fetchone():
                duplicate += 1
                continue

            body = clean(row.get("copy") or row.get("body") or title)
            creator = clean(row.get("creator"))
            creator_id = clean(row.get("creator_id"))
            db.execute("""INSERT INTO videos
                (video_key, platform, video_id, title, body, creator, creator_id,
                 published_at, availability)
                VALUES (?, '淘宝光合', ?, ?, ?, ?, ?, ?, '可访问')""",
                (video_key, video_id, title, body, creator, creator_id, published_at))
            db.execute("""INSERT OR IGNORE INTO video_sources
                (video_key, source_type, source_file, source_row)
                VALUES (?, '淘宝优质视频', ?, ?)""",
                (video_key, SOURCE.name, source_row))

            platform_topics = [clean(item).lstrip("#＃ ") for item in row.get("topics", [])]
            content_tags = [clean(match.group(1)) for match in HASHTAG_RE.finditer(title + " " + body)]
            for topic in dict.fromkeys(platform_topics + content_tags):
                if topic:
                    extraction = "平台话题" if topic in platform_topics else "标题或文案#话题"
                    result = db.execute("""INSERT OR IGNORE INTO video_topics
                        (video_key, topic, extraction) VALUES (?, ?, ?)""",
                        (video_key, topic, extraction))
                    topic_count += result.rowcount
            added += 1

    print(json.dumps({"source_rows": len(rows), "added": added,
                      "skipped_existing_ids": duplicate, "new_topics": topic_count,
                      "database": str(DATABASE)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
