"""Import captured Guanghe ranking copy into the video reference database.

The capture is a resumable JSON file. Ranking positions remain separate from
videos so a work appearing in multiple charts is stored only once.
"""

from __future__ import annotations

import json
from pathlib import Path
import re
import sqlite3


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs" / "video_reference_db"
SOURCE = OUT / "光合热门榜_最近7天_文案话题.json"
DATABASE = OUT / "视频发布参考数据库.sqlite3"
TOPIC_RE = re.compile(r"[#＃]\s*([^\s#＃，,。！？!?；;、：:\[\]（）()]{1,60})")


def clean(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def import_capture(db: sqlite3.Connection, source: Path = SOURCE) -> dict[str, int]:
    data = json.loads(source.read_text(encoding="utf-8"))
    records = data["records"]
    positions: set[tuple[str, str, int]] = set()
    for row in records:
        position = (clean(row.get("date_range")), clean(row.get("chart")), int(row["rank"]))
        if not all(position) or position in positions:
            raise ValueError(f"缺少榜单位置或重复：{position}")
        positions.add(position)
        if not all(clean(row.get(field)) for field in ("video_id", "title", "copy")):
            raise ValueError(f"榜单 {position} 缺少作品 ID、标题或正文")

    db.execute("PRAGMA foreign_keys=ON")
    with db:
        db.execute("""CREATE TABLE IF NOT EXISTS guanghe_hot_occurrences (
            video_key TEXT NOT NULL REFERENCES videos(video_key),
            time_range TEXT NOT NULL, chart TEXT NOT NULL, rank INTEGER NOT NULL,
            captured_at TEXT NOT NULL,
            PRIMARY KEY(time_range, chart, rank)
        )""")
        for row_number, row in enumerate(records, 1):
            video_id = clean(row["video_id"])
            key = "taobao:" + video_id
            title = clean(row["title"])
            body = clean(row["copy"])
            creator = clean(row.get("creator"))
            creator_id = clean(row.get("creator_id"))
            has_description = row.get("copy_source") != "title_fallback"
            db.execute("""INSERT INTO videos
                (video_key, platform, video_id, title, body, creator, creator_id, availability)
                VALUES (?, '淘宝光合', ?, ?, ?, ?, ?, '可访问')
                ON CONFLICT(video_key) DO UPDATE SET
                  title=CASE WHEN length(excluded.title)>length(videos.title)
                    THEN excluded.title ELSE videos.title END,
                  body=CASE WHEN ? OR videos.body='' OR videos.body=videos.title
                    THEN excluded.body ELSE videos.body END,
                  creator=CASE WHEN videos.creator='' THEN excluded.creator ELSE videos.creator END,
                  creator_id=CASE WHEN videos.creator_id='' THEN excluded.creator_id ELSE videos.creator_id END,
                  availability='可访问'""",
                (key, video_id, title, body, creator, creator_id, has_description))
            db.execute("""INSERT OR IGNORE INTO video_sources
                (video_key, source_type, source_file, source_row) VALUES (?, ?, ?, ?)""",
                (key, "光合热门榜", source.name, row_number))
            platform_topics = {clean(topic).lstrip("#＃ ") for topic in row.get("topics", [])}
            topics = list(dict.fromkeys(
                list(platform_topics)
                + [clean(match.group(1)) for match in TOPIC_RE.finditer(title + " " + body)]
            ))
            for topic in topics:
                if topic:
                    db.execute("""INSERT OR IGNORE INTO video_topics
                        (video_key, topic, extraction) VALUES (?, ?, ?)""",
                        (key, topic, "平台话题" if topic in platform_topics else "文案#话题"))
            db.execute("""INSERT INTO guanghe_hot_occurrences
                (video_key, time_range, chart, rank, captured_at) VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(time_range, chart, rank) DO UPDATE SET
                  video_key=excluded.video_key, captured_at=excluded.captured_at""",
                (key, clean(row["date_range"]), clean(row["chart"]), int(row["rank"]),
                 clean(data.get("captured_at"))))
    unique_ids = {clean(row["video_id"]) for row in records}
    return {"ranking_rows": len(records), "unique_videos": len(unique_ids),
            "skipped": len(data.get("skipped", []))}


def main() -> None:
    with sqlite3.connect(DATABASE) as db:
        result = import_capture(db)
        video_count = db.execute("SELECT COUNT(*) FROM videos").fetchone()[0]
        topic_count = db.execute("SELECT COUNT(*) FROM video_topics").fetchone()[0]
    check_path = OUT / "数据校验.json"
    if check_path.exists():
        checks = json.loads(check_path.read_text(encoding="utf-8"))
        checks.update(unique_videos=video_count, video_topics=topic_count, guanghe_hot=result)
        check_path.write_text(json.dumps(checks, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"database": str(DATABASE), **result}, ensure_ascii=False))


if __name__ == "__main__":
    main()
