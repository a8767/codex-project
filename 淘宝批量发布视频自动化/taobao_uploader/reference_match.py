"""Choose relevant ranking examples and safe themes for upload drafts.

Reference copy is evidence for matching, never copied into a publish field.
Only activity topics already confirmed in the Taobao publisher may be selected.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import date
import re
import sqlite3


CONFIRMED_TOPICS = ("山地骑行的快乐", "进来一起享受骑行")
GENERIC_TOPIC = "进来一起享受骑行"
THEMES = {
    "通勤": ("通勤", "上下班", "城市", "地铁"),
    "户外": ("户外", "周末", "山野", "山林", "远方", "出发"),
    "日常": ("日常", "生活", "每天", "平时"),
    "细节": ("细节", "看看", "看点", "关注", "第一眼"),
    "骑行": ("骑行", "车友", "喜欢骑", "热爱骑"),
    "记录": ("记录", "镜头", "画面", "片段", "分享"),
}
ACCESSORY_WORDS = (
    "眼镜", "风镜", "墨镜", "头盔", "坐垫", "手套", "车筐", "车篮", "水杯架",
    "车锁", "防盗锁", "安全锁", "车灯", "停车架", "脚撑", "码表", "护膝", "骑行服",
    "轮组", "轮胎", "打气筒", "手机支架", "收纳架", "电动车车筐",
    "车架", "曲柄", "链条", "刹车片", "挡泥板", "骑行裤",
    "存放", "收纳", "小黑盒", "配件", "神器",
)
BIKE_WORDS = ("自行车", "山地车", "公路车", "折叠车", "童车", "单车", "骑行")
WHOLE_BIKE_WORDS = BIKE_WORDS[:-1]
PROMO_WORDS = ("七夕", "双十一", "福利", "特惠", "优惠", "优惠券", "直播", "抽奖", "低价", "百元带走")
OTHER_MODEL_RE = re.compile(r"(?<![A-Za-z0-9])[A-Za-z]{1,10}[._ -]*\d{1,6}[A-Za-z0-9_-]*(?![A-Za-z0-9])")
NICHE_BIKES = ("三轮", "躺车", "平衡车", "电助力", "电动自行车")


@dataclass(frozen=True)
class ReferenceAdvice:
    count: int = 0
    title: str = ""
    body: str = ""
    source: str = ""
    video_key: str = ""
    topic: str = GENERIC_TOPIC
    related_topics: tuple[str, ...] = ()
    themes: tuple[tuple[str, float], ...] = ()


def _table_exists(db: sqlite3.Connection, name: str) -> bool:
    return db.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)).fetchone() is not None


def _model_pattern(model: str) -> re.Pattern[str]:
    chunks = re.findall(r"[A-Za-z]+|\d+", model)
    if not chunks:
        return re.compile(re.escape(model), re.IGNORECASE)
    return re.compile(r"(?<![A-Za-z0-9])" + r"[._ -]*".join(map(re.escape, chunks)) +
                      r"(?![A-Za-z0-9])", re.IGNORECASE)


def _category(title: str, body: str) -> str:
    # The title names the subject more reliably than a long description or tag list.
    if any(word in title for word in ACCESSORY_WORDS) or any(
        word in title for word in ("摩托车", "机车", "电动车")
    ):
        return "配件或非自行车"
    if not any(word in title for word in WHOLE_BIKE_WORDS) and any(
        word in body[:100] for word in ACCESSORY_WORDS
    ):
        return "配件或非自行车"
    text = title + " " + body[:120]
    for category, words in (
        ("儿童自行车", ("儿童自行车", "童车")),
        ("山地车", ("山地自行车", "山地车")),
        ("公路车", ("公路自行车", "公路车")),
        ("折叠车", ("折叠自行车", "折叠车")),
        ("自行车", ("自行车", "单车")),
    ):
        if any(word in text for word in words):
            return category
    return "骑行" if "骑行" in text else ""


def _sources(db: sqlite3.Connection) -> dict[str, dict[str, int]]:
    sources: dict[str, dict[str, int]] = defaultdict(dict)
    if _table_exists(db, "guanghe_hot_occurrences"):
        columns = {row[1] for row in db.execute("PRAGMA table_info(guanghe_hot_occurrences)")}
        if "captured_at" in columns:
            latest_capture = db.execute(
                "SELECT MAX(captured_at) FROM guanghe_hot_occurrences"
            ).fetchone()[0]
            hot_query = """SELECT video_key,MIN(rank),COUNT(*)
                FROM guanghe_hot_occurrences WHERE captured_at=? GROUP BY video_key"""
            hot_rows = db.execute(hot_query, (latest_capture,))
        else:
            hot_rows = db.execute("""SELECT video_key,MIN(rank),COUNT(*)
                FROM guanghe_hot_occurrences GROUP BY video_key""")
        for key, rank, charts in hot_rows:
            sources[key]["hot_rank"] = rank
            sources[key]["hot_charts"] = charts
    if _table_exists(db, "video_sources"):
        quality_rows = db.execute("""SELECT DISTINCT video_key,source_file FROM video_sources
                WHERE source_type='淘宝优质视频'""").fetchall()
        dated_files = {}
        for _, source_file in quality_rows:
            match = re.search(r"(20\d{6})", source_file or "")
            if match:
                dated_files[source_file] = match.group(1)
        latest_quality_date = max(dated_files.values(), default="")
        latest_quality_files = {name for name, date in dated_files.items()
                                if date == latest_quality_date}
        if not latest_quality_files and quality_rows:
            latest_quality_files = {max((file or "") for _, file in quality_rows)}
        for key, source_file in quality_rows:
            sources[key]["quality"] = 1
            if source_file in latest_quality_files:
                sources[key]["quality_latest"] = 1
        # The newly collected Guanghe captions/topics are matching evidence too.
        # Keep them out of publish fields; use their dates for a bounded freshness
        # boost so recent platform examples influence theme/topic ranking.
        if _table_exists(db, "videos"):
            copy_rows = db.execute("""SELECT DISTINCT v.video_key,v.published_at
                FROM videos v JOIN video_sources s USING(video_key)
                WHERE s.source_type='光合发布视频列表' AND v.published_at IS NOT NULL
                  AND v.published_at<>''""").fetchall()
            dates = {}
            for key, published_at in copy_rows:
                date_text = str(published_at)[:10]
                try:
                    parsed = date.fromisoformat(date_text)
                except ValueError:
                    continue
                dates[key] = parsed
                sources[key]["guanghe_copy_reference"] = 1
            newest_copy_date = max(dates.values(), default=None)
            if newest_copy_date:
                for key, published_at in dates.items():
                    age_days = max(0, (newest_copy_date - published_at).days)
                    sources[key]["guanghe_copy_freshness"] = (
                        18 if age_days <= 30 else 12 if age_days <= 90 else 0
                    )
    if _table_exists(db, "rank_occurrences"):
        columns = {row[1] for row in db.execute("PRAGMA table_info(rank_occurrences)")}
        if "month" in columns:
            latest_month = db.execute("SELECT MAX(month) FROM rank_occurrences").fetchone()[0]
            for key, rank, month in db.execute("""SELECT current.video_key,MIN(current.rank),current.month
                    FROM rank_occurrences AS current
                    JOIN (SELECT video_key,MAX(month) AS latest_month FROM rank_occurrences GROUP BY video_key)
                         AS latest ON latest.video_key=current.video_key AND latest.latest_month=current.month
                    GROUP BY current.video_key,current.month"""):
                sources[key]["douyin_rank"] = rank
                sources[key]["douyin_month"] = month
                sources[key]["douyin_latest_month"] = int(month == latest_month)
        else:
            for key, rank in db.execute("""SELECT video_key,MIN(rank) FROM rank_occurrences
                    GROUP BY video_key"""):
                sources[key]["douyin_rank"] = rank
    return sources


def recommend(db: sqlite3.Connection, model: str | None, category: str,
              folder_name: str = "", target_platform: str = "淘宝光合") -> ReferenceAdvice:
    if not {"video_key", "platform", "title", "body"}.issubset(
        {row[1] for row in db.execute("PRAGMA table_info(videos)")}
    ):
        return ReferenceAdvice(topic="山地骑行的快乐" if category == "山地车" else GENERIC_TOPIC)
    sources = _sources(db)
    if not sources:
        return ReferenceAdvice(topic="山地骑行的快乐" if category == "山地车" else GENERIC_TOPIC)

    model_re = _model_pattern(model) if model else None
    candidates = []
    for key, platform, title, body in db.execute("SELECT video_key,platform,title,body FROM videos"):
        meta = sources.get(key)
        if not meta:
            continue
        title, body = title or "", body or ""
        if (len(title) > 100 or any(word in title for word in PROMO_WORDS)
                or any(word in title for word in NICHE_BIKES)):
            continue
        if not any(word in title for word in BIKE_WORDS):
            continue
        exact_model = bool(model_re and model_re.search(title + " " + body[:100]))
        if OTHER_MODEL_RE.search(title) and not exact_model:
            continue
        kind = _category(title, body)
        if not kind or kind == "配件或非自行车":
            continue
        if category and kind not in {category, "自行车", "骑行"}:
            continue
        if not category and kind not in {"自行车", "骑行"}:
            # Unknown models and atmosphere clips must not borrow a specific bike claim.
            continue
        score = 25 if exact_model else 0
        score += 22 if category and kind == category else 7
        same_platform = platform == target_platform
        if same_platform:
            score += 65
        score += meta.get("guanghe_copy_freshness", 0)
        if "hot_rank" in meta:
            score += 50 + 12 * (51 - meta["hot_rank"]) / 50 + min(meta["hot_charts"], 3)
        if meta.get("quality"):
            # The current quality-video source is 淘宝优质视频. It should dominate
            # cross-platform rankings for Taobao publishing, while remaining a
            # low-priority fallback for other platforms.
            score += 85 if target_platform == "淘宝光合" else 8
            if meta.get("quality_latest"):
                score += 12
        if "douyin_rank" in meta:
            # Douyin charts are strong evidence for Douyin and fallback evidence
            # for Taobao, rather than outranking Taobao's own quality examples.
            score += (70 if target_platform == "抖音" else 7) + 2 * (101 - meta["douyin_rank"]) / 100
            if meta.get("douyin_latest_month"):
                score += 5
        if target_platform == "淘宝光合" and platform == "抖音":
            score -= 10
        candidates.append((score, key, title, body, platform, meta))
    candidates.sort(key=lambda item: (-item[0], item[1]))
    if not candidates:
        return ReferenceAdvice(topic="山地骑行的快乐" if category == "山地车" else GENERIC_TOPIC)

    top = candidates[0]
    meta = top[5]
    source = (f"光合热门榜第{meta['hot_rank']}名" if "hot_rank" in meta else
              "淘宝优质视频" if meta.get("quality") else
              "光合发布视频文案/话题参考" if meta.get("guanghe_copy_reference") else
              f"抖音榜单第{meta['douyin_rank']}名")
    theme_scores: Counter[str] = Counter()
    related: Counter[str] = Counter()
    keys = [item[1] for item in candidates[:30]]
    topics_by_key: dict[str, list[tuple[str, str]]] = defaultdict(list)
    if _table_exists(db, "video_topics"):
        placeholders = ",".join("?" for _ in keys)
        extraction_column = ("extraction" if "extraction" in
                             {row[1] for row in db.execute("PRAGMA table_info(video_topics)")} else "''")
        for key, topic, extraction in db.execute(
            f"SELECT video_key,topic,{extraction_column} FROM video_topics WHERE video_key IN ({placeholders})", keys
        ):
            topics_by_key[key].append((topic, extraction))
    for score, key, title, body, platform, meta in candidates[:30]:
        weight = score / 25
        text = title + " " + body[:160]
        for theme, words in THEMES.items():
            if any(word in text for word in words):
                theme_scores[theme] += weight
        if platform == target_platform and target_platform == "淘宝光合":
            for topic, extraction in topics_by_key[key]:
                if (extraction != "平台话题" or
                        not any(word in topic for word in ("骑行", "自行车")) or
                        any(word in topic for word in ACCESSORY_WORDS + PROMO_WORDS +
                            ("年终", "年货", "省钱", "必买", "装备", "好物"))):
                    continue
                related[topic] += weight
        elif platform == target_platform and target_platform == "抖音":
            for topic, extraction in topics_by_key[key]:
                if extraction not in {"标题#话题", "榜单文案#话题"}:
                    continue
                cleaned = topic.lstrip("#＃").strip()
                if (not any(word in cleaned for word in ("骑行", "自行车", "山地车", "公路车"))
                        or any(word in cleaned for word in ACCESSORY_WORDS + PROMO_WORDS)):
                    continue
                related[cleaned] += weight
    if folder_name:
        for theme, words in THEMES.items():
            if any(word in folder_name for word in words):
                theme_scores[theme] += 8
        if "氛围" in folder_name:
            theme_scores["记录"] += 100
    selected_topic = "山地骑行的快乐" if category == "山地车" else GENERIC_TOPIC
    if selected_topic not in CONFIRMED_TOPICS:
        raise ValueError("推荐的话题未经发布器验证")
    return ReferenceAdvice(
        count=len(candidates), title=top[2], body=top[3], source=source, video_key=top[1],
        topic=selected_topic,
        related_topics=tuple(topic for topic, _ in related.most_common(5)),
        themes=tuple(sorted(((theme, weight) for theme, weight in theme_scores.items()
                             if theme != "骑行"), key=lambda item: (-item[1], item[0]))),
    )
