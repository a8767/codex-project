"""Original, neutral copy for folder-based video drafts.

These are suggestions, not facts extracted from a video. They deliberately avoid
specifications, prices, performance and claims about scenes we have not watched.
"""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from hashlib import sha1
import re
import sqlite3

from .reference_match import THEMES


# A title fragment is prefixed with the identified model/category when available.
# {subject} in the body is a known model/category, or simply "自行车".
COMMON_COPY = (
    ("留住这一刻", "把这段关于{subject}的画面记录下来，慢慢看。"),
    ("今天的骑行记录", "留下一段与{subject}有关的日常记录。"),
    ("一起看看这辆车", "镜头里的{subject}，你会先注意到哪里？"),
    ("骑行生活的一页", "用一段画面，记录与{subject}有关的日常。"),
    ("这一段值得收藏", "关于{subject}的一段影像，留着以后再看。"),
    ("从这一眼开始", "先从画面认识{subject}，细节可以慢慢看。"),
    ("多看一会儿", "这段{subject}的画面，适合放慢一点看。"),
    ("把喜欢留在镜头里", "把喜欢{subject}的心情，留在这段视频里。"),
    ("车与日常", "{subject}也是日常的一部分，记录这一段画面。"),
    ("关于这辆车", "来看看{subject}，也欢迎说说你的第一印象。"),
    ("随手记一段", "一段关于{subject}的记录，分享给同样喜欢骑行的人。"),
    ("今天看点不一样的", "换个角度看看{subject}，也许会有新的发现。"),
    ("把镜头停在这里", "镜头停留的这一刻，记录了{subject}的样子。"),
    ("慢慢欣赏", "关于{subject}的短片，留一点时间慢慢看。"),
    ("骑行从喜欢开始", "喜欢骑行，也喜欢记录与{subject}有关的片段。"),
    ("这一眼的感觉", "看看{subject}，你对这辆车的第一感觉是什么？"),
    ("把日常拍成片段", "日常里的一段{subject}画面，分享给你。"),
    ("再看一遍也喜欢", "这段关于{subject}的画面，值得再看一遍。"),
    ("让画面说话", "先看看{subject}的画面，再聊聊你喜欢的骑行方式。"),
    ("喜欢车的人会懂", "记录{subject}，也记录对骑行的喜欢。"),
    ("骑行灵感来自日常", "把{subject}放进镜头，留住这一段日常。"),
    ("这一段有点意思", "关于{subject}的一段视频，你看到了什么？"),
    ("看车也看心情", "看看{subject}，让骑行的心情从这里开始。"),
    ("今天分享这辆车", "今天分享一段{subject}的画面，欢迎交流感受。"),
    ("一个小小的记录", "用镜头留下{subject}，也留下一点骑行的期待。"),
    ("骑行画面收藏夹", "这一段{subject}的影像，收进骑行画面收藏夹。"),
    ("这一刻想分享", "把{subject}的这一段画面，分享给喜欢自行车的你。"),
    ("简单看看就好", "先简单看看{subject}，关于它你最想了解什么？"),
    ("放慢一点看", "放慢节奏，看看{subject}在镜头里的样子。"),
    ("关于骑行的想法", "从{subject}聊起，你喜欢怎样的骑行日常？"),
    ("把兴趣拍下来", "喜欢{subject}，就把这一段画面拍下来。"),
    ("今天的车友时间", "来看看{subject}，车友们会注意到哪些地方？"),
    ("镜头里的日常", "记录{subject}，让平常的一段画面也有记忆点。"),
    ("看完再聊", "先看这一段{subject}视频，再聊聊你的看法。"),
    ("一次简单分享", "分享关于{subject}的画面，不急着下结论。"),
    ("留给喜欢骑行的人", "这段{subject}的画面，留给同样喜欢骑行的人。"),
    ("这一段刚刚好", "用一段关于{subject}的短片，记录此刻的喜欢。"),
    ("骑行兴趣小记录", "关于{subject}的一个小记录，欢迎一起交流。"),
    ("看看你的关注点", "看过{subject}之后，你最关注哪一处？"),
    ("下一段骑行想象", "从{subject}开始，想想下一次骑行想去哪里。"),
)


CATEGORY_COPY = {
    "山地车": (
        ("看看山地车", "今天分享{subject}，一起看看这辆山地车。"),
        ("山地车的日常", "记录{subject}，也记录喜欢山地车的日常。"),
        ("换个角度看山地车", "换个角度看看{subject}，说说你的第一印象。"),
        ("山地车画面记录", "把{subject}的画面留在这里，慢慢欣赏。"),
        ("聊聊山地骑行", "从{subject}聊起，分享你喜欢的山地骑行。"),
        ("今天关注这辆山地车", "今天的镜头留给{subject}，欢迎车友交流。"),
        ("山地车的小片段", "一段关于{subject}的短片，记录骑行兴趣。"),
        ("从这辆山地车说起", "看过{subject}，你最想先了解什么？"),
    ),
    "公路车": (
        ("看看公路车", "今天分享{subject}，一起看看这辆公路车。"),
        ("公路车的日常", "记录{subject}，也记录喜欢公路车的日常。"),
        ("换个角度看公路车", "换个角度看看{subject}，说说你的第一印象。"),
        ("公路车画面记录", "把{subject}的画面留在这里，慢慢欣赏。"),
        ("聊聊公路骑行", "从{subject}聊起，分享你喜欢的公路骑行。"),
        ("今天关注这辆公路车", "今天的镜头留给{subject}，欢迎车友交流。"),
        ("公路车的小片段", "一段关于{subject}的短片，记录骑行兴趣。"),
        ("从这辆公路车说起", "看过{subject}，你最想先了解什么？"),
    ),
    "折叠车": (
        ("看看折叠车", "今天分享{subject}，一起看看这辆折叠车。"),
        ("折叠车的日常", "记录{subject}，也记录喜欢折叠车的日常。"),
        ("换个角度看折叠车", "换个角度看看{subject}，说说你的第一印象。"),
        ("折叠车画面记录", "把{subject}的画面留在这里，慢慢欣赏。"),
        ("聊聊折叠车", "从{subject}聊起，分享你喜欢的骑行方式。"),
        ("今天关注这辆折叠车", "今天的镜头留给{subject}，欢迎车友交流。"),
        ("折叠车的小片段", "一段关于{subject}的短片，记录骑行兴趣。"),
        ("从这辆折叠车说起", "看过{subject}，你最想先了解什么？"),
    ),
    "儿童自行车": (
        ("看看儿童自行车", "今天分享{subject}，一起看看这辆儿童自行车。"),
        ("童车画面记录", "把{subject}的画面留在这里，慢慢欣赏。"),
        ("换个角度看童车", "换个角度看看{subject}，说说你的第一印象。"),
        ("关于这辆童车", "今天的镜头留给{subject}，欢迎交流。"),
        ("儿童自行车小片段", "一段关于{subject}的短片，记录这一刻。"),
        ("从这辆童车说起", "看过{subject}，你最想先了解什么？"),
        ("再看看这辆童车", "来看看{subject}，留下你的真实感受。"),
        ("童车的日常记录", "记录{subject}，留下一段关于自行车的画面。"),
    ),
    "自行车": (
        ("看看自行车", "今天分享{subject}，一起看看这辆车。"),
        ("自行车画面记录", "把{subject}的画面留在这里，慢慢欣赏。"),
        ("换个角度看车", "换个角度看看{subject}，说说你的第一印象。"),
        ("关于这辆自行车", "今天的镜头留给{subject}，欢迎车友交流。"),
        ("自行车小片段", "一段关于{subject}的短片，记录骑行兴趣。"),
        ("从这辆车说起", "看过{subject}，你最想先了解什么？"),
        ("再看看这辆车", "来看看{subject}，留下你的真实感受。"),
        ("车友日常记录", "记录{subject}，也记录对自行车的喜欢。"),
    ),
}

# Original prompts built around the recurring angles in the reference rankings.
# Questions avoid claiming that the unseen footage proves a feature or a scene.
RANK_STYLE_COPY = (
    ("通勤会考虑它吗", "如果把{subject}加入通勤备选，你会先看哪些细节？"),
    ("日常骑行怎么选", "聊聊{subject}，你更在意日常骑行的哪一面？"),
    ("周末骑行的备选", "想安排一次周末骑行时，你会怎样考虑{subject}？"),
    ("户外骑行想法", "从{subject}聊起，你期待怎样的户外骑行？"),
    ("第一眼看哪里", "第一次看{subject}，你最先留意的是哪里？"),
    ("多看几个细节", "关于{subject}，有哪些细节值得继续了解？"),
    ("留一段骑行记录", "用这段画面记录{subject}，也聊聊你的骑行想法。"),
    ("骑行日常分享", "分享一段关于{subject}的记录，欢迎车友交流。"),
    ("你会怎么骑", "看过{subject}，你会想用它体验哪种骑行方式？"),
    ("从喜欢骑行开始", "关于{subject}，你最想和车友分享什么？"),
    ("镜头里的车", "把{subject}放进镜头，看看你会注意到什么。"),
    ("下一次想去哪", "看到{subject}，你会想到下一次去哪里骑行？"),
)

COPY_ENDINGS = (
    "欢迎聊聊你的看法。", "也欢迎分享你的骑行体验。", "留言说说你会怎么选。",
    "车友们，你们会关注哪里？", "期待听听你的想法。", "欢迎在评论区交流。",
    "你最在意哪个方面？", "也可以分享你的日常选择。", "说说你对这段画面的感受。",
    "一起聊聊骑行里的小发现。", "如果是你，会怎么考虑？", "把你的骑行灵感也留下来。",
)


@dataclass(frozen=True)
class CopyTemplate:
    template_id: str
    category: str
    title_fragment: str
    body_template: str


@dataclass(frozen=True)
class BikeVideoCopy:
    copy_id: str
    direction: str
    video_theme: str
    title: str
    body: str
    usage_scenario: str
    keywords: str = ""
    hashtags: str = ""


BIKE_THEME_WORDS = {
    "城市通勤": ("通勤", "上下班", "代步", "城市"),
    "山地骑行": ("山地", "户外", "路况", "坡"),
    "折叠收纳": ("折叠", "收纳", "小空间"),
    "新手入门": ("新手", "入门", "选车", "预算"),
    "亲子出行": ("亲子", "儿童", "童车", "孩子"),
    "公路骑行": ("公路", "运动", "训练"),
    "城市休闲": ("休闲", "周末", "短途"),
    "车型对比": ("对比", "攻略", "参数", "比较"),
    "轻便搬运": ("搬运", "上楼", "重量", "轻便"),
    "日常维护": ("维护", "细节", "装配", "保养"),
    "预算有限": ("预算", "价格", "性价比", "平价"),
    "购买决策": ("购买", "下单", "售后", "配送"),
    "清晨骑行": ("清晨", "早晨", "晨光"),
    "下班通勤": ("下班", "夜骑", "夜晚", "街灯"),
    "周末郊游": ("郊游", "郊外", "乡间", "周末"),
    "河岸绿道": ("河岸", "河边", "绿道"),
    "独自骑行": ("独自", "漫游", "街巷"),
    "朋友结伴": ("朋友", "搭子", "同行", "车友"),
    "亲子骑行": ("亲子", "陪伴", "孩子"),
    "雨后街道": ("雨后", "雨停", "云开"),
    "林间骑行": ("林间", "林荫", "树林", "自然"),
    "骑行记录": ("记录", "旅途", "风景", "片段"),
    "日常代步": ("日常", "生活", "代步", "买菜"),
    "夕阳骑行": ("夕阳", "黄昏", "傍晚", "归途"),
}
ATMOSPHERE_INTENT_WORDS = ("氛围", "风景", "夜骑", "清晨", "夕阳", "郊游", "旅拍",
                           "骑游", "河岸", "绿道", "郊外", "林间", "漫游")
SALES_INTENT_WORDS = ("带货", "种草", "选购", "测评", "价格", "性价比", "配置", "参数",
                      "推荐", "购买", "下单", "对比")
SPECIFIC_BIKE_WORDS = {
    "儿童自行车": ("儿童自行车", "儿童", "童车", "孩子"),
    "山地车": ("山地自行车", "山地车", "山地骑行"),
    "公路车": ("公路自行车", "公路车", "公路骑行"),
    "折叠车": ("折叠自行车", "折叠车", "折叠收纳"),
}
PRICE_INTENT_RE = re.compile(r"(?:只要|仅需|不到|低至)\s*\d{1,6}|[¥￥]\s*\d|\d{1,6}\s*(?:元|块)")
BUDGET_INTENT_RE = re.compile(
    r"(?:性价比|平价|预算|价格|只要|仅需|不到|低至|[¥￥]\s*\d|\d{1,6}\s*(?:元|块))"
)
BIKE_ACCESSORY_WORDS = (
    "骑行眼镜", "眼镜", "风镜", "墨镜", "头盔", "坐垫套", "坐垫", "手套",
    "车筐", "车篮", "水杯架", "车锁", "防盗锁", "安全锁", "车灯", "脚撑",
    "停车架", "码表", "护膝", "骑行服", "轮胎", "打气筒", "手机支架",
    "挡泥板", "骑行裤", "水壶", "把套", "铃铛", "刹车", "脚踏", "链条",
    "车包", "尾包", "工具包", "手泵", "迷你泵", "车架", "曲柄", "刹车片",
    "轮组", "前叉", "牙盘", "车把", "车座", "变速器", "飞轮", "尾灯", "前灯",
)
BIKE_ACCESSORY_GROUPS = (
    ("骑行眼镜", "眼镜", "风镜", "墨镜"),
    ("坐垫套", "坐垫", "车座"),
    ("车筐", "车篮"),
    ("车锁", "防盗锁", "安全锁"),
    ("车灯", "前灯", "尾灯"),
    ("打气筒", "手泵", "迷你泵"),
    ("车包", "尾包", "工具包"),
    ("头盔",), ("手套",), ("水杯架",), ("脚撑",), ("停车架",), ("码表",),
    ("护膝",), ("骑行服",), ("轮胎",), ("手机支架",), ("挡泥板",),
    ("骑行裤",), ("水壶",), ("把套",), ("铃铛",), ("刹车",), ("脚踏",),
    ("链条",), ("车架",), ("曲柄",), ("刹车片",), ("轮组",), ("前叉",),
    ("牙盘",), ("车把",), ("变速器",), ("飞轮",),
)


def _accessory_groups(text: str) -> set[int]:
    return {index for index, group in enumerate(BIKE_ACCESSORY_GROUPS)
            if any(word in text for word in group)}


def is_bike_accessory_copy(title: str, body: str, theme: str = "", keywords: str = "") -> bool:
    wording = f"{title} {body} {theme} {keywords}"
    return any(word in wording for word in BIKE_ACCESSORY_WORDS)


def has_named_copy_theme(folder_name: str, video_name: str) -> bool:
    """Whether filenames state an actual scene or buying theme for matching."""
    signals = f"{folder_name} {video_name}"
    return any(word in signals for words in BIKE_THEME_WORDS.values() for word in words)


def default_templates() -> tuple[CopyTemplate, ...]:
    templates = [CopyTemplate(f"common-{i:02d}", "通用", title, body)
                 for i, (title, body) in enumerate(COMMON_COPY, 1)]
    for category, pairs in CATEGORY_COPY.items():
        templates.extend(CopyTemplate(f"{category}-{i:02d}", category, title, body)
                         for i, (title, body) in enumerate(pairs, 1))
    templates.extend(CopyTemplate(f"rank-style-{i:02d}", "通用", title, body)
                     for i, (title, body) in enumerate(RANK_STYLE_COPY, 1))
    return tuple(templates)


def seed_copy_templates(connection: sqlite3.Connection) -> int:
    """Add built-in copy without altering collected videos or user-edited templates."""
    connection.execute("""CREATE TABLE IF NOT EXISTS original_copy_templates (
        template_id TEXT PRIMARY KEY,
        category TEXT NOT NULL,
        title_fragment TEXT NOT NULL,
        body_template TEXT NOT NULL,
        source TEXT NOT NULL DEFAULT '原创草稿'
    )""")
    connection.executemany(
        "INSERT OR IGNORE INTO original_copy_templates VALUES(?,?,?,?,?)",
        ((item.template_id, item.category, item.title_fragment,
          item.body_template, "原创草稿") for item in default_templates()),
    )
    connection.commit()
    return connection.execute("SELECT count(*) FROM original_copy_templates").fetchone()[0]


def load_copy_templates(connection: sqlite3.Connection) -> dict[str, tuple[CopyTemplate, ...]]:
    try:
        rows = connection.execute("""SELECT template_id,category,title_fragment,body_template
            FROM original_copy_templates ORDER BY template_id""").fetchall()
        templates = tuple(CopyTemplate(*row) for row in rows)
    except sqlite3.OperationalError as exc:
        if "no such table" not in str(exc):
            raise
        templates = default_templates()
    grouped: defaultdict[str, list[CopyTemplate]] = defaultdict(list)
    for item in templates:
        grouped[item.category].append(item)
    return {key: tuple(value) for key, value in grouped.items()}


def load_bike_video_copies(connection: sqlite3.Connection) -> tuple[BikeVideoCopy, ...]:
    """Load curated bicycle copy and bicycle-relevant captions from the video table."""
    copies: list[BikeVideoCopy] = []
    sources = (
        ("""SELECT copy_id,direction,video_theme,title,body,usage_scenario,''
           FROM bike_video_copies ORDER BY copy_id""",),
        ("""SELECT copy_id,direction,video_theme,title,body,angle,keywords
           FROM original_copy_bank ORDER BY copy_id""",),
    )
    for (query,) in sources:
        try:
            copies.extend(BikeVideoCopy(*row) for row in connection.execute(query).fetchall())
        except sqlite3.OperationalError as exc:
            if "no such table" not in str(exc):
                raise
    copies.extend(load_reference_video_copies(connection))
    # Curated text takes precedence over identical platform captions.
    seen_ids: set[str] = set()
    seen_text: set[str] = set()
    unique: list[BikeVideoCopy] = []
    for item in copies:
        text_key = re.sub(r"\s+", "", f"{item.title}|{item.body}").casefold()
        if item.copy_id in seen_ids or text_key in seen_text:
            continue
        seen_ids.add(item.copy_id)
        seen_text.add(text_key)
        unique.append(item)
    return tuple(unique)


def load_reference_video_copies(connection: sqlite3.Connection) -> tuple[BikeVideoCopy, ...]:
    """Read bicycle-related title/body/topic examples from the normalized video table."""
    tables = {row[0] for row in connection.execute(
        "SELECT name FROM sqlite_master WHERE type='table'")}
    if "videos" not in tables:
        return ()
    columns = {row[1] for row in connection.execute("PRAGMA table_info(videos)")}
    if not {"video_key", "title"}.issubset(columns):
        return ()
    body = "COALESCE(NULLIF(v.body,''),v.title)" if "body" in columns else "v.title"
    published = "v.published_at" if "published_at" in columns else "''"
    if "video_topics" in tables:
        topic_rows = "COALESCE(group_concat(DISTINCT t.topic),'')"
        join = "LEFT JOIN video_topics t USING(video_key)"
    else:
        topic_rows, join = "''", ""
    query = f"""SELECT v.video_key,v.platform,v.title,{body},{topic_rows},{published}
        FROM videos v {join} GROUP BY v.video_key ORDER BY {published} DESC,v.video_key"""
    hashtag_pattern = re.compile(r"[#＃]\s*([^\s#＃,，。！？!?；;、]+)")
    rank_sources: defaultdict[str, list[str]] = defaultdict(list)
    if "rank_occurrences" in tables:
        rank_columns = {row[1] for row in connection.execute("PRAGMA table_info(rank_occurrences)")}
        if {"video_key", "chart", "month", "rank"}.issubset(rank_columns):
            for key, chart, month, rank in connection.execute(
                "SELECT video_key,chart,month,rank FROM rank_occurrences ORDER BY month DESC,rank"
            ):
                source = f"榜单月度记录：{chart} 2026-{int(month):02d} 第{rank}名"
                if source not in rank_sources[key]:
                    rank_sources[key].append(source)
    bike_words = ("自行车", "山地车", "山地自行车", "公路车", "公路自行车",
                  "折叠车", "折叠自行车", "儿童自行车", "童车", "单车", "骑行")
    excluded = ("摩托车", "电动自行车", "电助力", "平衡车", "三轮车")
    copies: list[BikeVideoCopy] = []
    for key, _platform, raw_title, raw_body, topic_text, _published_at in connection.execute(query):
        raw_title = (raw_title or "").strip()
        raw_body = (raw_body or raw_title).strip()
        topic_values = [value.strip().lstrip("#＃") for value in (topic_text or "").split(",")
                        if value.strip().lstrip("#＃")]
        extracted = hashtag_pattern.findall(f"{raw_title} {raw_body}")
        hashtags = list(dict.fromkeys([tag.strip().lstrip("#＃") for tag in extracted] + topic_values))
        title = re.sub(r"[#＃]", "", hashtag_pattern.sub("", raw_title))
        body_text = re.sub(r"[#＃]", "", hashtag_pattern.sub("", raw_body))
        title = title.strip(" \t\r\n，,。.!！?？；;、")
        body_text = body_text.strip(" \t\r\n，,。.!！?？；;、")
        body_text = body_text or title
        combined = f"{title} {body_text} {' '.join(topic_values)}"
        if not title or not any(word in combined for word in bike_words):
            continue
        if any(word in combined for word in excluded) and not any(
            word in combined for word in ("山地自行车", "公路自行车", "折叠自行车", "儿童自行车")
        ):
            continue
        atmosphere_hits = sum(word in combined for word in ATMOSPHERE_INTENT_WORDS)
        sales_hits = sum(word in combined for word in SALES_INTENT_WORDS)
        direction = ("氛围感向" if atmosphere_hits > sales_hits else
                     "带货种草向" if sales_hits > atmosphere_hits else
                     "带货种草向" if sha1(str(key).encode("utf-8")).digest()[0] % 2 else "氛围感向")
        tags = "；".join(topic_values)
        source_note = "；".join(rank_sources.get(key, ()))
        scenario = "视频总表自行车相关历史文案"
        if source_note:
            scenario += f"；{source_note}"
        scenario += "；作为表达参考，发布前核对画面与商品信息"
        copies.append(BikeVideoCopy(
            f"video-{key}", direction, tags or "视频总表参考", title, body_text,
            f"{scenario}；话题：{tags or '无'}", tags,
            " ".join(f"#{tag}" for tag in hashtags),
        ))
    return tuple(copies)


def suggest_bike_video_copy(
    copies: tuple[BikeVideoCopy, ...], category: str, folder_name: str, video_name: str,
    index: int, themes: tuple[tuple[str, float], ...], used_copy_ids: set[str],
    rotation_key: str,
) -> BikeVideoCopy | None:
    """Pick an unused, relevant copy; rotate close matches between draft runs."""
    if not copies:
        return None

    signals = f"{folder_name} {video_name}"
    budget_intent = bool(BUDGET_INTENT_RE.search(signals))
    if any(word in signals for word in ATMOSPHERE_INTENT_WORDS):
        direction = "氛围感向"
    elif any(word in signals for word in SALES_INTENT_WORDS) or PRICE_INTENT_RE.search(signals):
        direction = "带货种草向"
    else:
        # Unknown footage alternates between the two requested directions.
        direction = "带货种草向" if index % 2 else "氛围感向"

    matching_category = category or next(
        (name for name, words in SPECIFIC_BIKE_WORDS.items()
         if any(word in signals for word in words)),
        "",
    )
    candidates: list[tuple[float, str, BikeVideoCopy]] = []
    for item in copies:
        if item.copy_id in used_copy_ids or item.direction != direction:
            continue
        copy_text = (f"{item.video_theme} {item.title} {item.body} "
                     f"{item.usage_scenario} {item.keywords} {item.hashtags}")
        is_video_reference = item.copy_id.startswith("video-")
        copy_terms = f"{item.title} {item.body} {item.video_theme} {item.keywords} {item.hashtags}"
        copy_accessory_groups = _accessory_groups(copy_terms)
        signal_accessory_groups = _accessory_groups(signals)
        if copy_accessory_groups and not signal_accessory_groups:
            continue
        if signal_accessory_groups and not (copy_accessory_groups & signal_accessory_groups):
            continue
        if is_video_reference:
            # Never inherit a historical price/specification from another video.
            if PRICE_INTENT_RE.search(copy_text) or re.search(
                r"(?:[¥￥]\s*\d|\d+\s*(?:元|块|寸|英寸|速|档|挡|斤|公斤|kg|公里)|[A-Za-z]{1,10}[._ -]*\d{1,6}[A-Za-z0-9_-]*)",
                copy_text,
                re.IGNORECASE,
            ):
                continue
        promotional = any(word in f"{item.video_theme} {item.keywords} {item.title}"
                          for word in ("限时活动", "直播福利", "大促", "优惠"))
        if promotional and not any(word in signals for word in ("活动", "直播", "优惠", "大促", "限时")):
            continue
        specific_category = next((name for name, words in SPECIFIC_BIKE_WORDS.items()
                                  if any(word in copy_text for word in words)), "")
        if specific_category and matching_category != specific_category:
            continue

        score = 0.0
        for theme, weight in themes:
            if theme == "骑行":
                continue
            aliases = THEMES.get(theme, ())
            if any(word in copy_text for word in aliases):
                score += min(max(weight, 0.0), 4.0)
        if budget_intent:
            # A concrete affordability/price cue takes precedence over other
            # filename words (for example, a bike type in a hashtag).
            if any(word in copy_text for word in BIKE_THEME_WORDS["预算有限"]):
                score += 120.0
            direct_budget_terms = tuple(word for word in
                                        ("性价比", "预算", "价位", "价格", "平价")
                                        if word in signals)
            if any(word in f"{item.video_theme} {item.title} {item.keywords}"
                   for word in direct_budget_terms):
                score += 160.0
            elif PRICE_INTENT_RE.search(signals) and any(
                    word in f"{item.video_theme} {item.title} {item.keywords}"
                    for word in ("预算", "价格", "价位")):
                score += 120.0
        else:
            for theme, words in BIKE_THEME_WORDS.items():
                if any(word in signals for word in words) and any(word in copy_text for word in words):
                    score += 40.0
        if is_video_reference and score <= 0 and not has_named_copy_theme(folder_name, video_name):
            score = 1.0

        # Prefer relevant items, but rotate within a near-match band so repeated
        # draft generation does not keep selecting the same first template.
        rotation = sha1(f"{rotation_key}:{item.copy_id}".encode("utf-8")).hexdigest()
        candidates.append((score, rotation, item))

    if not candidates:
        return None
    best_score = max(score for score, _, _ in candidates)
    if best_score <= 0:
        return None
    band = max(3.0, best_score * 0.2)
    shortlist = [(rotation, item) for score, rotation, item in candidates
                 if score >= best_score - band]
    shortlist.sort(key=lambda pair: pair[0])
    return shortlist[0][1]


def suggest_copy(templates: dict[str, tuple[CopyTemplate, ...]], category: str,
                 model: str | None, index: int,
                 themes: tuple[tuple[str, float], ...] = ()) -> tuple[str, str]:
    common = templates.get("通用", ())
    specific = templates.get(category, ()) if model else ()
    if not common:
        raise ValueError("文案库缺少通用文案")
    pool: list[CopyTemplate] = []
    for position, item in enumerate(common, 1):
        pool.append(item)
        if position % 4 == 0 and position // 4 <= len(specific):
            pool.append(specific[position // 4 - 1])
    pool.extend(specific[min(len(specific), len(common) // 4):])
    if themes:
        def relevance(item: CopyTemplate) -> float:
            wording = item.title_fragment + " " + item.body_template
            thematic = sum(weight for theme, weight in themes if theme != "骑行"
                           if any(word in wording for word in THEMES.get(theme, ())))
            return thematic + (3 if model and item.category == category else 0)

        pool.sort(key=lambda item: (-relevance(item), item.template_id))
    selected = pool[(index - 1) % len(pool)]
    subject = category or "自行车"
    suffix = selected.title_fragment
    cycle = (index - 1) // len(pool)
    title = suffix[:30]
    body = selected.body_template.format(subject=subject)
    if cycle:
        body += COPY_ENDINGS[(cycle - 1) % len(COPY_ENDINGS)]
    return title, body
