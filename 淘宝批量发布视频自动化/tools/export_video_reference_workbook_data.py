"""Prepare spreadsheet rows from the normalized reference database."""

import collections
import json
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from taobao_uploader.copy_library import (
    is_bike_accessory_copy,
    load_reference_video_copies,
)


OUT = ROOT / "outputs" / "video_reference_db"
db = sqlite3.connect(OUT / "视频发布参考数据库.sqlite3")
db.row_factory = sqlite3.Row


def rows(sql, params=()):
    return [list(r) for r in db.execute(sql, params)]


def optional_rows(sql, params=()):
    """Return [] when an optional table has not been imported yet."""
    try:
        return rows(sql, params)
    except sqlite3.OperationalError:
        return []


stats = json.loads((OUT / "数据校验.json").read_text(encoding="utf-8"))
video_rows = rows("""
    SELECT v.platform, v.title, v.body, v.creator, v.published_at,
           COALESCE((SELECT group_concat(t.topic,'；') FROM video_topics t WHERE t.video_key=v.video_key),''),
           COALESCE((SELECT group_concat(DISTINCT r.chart) FROM rank_occurrences r WHERE r.video_key=v.video_key),''),
           COALESCE((SELECT group_concat(DISTINCT s.term) FROM video_search_terms s WHERE s.video_key=v.video_key),''),
           v.video_id, v.video_url, v.shop_name, v.account_type, v.relevance_status, v.video_key
    FROM videos v ORDER BY CASE WHEN v.platform='淘宝光合' THEN 0 ELSE 1 END,
        CASE WHEN EXISTS(SELECT 1 FROM rank_occurrences r WHERE r.video_key=v.video_key) THEN 0 ELSE 1 END,
        v.published_at DESC, v.video_key
""")
rank_rows = rows("""
    SELECT r.chart, '2026-' || printf('%02d',r.month), r.rank, v.title, v.body, v.creator,
           v.published_at, v.video_id, r.video_type, r.metrics_json, v.video_key
    FROM rank_occurrences r JOIN videos v USING(video_key)
    ORDER BY r.chart,r.month,r.rank
""")
topic_rows = rows("""
    SELECT t.topic,v.platform,v.title,v.creator,v.video_id,v.video_key,t.extraction
    FROM video_topics t JOIN videos v USING(video_key)
    ORDER BY t.topic,v.platform,v.video_key
""")
topic_stats = rows("""
    SELECT t.topic,count(DISTINCT t.video_key) video_count,
           sum(v.platform='抖音') douyin_videos,sum(v.platform='淘宝光合') taobao_videos
    FROM video_topics t JOIN videos v USING(video_key)
    GROUP BY t.topic ORDER BY video_count DESC,t.topic
""")
related_rows = rows("""
    SELECT s.chart,'2026-' || printf('%02d',s.month),s.rank,s.display_order,s.term,
           v.title,v.creator,v.video_id,v.video_key,s.scope
    FROM video_search_terms s JOIN videos v USING(video_key)
    ORDER BY s.month,s.rank,s.display_order
""")
term_rows = rows("""
    SELECT term,source_count,source_files FROM platform_search_terms ORDER BY source_count DESC,term
""")
term_raw_rows = rows("""
    SELECT term,search_result_exposure,search_users,search_payment,product_exposure,
           product_ctr,product_conversion,source_file,source_row
    FROM platform_search_term_records ORDER BY term,source_file,source_row
""")
scripts = rows("""
    SELECT s.chart,'2026-' || printf('%02d',s.month),s.rank,v.title,v.video_id,
           v.video_url,s.status,s.transcript,s.evidence,v.video_key
    FROM script_checks s JOIN videos v USING(video_key)
    ORDER BY s.chart,s.month,s.rank
""")
bike_copies = optional_rows("""
    SELECT copy_id,direction,video_theme,title,body,usage_scenario
    FROM bike_video_copies ORDER BY copy_id
""")
copy_bank = optional_rows("""
    SELECT copy_id,direction,video_theme,title,body,angle,COALESCE(keywords,''),source
    FROM original_copy_bank ORDER BY copy_id
""")
reference_copies = load_reference_video_copies(db)
reference_by_key = {item.copy_id.removeprefix("video-"): item for item in reference_copies}
reference_video_keys = set(reference_by_key)
filtered_video_rows = []
for row in video_rows:
    item = reference_by_key.get(row[13])
    if item is None:
        continue
    row[1], row[2], row[5] = item.title, item.body, item.hashtags
    filtered_video_rows.append(row)
video_rows = filtered_video_rows
filtered_rank_rows = []
for row in rank_rows:
    item = reference_by_key.get(row[10])
    if item is None:
        continue
    filtered_rank_rows.append([*row[:3], item.title, item.body, item.hashtags, *row[5:]])
rank_rows = filtered_rank_rows
stats["bicycle_reference_videos"] = len(reference_video_keys)
stats["bicycle_reference_rank_rows"] = len(rank_rows)


def split_copy_rows():
    """Build clean bike and accessory libraries with body and hashtags separated."""
    bike_rows, accessory_rows = [], []
    curated = [
        [copy_id, direction, theme, title, body, angle, keywords, source or "原创草稿", ""]
        for copy_id, direction, theme, title, body, angle, keywords, source in copy_bank
    ]
    curated += [
        [copy_id, direction, theme, title, body, scenario,
         "；".join(part.strip() for part in theme.replace("／", "/").split("/") if part.strip()),
         "原创草稿", ""]
        for copy_id, direction, theme, title, body, scenario in bike_copies
    ]
    references = [
        [item.copy_id, item.direction, item.video_theme, item.title, item.body,
         item.usage_scenario, item.keywords, "视频总表/榜单月度记录参考文案", item.hashtags]
        for item in reference_copies
    ]
    for row in curated + references:
        target = accessory_rows if is_bike_accessory_copy(row[3], row[4], row[2], row[6]) else bike_rows
        target.append(row)
    order = {"带货种草向": 0, "氛围感向": 1}
    for values in (bike_rows, accessory_rows):
        values.sort(key=lambda row: (order.get(row[1], 9), row[0]))
    return bike_rows, accessory_rows


copy_rows, accessory_copy_rows = split_copy_rows()

payload = {
    "stats": stats,
    "sheets": [
        {"name": "视频总表", "headers": ["平台","视频标题/文案","去话题正文","发布账号","发布时间","#话题","上榜类型","关联搜索词（页面前两项）","视频ID","视频链接","店铺","账号类型","行业相关性","数据库键"], "rows": video_rows},
        {"name": "榜单月度记录", "headers": ["榜单","月份","原始排名","视频标题/文案","去话题正文","#话题","发布账号","发布时间","视频ID","账号类型","页面指标JSON","数据库键"], "rows": rank_rows},
        {"name": "视频话题明细", "headers": ["话题","平台","所属视频标题或文案","发布账号","视频ID","数据库键","提取方式"], "rows": topic_rows},
        {"name": "话题频次", "headers": ["话题","关联视频数","抖音视频数","淘宝视频数"], "rows": topic_stats},
        {"name": "看后搜关联词", "headers": ["榜单","月份","原始排名","页面展示顺序","关联搜索词","视频标题或文案","发布账号","视频ID","数据库键","收集范围"], "rows": related_rows},
        {"name": "平台搜索词", "headers": ["搜索词","出现文件数","来源文件"], "rows": term_rows},
        {"name": "平台热词原始", "headers": ["搜索词","搜索结果曝光人数","搜索人数","搜索用户支付金额","商品曝光人数","商品点击率（人数）","商品点击-成交转化率（人数）","来源文件","Excel原始行"], "rows": term_raw_rows},
        {"name": "视频脚本核验", "headers": ["榜单","月份","原始排名","视频标题或文案","视频ID","视频链接","脚本状态","已取得内容（见状态）","核验依据","数据库键"], "rows": scripts},
        {"name": "自行车视频文案", "headers": ["文案编号","文案方向","适用视频主题","参考标题/文案原文","文案正文（不含#话题）","创作思路／使用场景","关键词","来源","#话题"], "rows": copy_rows},
        {"name": "自行车配件文案", "headers": ["文案编号","文案方向","适用配件/主题","参考标题/文案原文","文案正文（不含#话题）","创作思路／使用场景","关键词","来源","#话题"], "rows": accessory_copy_rows},
    ],
}
(OUT / "workbook_data.json").write_text(json.dumps(payload,ensure_ascii=False,separators=(",", ":")),encoding="utf-8")
print(json.dumps({s["name"]: len(s["rows"]) for s in payload["sheets"]}, ensure_ascii=False))
db.close()
