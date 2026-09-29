"""Collect visible Guanghe video titles and topics from the user's active BSK tab."""

from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
import sqlite3
import subprocess
import time


ROOT = Path(__file__).resolve().parents[1]
DB = ROOT / "outputs" / "video_reference_db" / "视频发布参考数据库.sqlite3"
OUTPUT = ROOT / "outputs" / "video_reference_db" / "光合发布视频_2026_文案话题.json"
SESSION = "ynab"
BATCH_PAGES = 10
PAGE_STATE = r"""({
  page: Number(document.querySelector('button.next-pagination-item.next-current')?.getAttribute('aria-label')?.match(/第(\d+)页/)?.[1] || 0),
  totalPages: Number(document.querySelector('button.next-pagination-item.next-current')?.getAttribute('aria-label')?.match(/共(\d+)页/)?.[1] || 0)
})"""

EXTRACT = r"""(async () => {
  const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
  const pageNumber = () => {
    const label = document.querySelector('button.next-pagination-item.next-current')?.getAttribute('aria-label') || '';
    return Number(label.match(/第(\d+)页/)?.[1] || 0);
  };
  const pageTotal = () => {
    const label = document.querySelector('button.next-pagination-item.next-current')?.getAttribute('aria-label') || '';
    return Number(label.match(/共(\d+)页/)?.[1] || 0);
  };
  const rows = () => {
    const table = document.querySelectorAll('table')[1];
    return Array.from(table?.querySelectorAll('tbody tr') || []).map((tr, index) => {
      const cells = tr.children;
      const info = cells[0]?.querySelector('.gg-multi-worksInfo')?.innerText?.trim() || '';
      const headlineLines = (cells[0]?.innerText || '').split(/\nID:/)[0].split(/\n/)
        .map(line => line.trim()).filter(line => line && line !== '置顶' && !/^\d{2}:\d{2}$/.test(line));
      const titleNode = cells[0]?.querySelector('[class*="ContentTitle_title__"]')?.textContent?.trim() || '';
      const title = titleNode || headlineLines.at(-1) || '';
      const id = info.match(/ID:\s*(\d+)/)?.[1] || '';
      const publishedAt = info.match(/(\d{4}-\d{2}-\d{2})(?:\s+(\d{2}:\d{2}:\d{2}))?/) || [];
      const topics = Array.from(cells[2]?.querySelectorAll('.topic span') || [])
        .map(node => node.textContent?.trim() || '').filter(Boolean);
      return {
        row: index + 1,
        video_id: id,
        title,
        title_fallback: !titleNode,
        published_at: publishedAt[1] ? `${publishedAt[1]}${publishedAt[2] ? ` ${publishedAt[2]}` : ''}` : '',
        scheduled: info.includes('定时发布'),
        topics,
        status: cells[1]?.innerText?.trim() || ''
      };
    });
  };
  const totalItems = Number(document.body.innerText.match(/共(\d+)条内容/)?.[1] || 0);
  const from = document.querySelector('input[placeholder="起始日期"]')?.value || '';
  const to = document.querySelector('input[placeholder="结束日期"]')?.value || '';
  const typeField = Array.from(document.querySelectorAll('[class*="RadioButton_field__"]'))
    .find(node => node.innerText.startsWith('作品类型'));
  const selectedType = typeField?.querySelector('[class*="fieldActive"]')?.innerText?.trim() || '';
  const startPage = pageNumber();
  const totalPages = pageTotal();
  const pages = [];
  for (let offset = 0; offset < __BATCH_PAGES__; offset++) {
    const current = pageNumber();
    if (!current || current !== startPage + offset) throw new Error(`分页意外：预期 ${startPage + offset}，当前 ${current}`);
    const items = rows();
    if (!items.length) throw new Error(`第 ${current} 页没有可见记录`);
    pages.push({page: current, items});
    if (current >= totalPages || offset + 1 >= __BATCH_PAGES__) break;
    const previousId = items[0]?.video_id || '';
    const next = document.querySelector('button.next-pagination-item.next-next');
    if (!next || next.disabled) throw new Error(`第 ${current} 页没有可用的下一页按钮`);
    next.click();
    let changed = false;
    for (let attempt = 0; attempt < 80; attempt++) {
      await sleep(250);
      const fresh = rows();
      if (pageNumber() === current + 1 && fresh.length && fresh[0].video_id !== previousId) {
        changed = true;
        break;
      }
    }
    if (!changed) throw new Error(`翻到第 ${current + 1} 页超时`);
  }
  return {ok: true, startPage, totalPages, totalItems, from, to, selectedType, pages};
})()"""


def bsk(*args: str, timeout: int = 120) -> dict:
    result = subprocess.run(
        ["bsk", *args], capture_output=True, text=True, encoding="utf-8",
        errors="replace", timeout=timeout, check=False,
    )
    if result.returncode:
        raise RuntimeError((result.stderr or result.stdout).strip()[:600])
    try:
        payload = json.loads(result.stdout)
    except json.JSONDecodeError:
        return {"ok": True, "stdout": result.stdout.strip()}
    if not payload.get("ok"):
        raise RuntimeError(str(payload.get("error") or payload)[:600])
    return payload


def save(state: dict) -> None:
    temporary = OUTPUT.with_suffix(".json.tmp")
    temporary.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(OUTPUT)


def main() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DB) as db:
        known_ids = {
            str(row[0]) for row in db.execute(
                "SELECT video_id FROM videos WHERE platform='淘宝光合' AND video_id IS NOT NULL"
            ) if row[0]
        }

    if OUTPUT.exists():
        state = json.loads(OUTPUT.read_text(encoding="utf-8"))
        if state.get("complete"):
            print(json.dumps({"complete": True, "output": str(OUTPUT)}, ensure_ascii=False))
            return
        if state.get("content_type") != "视频":
            raise RuntimeError("已有进度文件的类型筛选不是视频")
    else:
        state = {
            "source_url": "https://creator.guanghe.taobao.com/page/workspace/tb",
            "date_range": ["2026-01-01", "2026-09-30"],
            "content_type": "视频",
            "captured_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "total_items": 3385,
            "total_pages": 170,
            "pages": [],
            "records": [],
            "seen_video_ids": [],
            "scanned_items": 0,
            "skipped_existing": 0,
            "skipped_duplicate_ids": 0,
        }
        state.setdefault("seen_video_ids", [])
    state.setdefault("unresolved_rows", [])
    seen_ids: set[str] = set(state.get("seen_video_ids", []))
    page_cursor = max(state.get("pages", [0])) + 1
    while page_cursor <= state["total_pages"]:
        current = bsk("evaluate", "--session", SESSION, "--json", PAGE_STATE)["value"]["page"]
        if current != page_cursor:
            bsk("fill", "--selector", "input[aria-label='请输入跳转到第几页']",
                "--value", str(page_cursor), "--session", SESSION)
            bsk("click", "--selector", "button.next-pagination-jump-go", "--session", SESSION)
            wait_page = rf"""(async () => {{
              for (let i=0; i<80; i++) {{
                const label = document.querySelector('button.next-pagination-item.next-current')?.getAttribute('aria-label') || '';
                const page = Number(label.match(/第(\d+)页/)?.[1] || 0);
                if (page === {page_cursor}) return {{page}};
                await new Promise(resolve => setTimeout(resolve, 250));
              }}
              throw new Error('跳转到第 {page_cursor} 页超时');
            }})()"""
            bsk("evaluate", "--session", SESSION, "--json", wait_page, timeout=40)
        payload = bsk(
            "evaluate", "--session", SESSION, "--json",
            EXTRACT.replace("__BATCH_PAGES__", str(BATCH_PAGES)), timeout=240,
        )
        if payload.get("value", {}).get("startPage") != page_cursor:
            raise RuntimeError(
                f"浏览器当前位于第 {payload.get('value', {}).get('startPage')} 页，预期第 {page_cursor} 页"
            )
        page_result = payload["value"]
        if page_cursor == 1:
            state.update(
                total_items=page_result["totalItems"],
                total_pages=page_result["totalPages"],
                date_range=[page_result.get("from", ""), page_result.get("to", "")],
                content_type=page_result.get("selectedType", ""),
            )
            if state["content_type"] != "视频":
                raise RuntimeError(f"当前作品类型不是视频：{state['content_type']!r}")
        for page in page_result["pages"]:
            state["pages"].append(page["page"])
            state["scanned_items"] += len(page["items"])
            for item in page["items"]:
                video_id = str(item.get("video_id") or "").strip()
                if not video_id:
                    raise RuntimeError(f"第 {page['page']} 页第 {item['row']} 行没有作品 ID")
                if video_id in seen_ids:
                    state["skipped_duplicate_ids"] += 1
                    continue
                seen_ids.add(video_id)
                state["seen_video_ids"].append(video_id)
                if video_id in known_ids:
                    state["skipped_existing"] += 1
                    continue
                if not item.get("title"):
                    item["source_page"] = page["page"]
                    item["reason"] = "作品列表标题为空，待查看作品详情"
                    state["unresolved_rows"].append(item)
                    continue
                item["source_page"] = page["page"]
                state["records"].append(item)
        page_cursor = page_result["pages"][-1]["page"] + 1
        save(state)
        print(json.dumps({
            "pages": f"{page_result['startPage']}-{page_result['pages'][-1]['page']}/{state['total_pages']}",
            "scanned": sum(len(p["items"]) for p in page_result["pages"]),
            "new": len(state["records"]),
            "known_skipped": state["skipped_existing"],
        }, ensure_ascii=False), flush=True)

    if state["pages"] != list(range(1, state["total_pages"] + 1)):
        raise RuntimeError("页面采集范围不连续，已保留进度文件")
    state["unique_list_ids"] = state["scanned_items"] - state["skipped_duplicate_ids"]
    state["platform_count_difference"] = state["total_items"] - state["unique_list_ids"]
    state["page_row_count_difference"] = state["scanned_items"] - state["total_items"]
    state.update(scanned_unique_ids=len(seen_ids), complete=True)
    save(state)
    print(json.dumps({
        "complete": True,
        "pages": len(state["pages"]),
        "scanned_unique_ids": len(seen_ids),
        "list_unique_ids": state["unique_list_ids"],
        "platform_total": state["total_items"],
        "platform_count_difference": state["platform_count_difference"],
        "page_row_count_difference": state["page_row_count_difference"],
        "new_records": len(state["records"]),
        "unresolved_rows": len(state["unresolved_rows"]),
        "skipped_existing": state["skipped_existing"],
        "skipped_duplicate_ids": state["skipped_duplicate_ids"],
        "output": str(OUTPUT),
    }, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
