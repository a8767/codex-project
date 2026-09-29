"""Read the visible Guanghe bicycle-industry hot charts through a BSK session.

Only DOM-backed page text and browser clicks are used; no private endpoints are
called. The JSON is a resumable capture that is imported only after validation.
"""

from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
import re
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from taobao_uploader.browser import BskBrowser, BrowserError
from taobao_uploader.tasks import TZ


URL = "https://creator.guanghe.taobao.com/page/unify/inspiration?tab=1"
OUTPUT = ROOT / "outputs" / "video_reference_db" / "光合热门榜_最近7天_采集.json"
CHARTS = ("热度榜-查看人数", "引流榜-商品点击人数", "成交榜-种草成交金额")
NEXT = ".g-pc-gh-content-dialog-pagination .g-pc-gh-content-dialog-btn:last-child"
CLOSE = ".g-pc-gh-content-dialog-extra > .g-pc-gh-content-dialog-btn:first-child"
EXTRACT = r"""(() => {
  const d = document.querySelector('[role="dialog"]');
  if (!d) return null;
  const text = (selector) => d.querySelector(selector)?.textContent?.trim() || '';
  const nums = Array.from(d.querySelectorAll('.g-pc-gh-content-dialog-user-numwrap'));
  const number = (label) => nums.find(e => e.textContent?.startsWith(label))?.querySelector('span')?.textContent?.trim() || '';
  return {
    position: text('.g-pc-gh-content-dialog-total'),
    video_id: number('作品ID'),
    creator_id: number('账号ID'),
    creator: text('.g-pc-gh-content-dialog-user-nick'),
    title: text('.g-pc-gh-content-dialog-content-title'),
    copy: text('.g-pc-gh-content-dialog-content-summary'),
    topics: Array.from(d.querySelectorAll('.g-pc-gh-content-dialog-content-topicsTagText'))
      .map(e => e.textContent?.trim()).filter(Boolean)
  };
})()"""


def save(records: list[dict], skipped: list[dict]) -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    data = {
        "source_url": URL,
        "industry": "自行车/骑行装备/零配件",
        "date_range": "最近7天",
        "price_range": "全部价格",
        "product_type": "全部",
        "channel": "全部",
        "captured_at": datetime.now(TZ).isoformat(timespec="seconds"),
        "records": records,
        "skipped": skipped,
    }
    temporary = OUTPUT.with_suffix(".json.tmp")
    temporary.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(OUTPUT)


def extract(browser: BskBrowser) -> dict:
    payload = json.loads(browser._run("evaluate", EXTRACT, "--json", timeout=25))
    if not payload.get("ok") or not isinstance(payload.get("value"), dict):
        raise BrowserError("无法读取热门榜详情弹窗")
    return payload["value"]


def open_chart(browser: BskBrowser, chart_index: int) -> None:
    state = browser._run("observe", "--max-tokens", "12000", timeout=35)
    buttons = re.findall(r'@e(\d+) button "iconbofang"', state)
    target = chart_index * 10
    if len(buttons) < target + 1:
        raise BrowserError(f"找不到第 {chart_index+1} 个榜单的首条视频")
    browser._run("click", "--ref", f"@e{buttons[target]}", timeout=20)
    first = extract(browser)
    if first.get("position") != "1/50":
        raise BrowserError(f"榜单打开后分页不是 1/50：{first.get('position')}")


def main() -> None:
    browser = BskBrowser(print)
    if OUTPUT.exists():
        data = json.loads(OUTPUT.read_text(encoding="utf-8"))
        records = data["records"]
        skipped = data.get("skipped", [])
    else:
        records = []
        skipped = []
    try:
        browser.start()
        browser._run("navigate", URL, timeout=55)
        for chart_index, chart in enumerate(CHARTS):
            open_chart(browser, chart_index)
            for rank in range(1, 51):
                if any(row["chart"] == chart and row["rank"] == rank
                       for row in records + skipped):
                    if rank < 50:
                        browser._run("click", "--selector", NEXT, timeout=20)
                    continue
                expected = f"{rank}/50"
                for _ in range(15):
                    item = extract(browser)
                    if item.get("position") == expected:
                        break
                    time.sleep(0.2)
                else:
                    raise BrowserError(f"{chart} 第 {rank} 条未加载；最后读到 " +
                                       json.dumps(item, ensure_ascii=True)[:500])
                if not re.fullmatch(r"\d+", item.get("video_id", "")) or not (item.get("copy") or item.get("title")):
                    time.sleep(1)
                    item = extract(browser)
                if not re.fullmatch(r"\d+", item.get("video_id", "")) or not (item.get("copy") or item.get("title")):
                    skipped.append({"chart": chart, "rank": rank,
                                    "reason": "详情缺少作品 ID 或文案", "observed": item})
                    print(json.dumps({"chart": chart, "skipped": rank,
                                      "reason": "missing_id_or_copy"}, ensure_ascii=True), flush=True)
                else:
                    item["copy_source"] = "full_description" if item.get("copy") else "title_fallback"
                    item["copy"] = item.get("copy") or item["title"]
                    records.append({"chart": chart, "rank": rank, **item})
                save(records, skipped)
                if rank % 10 == 0:
                    print(json.dumps({"chart": chart, "processed": rank,
                                      "valid_total": len(records), "skipped_total": len(skipped)},
                                     ensure_ascii=True), flush=True)
                if rank < 50:
                    browser._run("click", "--selector", NEXT, timeout=20)
            if chart_index < len(CHARTS) - 1:
                browser._run("click", "--selector", CLOSE, timeout=20)
        print(json.dumps({"complete": True, "rows": len(records), "skipped": len(skipped),
                          "output": str(OUTPUT)},
                         ensure_ascii=True), flush=True)
    finally:
        browser.close()


if __name__ == "__main__":
    main()
