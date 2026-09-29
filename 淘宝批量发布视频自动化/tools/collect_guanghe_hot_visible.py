"""Collect visible Guanghe hot-chart copy, with resumable progress."""

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
OUTPUT = ROOT / "outputs" / "video_reference_db" / "光合热门榜_最近7天_文案话题.json"
PREVIOUS = ROOT / "outputs" / "video_reference_db" / "光合热门榜_最近1-30天_采集.json"
RANGES = ("最近7天",)
CHARTS = ("热度榜-查看人数", "引流榜-商品点击人数", "成交榜-种草成交金额")
CLOSE = ".g-pc-gh-content-dialog-extra > .g-pc-gh-content-dialog-btn:first-child"
EXTRACT = r"""(() => {
  const d = document.querySelector('[role="dialog"]');
  if (!d) return null;
  const text = (selector) => d.querySelector(selector)?.textContent?.trim() || '';
  const nums = Array.from(d.querySelectorAll('.g-pc-gh-content-dialog-user-numwrap'));
  const number = (label) => nums.find(e => e.textContent?.startsWith(label))?.querySelector('span')?.textContent?.trim() || '';
  return {
    video_id: number('作品ID'),
    creator_id: number('账号ID'),
    creator: text('.g-pc-gh-content-dialog-user-nick'),
    title: text('.g-pc-gh-content-dialog-content-title'),
    copy: text('.g-pc-gh-content-dialog-content-summary'),
    topics: Array.from(d.querySelectorAll('.g-pc-gh-content-dialog-content-topicsTagText'))
      .map(e => e.textContent?.trim()).filter(Boolean)
  };
})()"""
LIST = r"""(() => {
  const wrap = document.querySelector('[class*="listWrapper"]');
  if (!wrap) return null;
  return Array.from(wrap.children).map(column =>
    Array.from(column.querySelectorAll(':scope > .view.guang-feed > .view > [class*="cellWrapper"]'))
      .map((card, index) => ({
        rank: index + 1,
        list_title: card.querySelector('[class*="contentTitle"]')?.textContent?.trim() || '',
        list_creator: card.querySelector('[class*="nickName"]')?.textContent?.trim() || ''
      }))
  );
})()"""


def evaluate(browser: BskBrowser, script: str):
    payload = json.loads(browser._run("evaluate", script, "--json", timeout=25))
    if not payload.get("ok"):
        raise BrowserError("榜单页面读取失败：" + str(payload.get("error", ""))[:250])
    return payload.get("value")


def save(records: list[dict], skipped: list[dict]) -> None:
    data = {
        "source_url": URL,
        "industry": "自行车/骑行装备/零配件",
        "ranges": list(RANGES),
        "charts": list(CHARTS),
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


def change_range(browser: BskBrowser, name: str) -> None:
    dropdown = None
    for _ in range(12):
        state = browser._run("observe", "--max-tokens", "9500", timeout=35)
        dropdown = re.search(r'(?m)^\s*(@e\d+) combobox "筛选范围', state)
        if dropdown:
            break
        time.sleep(0.2)
    if dropdown is None:
        raise BrowserError("找不到榜单时间筛选")
    browser._run("click", "--ref", dropdown.group(1), timeout=20)
    state = browser._run("observe", "--max-tokens", "9500", timeout=35)
    option = re.search(rf'(?m)^\s*(@e\d+) option "{re.escape(name)}"', state)
    if not option:
        raise BrowserError(f"找不到时间范围 {name}")
    browser._run("click", "--ref", option.group(1), timeout=20)
    for _ in range(20):
        selected = evaluate(browser, "document.querySelector('[class*=listWrapper]')?.parentElement?.innerText.slice(0,300)")
        if selected and name in selected:
            break
        time.sleep(0.2)


def close_detail(browser: BskBrowser) -> None:
    for attempt in range(3):
        if not evaluate(browser, "!!document.querySelector('.g-pc-gh-content-dialog-total')"):
            return
        if attempt < 2:
            browser._run("click", "--selector", CLOSE, timeout=20)
        else:
            browser._run("press", "Escape", timeout=20)
        for _ in range(20):
            if not evaluate(browser, "!!document.querySelector('.g-pc-gh-content-dialog-total')"):
                return
            time.sleep(0.2)
    raise BrowserError("视频详情无法关闭；请人工处理测试窗口")


def main() -> None:
    if OUTPUT.exists():
        data = json.loads(OUTPUT.read_text(encoding="utf-8"))
        records, skipped = data["records"], data.get("skipped", [])
    elif PREVIOUS.exists():
        old = json.loads(PREVIOUS.read_text(encoding="utf-8"))
        records = [row for row in old["records"] if row.get("date_range") == "最近7天"]
        skipped = [row for row in old.get("skipped", []) if row.get("date_range") == "最近7天"]
    else:
        records, skipped = [], []
    skipped = [row for row in skipped
               if row.get("reason") not in {"视频详情缺少作品 ID", "视频详情没有可读文案"}]
    browser = BskBrowser(lambda message: print(message, flush=True))
    try:
        browser.start()
        browser._run("navigate", URL, timeout=55)
        for range_index, date_range in enumerate(RANGES):
            if range_index:
                change_range(browser, date_range)
            lists = evaluate(browser, LIST)
            for _ in range(8):
                if lists and len(lists) == 3 and all(len(rows) >= 50 for rows in lists):
                    break
                browser._run("wheel", "--delta-y", "1200", timeout=20)
                lists = evaluate(browser, LIST)
            if not lists or len(lists) != 3 or any(len(rows) != 50 for rows in lists):
                raise BrowserError(f"{date_range} 榜单未加载为三栏各 50 条：" +
                                   str([len(rows) for rows in lists] if lists else None))
            for chart_index, chart in enumerate(CHARTS):
                consecutive_misses = 0
                for card in lists[chart_index]:
                    rank = card["rank"]
                    if any(row.get("date_range") == date_range and row["chart"] == chart
                           and row["rank"] == rank for row in records + skipped):
                        continue
                    selector = (f'[class*="listWrapper"] > [class*="colWrapper"]:nth-child({chart_index+1})'
                                f' > .view.guang-feed > .view > [class*="cellWrapper"]:nth-child({rank})'
                                ' [class*="sourceWrapper"] svg.bofang')
                    try:
                        count = evaluate(browser, f"document.querySelectorAll({json.dumps(selector)}).length")
                        if count != 1:
                            raise BrowserError(f"卡片定位数量是 {count}，不是 1")
                        browser._run("scroll-to", "--selector", selector, timeout=20)
                        browser._run("click", "--selector", selector, timeout=20)
                        item = None
                        for _ in range(30):
                            item = evaluate(browser, EXTRACT)
                            if item and (item.get("copy") or item.get("title")):
                                break
                            time.sleep(0.2)
                        if not item or not (item.get("copy") or item.get("title")):
                            consecutive_misses += 1
                            skipped.append({"date_range": date_range, "chart": chart, **card,
                                            "reason": "视频详情没有可读文案"})
                        else:
                            displayed = card["list_title"]
                            detail = item.get("title") or ""
                            if detail and not (detail.startswith(displayed) or displayed.startswith(detail)):
                                raise BrowserError("详情标题与点击的榜单卡片不一致")
                            item["copy_source"] = "full_description" if item.get("copy") else "title_fallback"
                            item["copy"] = item.get("copy") or item.get("title") or card["list_title"]
                            records.append({"date_range": date_range, "chart": chart, **card, **item})
                            consecutive_misses = 0
                    except BrowserError as exc:
                        consecutive_misses += 1
                        skipped.append({"date_range": date_range, "chart": chart, **card,
                                        "reason": str(exc)[:250]})
                    finally:
                        close_detail(browser)
                    save(records, skipped)
                    if consecutive_misses >= 3:
                        raise BrowserError(f"{chart} 连续三条详情为空；可能出现验证或页面异常，已暂停采集")
                    # The site repeatedly challenged faster modal navigation.
                    # Keep future collection deliberately slow and resumable.
                    time.sleep(6)
                    if rank % 10 == 0:
                        time.sleep(20)
                print(json.dumps({"range": date_range, "chart": chart,
                                  "valid": len(records), "skipped": len(skipped)},
                                 ensure_ascii=True), flush=True)
        print(json.dumps({"complete": True, "rows": len(records),
                          "skipped": len(skipped), "output": str(OUTPUT)},
                         ensure_ascii=True), flush=True)
    finally:
        browser.close()


if __name__ == "__main__":
    main()
