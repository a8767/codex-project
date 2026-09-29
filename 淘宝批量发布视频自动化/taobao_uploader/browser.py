from __future__ import annotations

import json
import re
import subprocess
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable

from .tasks import SCHEDULE_MAX_AHEAD, SCHEDULE_MIN_LEAD, TZ, Task


PUBLISH_URL = (
    "https://huodong.taobao.com/wow/z/guang/gg_publish/gg-video"
    "?ugc_scene=pc_newcreator_video&pageType=video&site=guangguang"
)
WORKSPACE_URL = "https://creator.guanghe.taobao.com/page/workspace/tb"


class BrowserError(RuntimeError):
    pass


class PauseForHuman(BrowserError):
    pass


class DuplicateVideoWarning(BrowserError):
    """The platform warned that the selected video may already be published."""


DUPLICATE_VIDEO_MARKERS = (
    "视频可能重复发布", "视频疑似重复", "视频内容重复", "视频重复发布",
    "视频重复上传", "该视频已发布", "视频已上传过",
)


class SubmissionUncertain(BrowserError):
    pass


@dataclass(frozen=True)
class VerifiedPublication:
    work_id: str
    message: str
    verified_at: datetime
    audit_status: str = ""


class BskBrowser:
    """Only the agent window is used. Commands never borrow the user's existing tabs."""

    def __init__(self, log: Callable[[str], None] = print):
        self.log = log
        self.session: str | None = None
        self.browser_id: str | None = None

    def _run(self, *args: str, timeout: int = 45) -> str:
        command = ["bsk", *args]
        if self.session and "--session" not in command and args[:2] not in {("session", "stop"), ("session", "list")}:
            command.extend(["--session", self.session])
        try:
            result = subprocess.run(command, capture_output=True, text=True, encoding="utf-8",
                                    errors="replace", timeout=timeout, check=False,
                                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        except FileNotFoundError as exc:
            raise BrowserError("找不到 bsk 命令。请先安装并连接 browser-skill 浏览器扩展。") from exc
        except subprocess.TimeoutExpired as exc:
            raise BrowserError(f"浏览器操作超时：{' '.join(args[:2])}") from exc
        output = (result.stdout + "\n" + result.stderr).strip()
        if result.returncode:
            raise BrowserError(output[-1200:] or f"bsk 退出码 {result.returncode}")
        return output

    def start(self) -> None:
        listing = self._run("browsers", timeout=20)
        chrome = re.findall(r"(?m)^([a-f0-9]{8})\s+chrome\s", listing, flags=re.I)
        if len(chrome) != 1:
            raise PauseForHuman("请保持一个 Chrome 实例和 bsk 扩展连接；当前未检测到唯一的 Chrome。")
        self.browser_id = chrome[0]
        output = self._run("session", "start", "--browser", self.browser_id, "--no-focus",
                           "--width", "1500", "--height", "900", timeout=35)
        found = re.search(r"(?m)^([a-z]{4})$", output)
        if not found:
            raise BrowserError("无法从 bsk 返回值读取会话 ID：" + output[-400:])
        self.session = found.group(1)

    def close(self) -> None:
        if not self.session:
            return
        session = self.session
        try:
            self._run("session", "stop", session, timeout=12)
            self.session = None
            return
        except BrowserError as exc:
            detail = str(exc)

        # A timed-out stop can keep running in the daemon. Do not issue a second
        # stop request: that produces "session stop is already in progress".
        self.log("浏览器关闭遇到确认窗口，正在请求你处理。")
        try:
            self._run(
                "request-help", "--title", "请确认关闭发布器",
                "--prompt", "浏览器正在等待“是否离开网站”确认。请点击“离开”完成关闭；"
                           "如果要保留草稿，请点击“取消”。处理后不要重试发布。",
                "--timeout", "25s", timeout=30,
            )
        except BrowserError as exc:
            detail = f"{detail}; request-help: {exc}"

        deadline = time.monotonic() + 15
        last_error = detail
        while time.monotonic() < deadline:
            try:
                listing = self._run("session", "list", timeout=5)
                if not re.search(rf"(?m)^{re.escape(session)}\s", listing):
                    self.log("本次自动化网页会话已关闭。")
                    self.session = None
                    return
            except BrowserError as exc:
                last_error = str(exc)
            time.sleep(1)
        self.log(
            f"自动化网页会话仍在关闭（{last_error}）。如看到“离开网站”确认，请点击“离开”；"
            "程序不会重复提交发布或重复发送关闭命令。"
        )
        self.session = None

    def observe(self) -> str:
        state = self._run("observe", "--max-tokens", "9000", timeout=35)
        if any(marker in state for marker in ("滑块验证", "拼图验证", "安全验证", "请完成验证", "验证码")):
            raise PauseForHuman("页面出现验证码，请在 Chrome 的任务窗口中人工完成后继续。")
        if "登录" in state and "发布视频" not in state and "发布作品" not in state:
            raise PauseForHuman("淘宝登录状态已失效，请在 Chrome 中登录后继续。")
        return state

    def duplicate_warning(self, seconds: int = 8) -> str:
        """Wait briefly for the publisher's visible duplicate-video warning."""
        markers = DUPLICATE_VIDEO_MARKERS
        until = time.monotonic() + max(0, seconds)
        while True:
            try:
                state = self.observe()
            except PauseForHuman:
                raise
            except BrowserError as exc:
                raise PauseForHuman(
                    "视频上传后无法检查平台重复提示；页面保持未提交，请人工检查后再继续。"
                ) from exc
            for marker in markers:
                if marker in state:
                    return marker
            if time.monotonic() >= until:
                return ""
            time.sleep(1)

    @staticmethod
    def ref(state: str, role: str, name: str, *, exact: bool = True) -> str:
        role = re.escape(role)
        name = re.escape(name)
        ending = r'"' if exact else r'[^\"]*"'
        match = re.search(rf'(?m)^\s*(@e\d+)\s+{role}\s+"{name}{ending}', state)
        if not match:
            raise BrowserError(f"页面上找不到 {role}「{name}」。请核对页面是否更新。")
        return match.group(1)

    def click(self, role: str, name: str, *, exact: bool = True) -> str:
        state = self.observe()
        ref = self.ref(state, role, name, exact=exact)
        self._run("click", ref)
        try:
            return self.observe()
        except BrowserError as exc:
            # Publishing can replace the document between a successful click and
            # its observation. A fresh read is safe; clicking again is not.
            if "Document changed during observation" not in str(exc):
                raise
            return self.observe()

    def fill(self, role: str, name: str, value: str, *, exact: bool = True) -> str:
        state = self.observe()
        ref = self.ref(state, role, name, exact=exact)
        self._run("fill", ref, "--value", value)
        return self.observe()

    def navigate(self) -> str:
        self._run("navigate", PUBLISH_URL, timeout=55)
        try:
            return self._wait_for(
                lambda s: 'RootWebArea "发布器"' in s and "发布视频" in s,
                "发布器加载", seconds=30,
            )
        except PauseForHuman:
            raise
        except BrowserError as exc:
            raise PauseForHuman("发布器未加载，检查当前店铺、登录状态和页面访问权限。") from exc

    def _wait_for(self, predicate: Callable[[str], bool], description: str, seconds: int = 120) -> str:
        until = time.monotonic() + seconds
        while time.monotonic() < until:
            state = self.observe()
            if predicate(state):
                return state
            time.sleep(2)
        raise BrowserError(f"等待{description}超时")

    def _upload(self, path: Path, kind: str) -> None:
        # Open the publisher URL directly so CSS selectors resolve in its document.
        candidates = {
            "video": ['span.upload-button', 'input[type="file"][accept*="video"]',
                      'input[type="file"][accept*="mp4"]'],
            "cover": ['input[type="file"][accept*="image"]'],
        }[kind]
        errors = []
        for selector in candidates:
            try:
                self._run("upload", "--selector", selector, "--file", str(path), "--timeout", "2m", timeout=140)
                return
            except BrowserError as exc:
                detail = str(exc)
                if "not allowed" in detail.lower() and "file url" in detail.lower():
                    self.log("Chrome 扩展没有文件 URL 访问权限，请在任务窗口中手动选择文件。")
                    result = self._run(
                        "request-help", "--title", "请手动选择文件",
                        "--prompt", f"请在当前发布器手动选择这个{kind}文件：{path}。等待上传完成后点击继续；不要点击发布。",
                        "--target", selector, "--timeout", "5m", timeout=310,
                    )
                    if "outcome=continued" not in result and "outcome=completed" not in result:
                        raise PauseForHuman("手动上传未完成，请在任务窗口中选择文件后继续。")
                    return
                if "effect_state=unknown" in detail or "effect_state=committed" in detail:
                    raise BrowserError(f"上传动作结果不明，已停止以免重复上传：{detail}") from exc
                safe_to_retry = (
                    "file_input_not_activated" in detail and "effect_state=none" in detail
                ) or any(marker in detail.lower() for marker in ("not found", "no element", "no node", "did not match"))
                if not safe_to_retry:
                    raise
                errors.append(detail)
        raise BrowserError(f"无法上传{kind}文件：" + " | ".join(errors)[-900:])

    def _video(self, task: Task) -> None:
        self._upload(task.video, "video")
        state = self._wait_for(
            lambda s: any(marker in s for marker in DUPLICATE_VIDEO_MARKERS)
            or ("等待视频上传..." not in s and any(
                x in s for x in ("更换视频", "视频封面", "重新上传")
            )),
            "视频上传处理", seconds=600,
        )
        if "上传失败" in state or "视频处理失败" in state:
            raise BrowserError("平台提示视频上传或处理失败")
        duplicate_warning = self.duplicate_warning()
        if duplicate_warning:
            raise DuplicateVideoWarning(f"平台提示：{duplicate_warning}")
        if task.cover:
            self._upload(task.cover, "cover")
            self._wait_for(lambda s: "视频封面" in s and "上传失败" not in s, "封面上传")
        else:
            cover_section = state.split('StaticText "视频封面"', 1)[-1].split('StaticText "视频描述"', 1)[0]
            if any(text in cover_section for text in ("封面生成失败", "请上传封面", "暂无封面")):
                raise PauseForHuman("平台没有生成封面，请补充封面图片后重新导入任务。")

    def _body(self, text: str) -> None:
        if not text:
            return
        state = self.observe()
        # The rich-text editor is exposed as the unnamed textbox in the publisher.
        matches = re.findall(r'(?m)^\s*(@e\d+)\s+textbox(?:\s+\[empty\])?\s*$', state)
        if len(matches) != 1:
            raise BrowserError("无法唯一定位视频正文编辑框")
        self._run("fill", matches[0], "--value", text)
        state = self.observe()
        if f'StaticText "{text}"' not in state and f'textbox "{text}"' not in state:
            raise BrowserError("视频正文填写后未在页面显示指定内容")

    def _products(self, ids: tuple[str, ...]) -> None:
        if not ids:
            return
        self.click("button", "添加商品")
        for product_id in ids:
            self.fill("searchbox", "搜索", product_id)
            state = self.click("button", "搜索")
            if 'dialog "关联商品"' not in state or "没有找到" in state:
                raise BrowserError(f"商品 {product_id} 没有搜索结果")
            # Search-by-ID must yield exactly one selectable product. Ambiguous hits are rejected.
            checkboxes = re.findall(r'(?m)^\s*(@e\d+)\s+checkbox\s+""[^\n]*\[unchecked\]', state)
            if len(checkboxes) != 1:
                raise BrowserError(f"商品 {product_id} 搜索到 {len(checkboxes)} 个候选，无法安全关联")
            self._run("click", checkboxes[0])
            self.observe()
        self.click("button", "确定")

    def _tags(self, task: Task) -> None:
        for tag in task.brand_tags:
            state = self.click("button", "品牌标签")
            inputs = re.findall(r'(?m)^\s*(@e\d+)\s+(?:searchbox|textbox)\s+"(?:搜索|输入标签|请输入标签)[^\"]*"', state)
            if len(inputs) != 1:
                raise BrowserError("无法定位品牌标签输入框")
            self._run("fill", inputs[0], "--value", tag)
            state = self.observe()
            exact = re.findall(rf'(?m)^\s*(@e\d+)\s+(?:option|button|checkbox)\s+"{re.escape(tag)}"', state)
            if len(exact) != 1:
                raise BrowserError(f"品牌标签「{tag}」无唯一匹配")
            self._run("click", exact[0])
        for raw_tag in task.content_tags:
            tag = raw_tag.lstrip("#＃").strip()
            if not tag:
                continue
            state = self.click("button", "内容标签")
            # The current rich-text editor searches through its focused, unnamed
            # textarea. The title input has a placeholder and must not be filled.
            inputs = re.findall(r'(?m)^\s*(@e\d+)\s+textbox(?:\s+"[^\"]*")?\s+' 
                                r'\[(?:empty|filled)\](?![^\n]*placeholder)', state)
            if len(inputs) != 1:
                raise BrowserError("无法唯一定位内容标签搜索输入框")
            self._run("fill", inputs[0], "--value", tag)
            # Search results are debounced by the publisher and may arrive after
            # the first observation, especially while video processing finishes.
            until = time.monotonic() + 12
            while True:
                state = self.observe()
                exact = re.findall(
                    rf'(?m)^\s*(@e\d+)\s+(?:button|option)\s+"#?\s*{re.escape(tag)}"(?:\s|$)',
                    state,
                )
                if len(exact) == 1 or time.monotonic() >= until:
                    break
                time.sleep(1)
            if len(exact) != 1:
                candidates = [line.strip() for line in state.splitlines()
                              if re.search(r'(?m)^\s*@e\d+\s+(?:button|option)\s+', line)]
                detail = "；当前候选：" + " | ".join(candidates[:8]) if candidates else "；页面未返回按钮候选"
                raise BrowserError(
                    f"平台内容标签「{tag}」精确匹配 {len(exact)} 条；请核对搜索结果{detail}"
                )
            self._run("click", exact[0])
            self.observe()
            # Plain '#tag' text is not a platform tag. Confirm a real editor token.
            result = json.loads(self._run(
                "evaluate",
                "Array.from(document.querySelectorAll('.rich-text-content [data-cangjie-void]'))"
                ".map(e=>e.textContent.trim())",
                "--json"))
            if not result.get("ok") or tag not in result.get("value", []):
                raise BrowserError(f"内容标签「{tag}」未插入视频描述")
        if task.topic:
            state = self.click("button", "点击添加话题")
            if 'dialog "话题选择"' in state:
                self.fill("textbox", "输入关键词搜索", task.topic)
                state = self.click("button", "搜索")
                if "没有找到相关话题" in state:
                    raise BrowserError(f"平台没有可选话题「{task.topic}」，请填写完整的平台话题名称")
                # The current topic cards are clickable divs rather than options.
                # Match their complete platform title so a keyword hit cannot be selected by accident.
                selector = ('[data-autolog-container="topic-item-card"]'
                            f'[data-autolog="key=topic_item_card&text=可选话题:{task.topic}"]')
                self._run("click", "--selector", selector)
                state = self.observe()
                if ("当前选中话题：" not in state or
                        f'StaticText "{task.topic}"' not in state):
                    raise BrowserError(f"无法确认已选中话题「{task.topic}」")
                self.click("button", "确认提交")
                self._wait_for(lambda s: 'dialog "话题选择"' not in s and
                               f'button "{task.topic}"' in s,
                               "话题回填", seconds=15)
            else:
                inputs = re.findall(r'(?m)^\s*(@e\d+)\s+(?:searchbox|textbox)\s+"[^\"]*"[^\n]*(?:placeholder="[^\"]*话题[^\"]*")', state)
                if len(inputs) != 1:
                    raise BrowserError("无法定位话题搜索框")
                self._run("fill", inputs[0], "--value", task.topic)
                state = self.observe()
                exact = re.findall(rf'(?m)^\s*(@e\d+)\s+(?:option|button)\s+"{re.escape(task.topic)}"', state)
                if len(exact) != 1:
                    raise BrowserError("话题搜索结果不唯一")
                self._run("click", exact[0])

    def _schedule(self, task: Task) -> None:
        if task.publish_at is None:
            return
        now = datetime.now(TZ)
        if not now + SCHEDULE_MIN_LEAD <= task.publish_at <= now + SCHEDULE_MAX_AHEAD:
            raise BrowserError("定时发布时间须在当前北京时间之后 2 小时至 30 天内，请修改 Excel 并重新导入")
        state = self.observe()
        schedule_part = state.split('StaticText "定时发布"', 1)
        radios = re.findall(r'(?m)^\s*(@e\d+)\s+radio\s+\[unchecked\][^\n]*$', schedule_part[0][-240:])
        if not radios or len(schedule_part) != 2:
            raise BrowserError("无法定位定时发布开关")
        self._run("click", radios[-1])
        state = self.observe()
        if "[disabled]" in next((line for line in state.splitlines() if 'combobox "请选择日期和时间' in line), ""):
            raise BrowserError("定时日期选择器仍处于禁用状态")
        target = task.publish_at.strftime("%Y/%m/%d %H:%M")
        date = task.publish_at.strftime("%Y/%m/%d")
        # The publisher's outer combobox is readonly. Operate the calendar and
        # time picker in one bsk evaluation so a popup blur cannot discard it.
        script = """(() => {
          const picker = document.querySelector('#date-picker');
          if (!picker) throw Error('date picker missing');
          let body = document.querySelector('.next-date-picker-body[aria-hidden="false"]');
          if (!body) picker.querySelector('.next-date-picker-trigger').click();
          body = document.querySelector('.next-date-picker-body[aria-hidden="false"]');
          if (!body) throw Error('date picker did not open');
          const button = name => Array.from(body.querySelectorAll('button'))
              .find(x => x.textContent.trim() === name);
          button('选择日期')?.click();
          const panel = body.querySelector('.next-calendar-panel');
          if (!panel) throw Error('calendar panel missing');
          const months = ['一月','二月','三月','四月','五月','六月',
                          '七月','八月','九月','十月','十一月','十二月'];
          const names = Array.from(panel.querySelectorAll(
              '.next-calendar-panel-header-full button')).map(x => x.textContent.trim());
          const currentMonth = months.indexOf(names[0]) + 1;
          const currentYear = Number(names[1]);
          if (!currentMonth || !Number.isInteger(currentYear)) throw Error('calendar header unknown');
          const delta = (__YEAR__ - currentYear) * 12 + __MONTH__ - currentMonth;
          if (Math.abs(delta) > 3) throw Error('calendar month out of expected range');
          const direction = delta < 0 ? '上个月' : '下个月';
          for (let i = 0; i < Math.abs(delta); i++) {
            const arrow = panel.querySelector(`button[title="${direction}"]`);
            if (!arrow || arrow.disabled) throw Error('calendar month navigation unavailable');
            arrow.click();
          }
          const cell = panel.querySelector(`td[title="${__DATE__}"]`);
          if (!cell || cell.getAttribute('aria-disabled') !== 'false')
            throw Error('requested calendar day is unavailable');
          cell.click();
          const timeButton = button('选择时间');
          if (!timeButton) throw Error('time panel unavailable');
          timeButton.click();
          const hour = body.querySelector(`.next-time-picker-menu-hour [role="option"][title="${__HOUR__}"]`);
          if (!hour || hour.getAttribute('aria-disabled') === 'true')
            throw Error('requested hour unavailable');
          hour.click();
          const minute = body.querySelector(`.next-time-picker-menu-minute [role="option"][title="${__MINUTE__}"]`);
          if (!minute || minute.getAttribute('aria-disabled') === 'true')
            throw Error('requested minute unavailable');
          minute.click();
          if (body.querySelector('input[placeholder="YYYY/MM/DD"]')?.value !== __DATE__ ||
              body.querySelector('input[placeholder="HH:mm"]')?.value !== __CLOCK__)
            throw Error('calendar selection did not stick');
          const confirm = button('确定');
          if (!confirm) throw Error('calendar confirmation missing');
          confirm.click();
          const actual = picker.querySelector('input[role="combobox"]')?.value;
          if (actual !== __TARGET__) throw Error(`scheduled value mismatch: ${actual}`);
          return actual;
        })()"""
        for marker, value in {
            "__YEAR__": str(task.publish_at.year),
            "__MONTH__": str(task.publish_at.month),
            "__DATE__": json.dumps(date),
            "__CLOCK__": json.dumps(task.publish_at.strftime("%H:%M")),
            "__HOUR__": str(task.publish_at.hour),
            "__MINUTE__": str(task.publish_at.minute),
            "__TARGET__": json.dumps(target),
        }.items():
            script = script.replace(marker, value)
        try:
            result = json.loads(self._run("evaluate", script, "--json", timeout=45))
        except (ValueError, TypeError) as exc:
            raise BrowserError("无法解析日期选择器的返回结果") from exc
        if not result.get("ok") or result.get("value") != target:
            detail = result.get("error", {}).get("text", "返回时间与目标不一致")
            raise BrowserError(f"设置定时发布时间失败：{detail}")
        state = self.observe()
        self._verify_schedule(task, state)

    @staticmethod
    def _verify_schedule(task: Task, state: str) -> None:
        if task.publish_at is None:
            return
        schedule_marker = 'StaticText "定时发布"'
        if schedule_marker not in state:
            raise BrowserError("页面未显示定时发布选项")
        before = state.split(schedule_marker, 1)[0]
        if not re.search(r'(?m)^\s*@e\d+\s+radio\s+\[checked\][^\n]*$', before[-240:]):
            raise BrowserError("定时发布开关未选中")
        target = task.publish_at.strftime("%Y/%m/%d %H:%M")
        field = re.search(r'(?m)^\s*@e\d+\s+combobox\s+"请选择日期和时间[^\n]*$', state)
        if not field or f'="{target}"' not in field.group(0):
            raise BrowserError(f"定时发布时间未设置为 {target}")

    def prepare(self, task: Task) -> None:
        self.navigate()
        self.log(f"{task.task_id}：上传视频")
        self._video(task)
        self.fill("textbox", "加个标题让内容更吸引人", task.title)
        self._body(task.body)
        self._products(task.product_ids)
        self._tags(task)
        self._schedule(task)
        self.click("radio", task.declaration)
        state = self.observe()
        declaration = re.search(
            rf'(?m)^\s*@e\d+\s+radio\s+"{re.escape(task.declaration)}"[^\n]*\[checked\]',
            state,
        )
        if task.title not in state or declaration is None:
            raise BrowserError("提交前标题或创作者声明校验失败")

    @staticmethod
    def recent_work_ids(state: str, title: str, submitted_at: datetime) -> set[str]:
        found: set[str] = set()
        pattern = re.compile(
            r'(?m)^\s*row "(?:置顶 )?\d{2}:\d{2}(?::\d{2})? (?P<title>.*?) '
            r'ID:\s*(?P<id>\d+)\s*\|\s*(?P<created>\d{4}-\d{2}-\d{2} '
            r'\d{2}:\d{2}:\d{2})'
        )
        for match in pattern.finditer(state):
            if match.group("title") != title:
                continue
            created = datetime.strptime(match.group("created"), "%Y-%m-%d %H:%M:%S").replace(tzinfo=TZ)
            if submitted_at - timedelta(minutes=3) <= created <= datetime.now(TZ) + timedelta(minutes=2):
                found.add(match.group("id"))
        return found

    @staticmethod
    def scheduled_work_ids(state: str, title: str, publish_at: datetime) -> set[str]:
        """Find an exact scheduled title/time pair in the workspace results."""
        scheduled_label = f"定时发布 {publish_at:%Y-%m-%d %H:%M}"
        found: set[str] = set()
        for row in re.findall(r'(?m)^\s*row "[^\n]*', state):
            if f"{title} ID:" not in row or scheduled_label not in row:
                continue
            match = re.search(r"ID:\s*(\d+)", row)
            if match:
                found.add(match.group(1))
        return found

    def _verify_workspace_record(self, task: Task, submitted_at: datetime) -> VerifiedPublication:
        try:
            self._run("navigate", WORKSPACE_URL, timeout=55)
            state = self.observe()
            if 'RootWebArea "作品管理' not in state:
                raise BrowserError("作品管理页未加载")
            self.fill("textbox", "请输入作品ID或关键字", task.title)
            self.click("button", "搜索")
            until = time.monotonic() + 30
            while time.monotonic() < until:
                state = self.observe()
                matches = (self.scheduled_work_ids(state, task.title, task.publish_at)
                           if task.publish_at else
                           self.recent_work_ids(state, task.title, submitted_at))
                if len(matches) == 1:
                    work_id = next(iter(matches))
                    message = f"后台已找到作品 ID {work_id}"
                    if task.publish_at:
                        message += f"，定时发布 {task.publish_at:%Y-%m-%d %H:%M}"
                    audit_status = ""
                    for row in re.findall(r'(?m)^\s*row "[^\n]*', state):
                        if f"{task.title} ID:" in row and f"ID: {work_id}" in row:
                            audit_status = next((label for label in
                                ("公域审核通过", "公域审核中", "公域审核不通过", "作品违规")
                                if label in row), "")
                            if audit_status:
                                break
                    return VerifiedPublication(work_id, message, datetime.now(TZ), audit_status)
                if len(matches) > 1:
                    raise SubmissionUncertain("后台存在多条同名新作品，请人工核查")
                time.sleep(2)
        except BrowserError as exc:
            raise SubmissionUncertain(f"提交后后台核验失败：{exc}") from exc
        raise SubmissionUncertain("提交后 30 秒内未在作品管理中找到新作品，请人工核查")

    def submit_and_verify(self, task: Task) -> VerifiedPublication:
        submitted_at = datetime.now(TZ)
        before = self.observe()
        if task.publish_at:
            self._verify_schedule(task, before)
            if 'button "定时发布"' not in before:
                raise BrowserError("定时发布按钮未出现；已阻止误点立即发布")
            button = "定时发布"
        else:
            button = "立即发布"
        try:
            state = self.click("button", button)
        except BrowserError as exc:
            raise SubmissionUncertain(f"点击提交时状态不明：{exc}") from exc
        until = time.monotonic() + 30
        while time.monotonic() < until:
            if any(token in state for token in ("发布失败", "提交失败")):
                raise SubmissionUncertain("平台显示提交失败；请核查后台是否产生记录")
            if any(token in state for token in ("发布成功", "提交成功", "定时发布成功", "发布完成")) or "等待视频上传..." in state:
                break
            time.sleep(2)
            try:
                state = self.observe()
            except BrowserError as exc:
                raise SubmissionUncertain(f"提交后无法读取结果：{exc}") from exc
        return self._verify_workspace_record(task, submitted_at)
