from pathlib import Path
from datetime import datetime, timedelta
import unittest
from unittest.mock import patch

from taobao_uploader.browser import BskBrowser, BrowserError, DuplicateVideoWarning, PUBLISH_URL
from taobao_uploader.tasks import TZ, Task


class PublisherNavigationTests(unittest.TestCase):
    def test_close_waits_for_in_progress_stop_without_sending_a_second_stop(self):
        browser = BskBrowser()
        browser.session = "mxbq"
        calls = []
        listings = iter(("mxbq bf65b17e 557535231", "SESSION  BROWSER   AGENT WINDOW"))

        def run(*args, **_kwargs):
            calls.append(args)
            if args[:2] == ("session", "stop"):
                raise BrowserError("operation timed out; session stop is already in progress")
            if args[0] == "request-help":
                return "outcome=continued"
            if args[:2] == ("session", "list"):
                return next(listings)
            raise AssertionError(args)

        messages = []
        browser._run = run
        browser.log = messages.append
        with patch("taobao_uploader.browser.time.sleep", lambda _seconds: None):
            browser.close()
        self.assertEqual(sum(args[:2] == ("session", "stop") for args in calls), 1)
        self.assertIsNone(browser.session)
        self.assertIn("本次自动化网页会话已关闭。", messages)

    def test_prepare_accepts_checked_declaration_with_context_metadata(self):
        browser = BskBrowser()
        task = Task(2, "checked", Path("video.mp4"), "标题", "正文", (), None,
                    (), (), "", None, "内容含营销信息")
        browser.navigate = lambda: None
        browser._video = lambda _task: None
        browser.fill = lambda *_args: None
        browser._body = lambda _body: None
        browser._products = lambda _products: None
        browser._tags = lambda _task: None
        browser._schedule = lambda _task: None
        browser.duplicate_warning = lambda: ""
        browser.click = lambda *_args: None
        browser.observe = lambda: ('@e1 textbox "加个标题让内容更吸引人" ="标题"\n'
                                   '@e2 radio "内容含营销信息" [ctx: iframe] [checked] ="on"')
        browser.prepare(task)

    def test_click_reobserves_once_when_page_changes_without_reclicking(self):
        browser = BskBrowser()
        observations = iter((
            '@e1 button "定时发布"',
            BrowserError("Document changed during observation"),
            'RootWebArea "发布器"',
        ))
        def observe():
            value = next(observations)
            if isinstance(value, BrowserError):
                raise value
            return value

        browser.observe = observe
        clicks = []
        browser._run = lambda *args, **_kwargs: clicks.append(args)
        self.assertEqual(browser.click("button", "定时发布"), 'RootWebArea "发布器"')
        self.assertEqual(clicks, [("click", "@e1")])

    def test_publisher_opens_directly_so_upload_selector_can_resolve(self):
        browser = BskBrowser()
        commands = []

        def run(*args, **_kwargs):
            commands.append(args)
            if args[0] == "observe":
                return 'RootWebArea "发布器"\n  StaticText "发布视频"'
            return ""

        browser._run = run
        browser.navigate()
        self.assertEqual(commands[0], ("navigate", PUBLISH_URL))
        self.assertTrue(PUBLISH_URL.startswith("https://huodong.taobao.com/"))

    def test_extension_file_permission_uses_manual_picker(self):
        browser = BskBrowser(lambda _: None)
        commands = []

        def run(*args, **_kwargs):
            commands.append(args)
            if args[0] == "upload":
                raise BrowserError('Not allowed. Enable "Allow access to file URLs" for the extension.')
            if args[0] == "request-help":
                return "outcome=continued"
            self.fail(f"unexpected command: {args}")

        browser._run = run
        browser._upload(Path("C:/video.mp4"), "video")
        self.assertEqual([command[0] for command in commands], ["upload", "request-help"])
        self.assertTrue(any("不要点击发布" in part for part in commands[1]))

    def test_session_cleanup_polls_after_timeout_without_reissuing_stop(self):
        messages = []
        browser = BskBrowser(messages.append)
        browser.session = "abcd"
        commands = []
        list_calls = 0

        def run(*args, **_kwargs):
            nonlocal list_calls
            commands.append(args)
            if len(commands) == 1:
                raise BrowserError("浏览器操作超时")
            if args[0] == "request-help":
                return "outcome=continued"
            list_calls += 1
            return "abcd bf65b17e 557535231" if list_calls == 1 else "(no active sessions)"

        browser._run = run
        with patch("taobao_uploader.browser.time.sleep", lambda _seconds: None):
            browser.close()
        self.assertEqual(commands, [("session", "stop", "abcd"),
                                    ("request-help", "--title", "请确认关闭发布器",
                                     "--prompt", "浏览器正在等待“是否离开网站”确认。请点击“离开”完成关闭；"
                                                "如果要保留草稿，请点击“取消”。处理后不要重试发布。",
                                     "--timeout", "25s"),
                                    ("session", "list"), ("session", "list")])
        self.assertIsNone(browser.session)
        self.assertIn("确认窗口", messages[0])

    def test_session_cleanup_accepts_late_completion(self):
        messages = []
        browser = BskBrowser(messages.append)
        browser.session = "abcd"
        commands = []

        def run(*args, **_kwargs):
            commands.append(args)
            if args[:2] == ("session", "stop"):
                raise BrowserError("session stop is already in progress")
            if args[0] == "request-help":
                return "outcome=continued"
            return "(no active sessions)"

        browser._run = run
        browser.close()
        self.assertEqual(commands[-1], ("session", "list"))
        self.assertEqual(sum(args[:2] == ("session", "stop") for args in commands), 1)
        self.assertTrue(any(args[0] == "request-help" for args in commands))
        self.assertFalse(any("关闭浏览器会话失败" in message for message in messages))

    def test_workspace_match_requires_exact_recent_title(self):
        now = datetime.now(TZ).replace(microsecond=0)
        recent = now.strftime("%Y-%m-%d %H:%M:%S")
        old = (now - timedelta(hours=1)).strftime("%Y-%m-%d %H:%M:%S")
        state = (
            f'row "00:14 测试视频 ID: 2264901856221090 | {recent} 0 0 0 0 公域审核通过"\n'
            f'row "00:14 测试视频合集 ID: 111 | {recent} 0 0 0 0 公域审核通过"\n'
            f'row "00:14 测试视频 ID: 222 | {old} 0 0 0 0 公域审核通过"'
        )
        self.assertEqual(BskBrowser.recent_work_ids(state, "测试视频", now),
                         {"2264901856221090"})

    def test_workspace_match_requires_exact_scheduled_title_and_time(self):
        at = datetime(2026, 9, 28, 10, 0, tzinfo=TZ)
        state = (
            'row "00:25 GT-05山地车｜骑行日常（定时测试） ID: 2312389332965071 |  '
            '定时发布 2026-09-28 10:00 修改 0 0 0 0"\n'
            'row "00:25 GT-05山地车｜骑行日常（定时测试）续集 ID: 222 |  '
            '定时发布 2026-09-28 10:00"\n'
            'row "00:25 GT-05山地车｜骑行日常（定时测试） ID: 333 |  '
            '定时发布 2026-09-28 11:00"'
        )
        self.assertEqual(BskBrowser.scheduled_work_ids(
            state, "GT-05山地车｜骑行日常（定时测试）", at), {"2312389332965071"})

    def test_scheduled_submission_never_falls_back_to_immediate(self):
        browser = BskBrowser()
        at = (datetime.now(TZ) + timedelta(hours=4)).replace(second=0, microsecond=0)
        task = Task(2, "scheduled", Path("video.mp4"), "标题", "", (), None,
                    (), (), "", at, "内容无需标注")
        browser.observe = lambda: (
            '@e1 radio [checked] ="on"\n'
            'StaticText "定时发布"\n'
            f'@e2 combobox "请选择日期和时间 [has-submenu]" ="{at:%Y/%m/%d %H:%M}"\n'
            '@e3 button "立即发布"'
        )
        clicked = []
        browser.click = lambda *args, **kwargs: clicked.append(args)
        with self.assertRaisesRegex(BrowserError, "已阻止误点立即发布"):
            browser.submit_and_verify(task)
        self.assertEqual(clicked, [])

    def test_scheduled_time_requires_exact_minute(self):
        at = (datetime.now(TZ) + timedelta(hours=4)).replace(second=0, microsecond=0)
        wrong_minute = (at.minute + 1) % 60
        task = Task(2, "scheduled", Path("video.mp4"), "标题", "", (), None,
                    (), (), "", at, "内容无需标注")
        state = (
            '@e1 radio [checked] ="on"\nStaticText "定时发布"\n'
            f'@e2 combobox "请选择日期和时间 [has-submenu]" ="{at:%Y/%m/%d %H}:{wrong_minute:02d}"'
        )
        with self.assertRaisesRegex(BrowserError, "未设置为"):
            BskBrowser._verify_schedule(task, state)

    def test_current_topic_dialog_selects_exact_platform_card(self):
        browser = BskBrowser()
        task = Task(2, "topic", Path("video.mp4"), "标题", "", (), None,
                    (), (), "山地骑行的快乐", None, "内容含营销信息")
        actions = []

        def click(role, name):
            actions.append(("click", role, name))
            if name == "点击添加话题":
                return 'dialog "话题选择"\n@e1 textbox "输入关键词搜索"'
            if name == "搜索":
                return '@e4 button "山地骑行的快乐 时尚 户外 运动"'
            return 'dialog "话题选择"'

        browser.click = click
        browser.fill = lambda role, name, value: actions.append(("fill", role, name, value))
        browser._run = lambda *args: actions.append(args)
        browser.observe = lambda: 'StaticText "当前选中话题："\nStaticText "山地骑行的快乐"'
        browser._wait_for = lambda predicate, description, seconds: actions.append(("wait", description))
        browser._tags(task)
        self.assertIn(("fill", "textbox", "输入关键词搜索", "山地骑行的快乐"), actions)
        self.assertIn(("click", "--selector",
                       '[data-autolog-container="topic-item-card"]'
                       '[data-autolog="key=topic_item_card&text=可选话题:山地骑行的快乐"]'), actions)
        self.assertIn(("click", "button", "确认提交"), actions)

    def test_content_tag_is_searched_exactly_and_verified_as_editor_token(self):
        browser = BskBrowser()
        task = Task(2, "tag", Path("video.mp4"), "标题", "正文", (), None,
                    (), ("#儿童自行车",), "", None, "内容含营销信息")
        actions = []

        def click(role, name):
            actions.append(("click", role, name))
            return ('@e1 textbox "加个标题让内容更吸引人" [empty] '
                    'placeholder="加个标题让内容更吸引人"\n'
                    '@e2 textbox [empty]\nStaticText "内容标签"')

        def run(*args):
            actions.append(args)
            if args[0] == "evaluate":
                return '{"ok":true,"value":["儿童自行车"]}'
            return ""

        browser.click = click
        browser._run = run
        browser.observe = lambda: '@e3 button "# 儿童自行车"\n@e4 button "# 儿童自行车推荐"'
        browser._tags(task)
        self.assertIn(("fill", "@e2", "--value", "儿童自行车"), actions)
        self.assertIn(("click", "@e3"), actions)
        self.assertNotIn(("click", "@e4"), actions)
        self.assertTrue(any(action[0] == "evaluate" for action in actions))

    def test_content_tag_accepts_platform_result_without_space_after_hash(self):
        browser = BskBrowser()
        task = Task(2, "tag", Path("video.mp4"), "标题", "正文", (), None,
                    (), ("骑行",), "", None, "内容含营销信息")
        actions = []

        def click(role, name):
            actions.append(("click", role, name))
            return '@e1 textbox "加个标题让内容更吸引人" [empty] placeholder="加个标题让内容更吸引人"\n@e2 textbox [empty]\nStaticText "内容标签"'

        def run(*args):
            actions.append(args)
            if args[0] == "evaluate":
                return '{"ok":true,"value":["骑行"]}'
            return ""

        browser.click = click
        browser._run = run
        browser.observe = lambda: '@e3 button "#骑行"'
        browser._tags(task)
        self.assertIn(("click", "@e3"), actions)

    def test_duplicate_warning_is_detected_from_visible_page_state(self):
        browser = BskBrowser()
        browser.observe = lambda: 'StaticText "视频可能重复发布"'
        self.assertEqual(browser.duplicate_warning(seconds=0), "视频可能重复发布")

    def test_prepare_stops_before_filling_when_video_may_be_duplicate(self):
        browser = BskBrowser()
        task = Task(2, "duplicate", Path("video.mp4"), "标题", "正文", (), None,
                    (), (), "", None, "内容含营销信息")
        actions = []
        browser.navigate = lambda: None
        browser._video = lambda _task: actions.append("video")
        browser.duplicate_warning = lambda: "视频可能重复发布"
        browser.fill = lambda *args: actions.append("fill")
        with self.assertRaises(DuplicateVideoWarning):
            browser.prepare(task)
        self.assertEqual(actions, ["video"])

    def test_current_topic_dialog_rejects_missing_platform_topic(self):
        browser = BskBrowser()
        task = Task(2, "topic", Path("video.mp4"), "标题", "", (), None,
                    (), (), "山地车", None, "内容含营销信息")
        browser.click = lambda role, name: ('dialog "话题选择"' if name == "点击添加话题"
                                            else 'StaticText "没有找到相关话题"')
        browser.fill = lambda *args: None
        with self.assertRaisesRegex(BrowserError, "平台没有可选话题"):
            browser._tags(task)


if __name__ == "__main__":
    unittest.main()
