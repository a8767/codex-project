from __future__ import annotations

from threading import Event
from time import sleep
from typing import Callable
from datetime import datetime
from pathlib import Path

from .browser import (BskBrowser, BrowserError, DuplicateVideoWarning, PauseForHuman,
                      SubmissionUncertain, VerifiedPublication)
from .publish_stats import DEFAULT_STATS_DIR, PublicationStats
from .state import StateStore
from .tasks import TZ, Task


class BatchRunner:
    def __init__(self, tasks: list[Task], store: StateStore,
                 notify: Callable[[str, str, str], None],
                 browser_factory: Callable[[], BskBrowser] | None = None,
                 *, dry_run: bool = False, source_workbook: Path | None = None,
                 stats_dir: Path = DEFAULT_STATS_DIR,
                 pause_on_uncertain: bool = True):
        self.tasks = tasks
        self.store = store
        self.notify = notify
        self.browser_factory = browser_factory or (lambda: BskBrowser(lambda s: notify("", "日志", s)))
        self.dry_run = dry_run
        self.source_workbook = source_workbook
        self.stats_dir = stats_dir
        self.pause_on_uncertain = pause_on_uncertain
        self.pause = Event()
        self.stop = Event()

    def request_pause(self) -> None:
        self.pause.set()

    def request_stop(self) -> None:
        self.stop.set()
        self.pause.clear()

    def resume(self) -> None:
        self.pause.clear()

    def _wait_while_paused(self) -> bool:
        if self.pause.is_set():
            self.notify("", "暂停", "已暂停；处理页面后点击继续")
        while self.pause.is_set() and not self.stop.is_set():
            sleep(0.2)
        return not self.stop.is_set()

    def _set(self, task: Task, status: str, message: str = "", increment: bool = False) -> None:
        self.store.set(task.task_id, status, message, increment=increment)
        self.notify(task.task_id, status, message)

    def run(self) -> None:
        browser = self.browser_factory()
        try:
            browser.start()
            for task in self.tasks:
                if not self._wait_while_paused():
                    break
                status, old_message, _ = self.store.get(task.task_id)
                if status in {"待核查", "失败"}:
                    continue
                attempt = 0
                prepared = False
                skipped_duplicate = False
                while attempt < 3 and not self.stop.is_set():
                    try:
                        self._set(task, "执行中", "准备页面", increment=attempt == 0)
                        browser.prepare(task)
                        prepared = True
                        break
                    except DuplicateVideoWarning as exc:
                        self._set(task, "跳过", str(exc))
                        stats = PublicationStats(self.stats_dir)
                        try:
                            stats.record_skipped(
                                task, source_excel=self.source_workbook,
                                skipped_at=datetime.now(TZ), remark="视频重复上传",
                            )
                            stats_path = stats.export()
                            self.notify(task.task_id, "统计", f"重复视频跳过记录已登记：{stats_path}")
                        except Exception as stats_exc:
                            self.notify(task.task_id, "统计失败",
                                        f"任务已跳过；统计数据库可能已保存，但工作簿未能更新：{stats_exc}")
                        finally:
                            stats.close()
                        skipped_duplicate = True
                        break
                    except PauseForHuman as exc:
                        self._set(task, "待执行", str(exc))
                        self.pause.set()
                        if not self._wait_while_paused():
                            break
                    except BrowserError as exc:
                        attempt += 1
                        if attempt >= 3:
                            self._set(task, "失败", str(exc))
                        else:
                            self.notify(task.task_id, "重试", f"准备阶段第 {attempt} 次失败：{exc}")
                            sleep(2)
                if self.stop.is_set():
                    break
                if skipped_duplicate:
                    continue
                if not prepared:
                    continue
                if self.dry_run:
                    self._set(task, "验证通过", "页面填写完成；未提交发布")
                    continue
                # A pause requested during preparation takes effect before any submission.
                if self.pause.is_set():
                    self._set(task, "待执行", "提交前暂停")
                    if not self._wait_while_paused():
                        break
                if self.stop.is_set():
                    self._set(task, "待执行", "提交前停止")
                    break
                self._set(task, "提交中", "正在提交，等待平台明确结果")
                submitted_at = datetime.now(TZ)
                try:
                    result = browser.submit_and_verify(task)
                    self._set(task, "成功", result.message if isinstance(result, VerifiedPublication)
                              else str(result))
                    if isinstance(result, VerifiedPublication):
                        stats = None
                        try:
                            stats = PublicationStats(self.stats_dir)
                            stats.record(task, result.work_id,
                                         source_excel=self.source_workbook,
                                         submitted_at=submitted_at,
                                         verified_at=result.verified_at,
                                         audit_status=result.audit_status)
                            stats_path = stats.export()
                            self.notify(task.task_id, "统计", f"已更新发布视频统计表：{stats_path}")
                        except Exception as exc:
                            self.notify(task.task_id, "统计失败",
                                        f"作品已发布，统计表暂未更新：{exc}。关闭 Excel 后重新导出。")
                        finally:
                            if stats is not None:
                                stats.close()
                except SubmissionUncertain as exc:
                    self._set(task, "待核查", str(exc))
                    if self.pause_on_uncertain:
                        self.pause.set()
                    continue
                except BrowserError as exc:
                    self._set(task, "待核查", f"提交结果不明确：{exc}")
                    if self.pause_on_uncertain:
                        self.pause.set()
                    continue
        except BrowserError as exc:
            self.pause.set()
            self.notify("", "错误", str(exc))
        finally:
            browser.close()
            self.notify("", "结束", "批量执行已结束" if not self.pause.is_set() else "已暂停")
