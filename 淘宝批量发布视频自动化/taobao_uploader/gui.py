from __future__ import annotations

import queue
from pathlib import Path
import sqlite3
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .folder_drafts import generate_folder_draft
from .publish_stats import DEFAULT_STATS_DIR, PublicationStats
from .runner import BatchRunner
from .state import StateStore
from .tasks import Task, TaskValidationError, read_tasks


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("淘宝视频批量发布")
        self.geometry("1120x650")
        self.minsize(880, 520)
        self.tasks: list[Task] = []
        self.store: StateStore | None = None
        self.runner: BatchRunner | None = None
        self.worker: threading.Thread | None = None
        self.events: queue.Queue[tuple[str, str, str]] = queue.Queue()
        self.workbook = tk.StringVar()
        self.dry_run = tk.BooleanVar(value=True)
        self._build()
        self.after(150, self._drain)
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _build(self) -> None:
        top = ttk.Frame(self, padding=12)
        top.pack(fill="x")
        ttk.Label(top, text="任务 Excel").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.workbook).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(top, text="选择并导入", command=self._choose).grid(row=0, column=2)
        ttk.Button(top, text="按文件夹生成草稿", command=self._generate_from_folder).grid(
            row=0, column=3, padx=(8, 0))
        top.columnconfigure(1, weight=1)

        actions = ttk.Frame(self, padding=(12, 0, 12, 10))
        actions.pack(fill="x")
        self.start_button = ttk.Button(actions, text="开始", command=self._start)
        self.start_button.pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="暂停", command=self._pause).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="继续", command=self._resume).pack(side="left", padx=(0, 8))
        ttk.Button(actions, text="重试失败", command=self._retry).pack(side="left")
        ttk.Button(actions, text="更新统计表", command=self._export_stats).pack(side="left", padx=(8, 0))
        ttk.Checkbutton(actions, text="仅验证填写，不提交", variable=self.dry_run).pack(side="left", padx=16)
        self.summary = ttk.Label(actions, text="请导入任务 Excel")
        self.summary.pack(side="right")

        columns = ("行", "编号", "视频", "标题", "话题", "方式", "状态", "说明")
        self.table = ttk.Treeview(self, columns=columns, show="headings")
        widths = (52, 100, 180, 220, 115, 165, 85, 260)
        for name, width in zip(columns, widths):
            self.table.heading(name, text=name)
            self.table.column(name, width=width, minwidth=50, stretch=name in {"视频", "标题", "说明"})
        self.table.pack(fill="both", expand=True, padx=12)
        scroll = ttk.Scrollbar(self.table, orient="vertical", command=self.table.yview)
        self.table.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        ttk.Label(self, text="说明：提交后状态不明的任务不会自动重试，请先到光合后台核查。",
                  padding=(12, 8)).pack(fill="x")

    def _choose(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("正在运行", "请先暂停并等待当前任务结束")
            return
        filename = filedialog.askopenfilename(title="选择任务 Excel", filetypes=[("Excel 工作簿", "*.xlsx")])
        if filename:
            self.workbook.set(filename)
            self._load(Path(filename))

    def _generate_from_folder(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("正在运行", "请等待当前任务结束后再生成草稿")
            return
        folder = filedialog.askdirectory(title="选择包含车型子文件夹的视频目录")
        if not folder:
            return
        drafts_dir = Path(__file__).resolve().parents[1] / "outputs" / "folder_drafts"
        drafts_dir.mkdir(parents=True, exist_ok=True)
        filename = filedialog.asksaveasfilename(
            title="保存发布任务草稿", initialdir=drafts_dir,
            initialfile="发布任务草稿.xlsx", defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")])
        if not filename:
            return
        reference_db = (Path(__file__).resolve().parents[1] / "outputs" /
                        "video_reference_db" / "视频发布参考数据库.sqlite3")
        try:
            result = generate_folder_draft(Path(folder), reference_db, Path(filename))
        except (ValueError, OSError) as exc:
            messagebox.showerror("草稿生成失败", str(exc))
            return
        self.workbook.set(str(result.output))
        self.summary.configure(text=f"草稿 {result.total} 条：车型匹配 {result.model_matched}，通用 {result.generic}")
        messagebox.showinfo(
            "草稿已生成",
            f"已保存：{result.output}\n\n"
            f"车型匹配 {result.model_matched} 条，通用标题 {result.generic} 条。\n"
            "请先打开草稿，核对标题和话题并填写创作者声明，再点击“选择并导入”。\n"
            "草稿生成不会上传或发布视频。")

    def _load(self, path: Path) -> None:
        try:
            tasks = read_tasks(path)
            new_store = StateStore(path.parent / ".taobao_uploader" / "tasks.sqlite3")
            new_store.register(tasks)
            new_store.recover_interrupted(tasks)
        except (TaskValidationError, ValueError, OSError) as exc:
            messagebox.showerror("导入失败", str(exc))
            return
        if self.store:
            self.store.close()
        self.store = new_store
        self.tasks = tasks
        self.table.delete(*self.table.get_children())
        for task in tasks:
            status, message, _ = new_store.get(task.task_id)
            mode = task.publish_at.strftime("定时 %Y-%m-%d %H:%M") if task.publish_at else "立即发布"
            self.table.insert("", "end", iid=task.task_id,
                              values=(task.row, task.task_id, task.video.name, task.title,
                                      task.topic, mode, status, message))
        self.summary.configure(text=f"共 {len(tasks)} 条任务")

    def _notify(self, task_id: str, status: str, message: str) -> None:
        self.events.put((task_id, status, message))

    def _start(self) -> None:
        if not self.tasks or not self.store:
            messagebox.showwarning("没有任务", "请先导入任务 Excel")
            return
        if self.worker and self.worker.is_alive():
            return
        self.runner = BatchRunner(self.tasks, self.store, self._notify,
                                  dry_run=self.dry_run.get(),
                                  source_workbook=Path(self.workbook.get()))
        self.worker = threading.Thread(target=self.runner.run, daemon=True)
        self.worker.start()
        self.start_button.configure(state="disabled")

    def _pause(self) -> None:
        if self.runner:
            self.runner.request_pause()
            self.summary.configure(text="将在提交前或当前任务结束后暂停")

    def _resume(self) -> None:
        if self.worker and self.worker.is_alive():
            if self.runner:
                self.runner.resume()
                self.summary.configure(text="继续执行")
            return
        self._start()

    def _retry(self) -> None:
        if not self.store or not self.tasks:
            return
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("正在运行", "请先暂停并等待当前任务结束")
            return
        count = self.store.retry_failed(self.tasks)
        for task in self.tasks:
            if self.table.exists(task.task_id):
                status, message, _ = self.store.get(task.task_id)
                self.table.set(task.task_id, "状态", status)
                self.table.set(task.task_id, "说明", message)
        self.summary.configure(text=f"已将 {count} 条失败任务设为待执行")

    def _export_stats(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("正在运行", "请等待当前任务结束后再更新统计表")
            return
        stats = None
        try:
            stats = PublicationStats(DEFAULT_STATS_DIR)
            if self.tasks and self.store and self.workbook.get():
                source = Path(self.workbook.get())
                stats.backfill_successes(
                    self.tasks, source.parent / ".taobao_uploader" / "tasks.sqlite3", source)
            path = stats.export()
        except (OSError, ValueError, sqlite3.Error) as exc:
            messagebox.showerror("更新统计表失败", f"{exc}\n如已在 Excel 打开统计表，请先关闭后重试。")
            return
        finally:
            if stats is not None:
                stats.close()
        messagebox.showinfo("统计表已更新", str(path))

    def _drain(self) -> None:
        try:
            while True:
                task_id, status, message = self.events.get_nowait()
                if task_id and self.table.exists(task_id) and status not in {
                        "重试", "统计", "统计失败"}:
                    self.table.set(task_id, "状态", status)
                    self.table.set(task_id, "说明", message)
                if status in {"错误", "结束", "日志", "重试", "统计", "统计失败"}:
                    self.summary.configure(text=message[:110])
                if status == "结束":
                    self.start_button.configure(state="normal")
        except queue.Empty:
            pass
        self.after(150, self._drain)

    def _close(self) -> None:
        if self.worker and self.worker.is_alive():
            if self.runner:
                self.runner.request_stop()
            messagebox.showinfo("正在执行", "已请求停止。请等待当前浏览器操作结束后再关闭窗口。")
            return
        if self.store:
            self.store.close()
        self.destroy()


def main() -> None:
    App().mainloop()
