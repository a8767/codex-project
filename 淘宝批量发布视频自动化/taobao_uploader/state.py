from __future__ import annotations

from pathlib import Path
import sqlite3
from threading import RLock

from .tasks import Task


class StateStore:
    def __init__(self, path: Path):
        path.parent.mkdir(parents=True, exist_ok=True)
        self.lock = RLock()
        self.db = sqlite3.connect(path, check_same_thread=False)
        self.db.execute("""CREATE TABLE IF NOT EXISTS tasks (
            task_id TEXT PRIMARY KEY, fingerprint TEXT NOT NULL, status TEXT NOT NULL,
            message TEXT NOT NULL DEFAULT '', attempts INTEGER NOT NULL DEFAULT 0,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )""")
        self.db.commit()

    def register(self, tasks: list[Task]) -> None:
        with self.lock:
            for task in tasks:
                row = self.db.execute("SELECT fingerprint, status FROM tasks WHERE task_id=?", (task.task_id,)).fetchone()
                if row and row[0] != task.fingerprint:
                    if row[1] in {"提交中", "待核查"}:
                        raise ValueError(f"任务 {task.task_id} 的内容已变化；请用新任务编号，避免重复发布")
                    self.db.execute("""UPDATE tasks SET fingerprint=?, status='待执行', message='任务内容已更新',
                        updated_at=CURRENT_TIMESTAMP WHERE task_id=?""", (task.fingerprint, task.task_id))
                if not row:
                    self.db.execute("INSERT INTO tasks(task_id,fingerprint,status) VALUES(?,?,'待执行')",
                                    (task.task_id, task.fingerprint))
            self.db.commit()

    def get(self, task_id: str) -> tuple[str, str, int]:
        with self.lock:
            row = self.db.execute("SELECT status,message,attempts FROM tasks WHERE task_id=?", (task_id,)).fetchone()
            if row is None:
                raise KeyError(task_id)
            return row[0], row[1], row[2]

    def set(self, task_id: str, status: str, message: str = "", *, increment: bool = False) -> None:
        with self.lock:
            self.db.execute("""UPDATE tasks SET status=?, message=?, attempts=attempts+?,
                updated_at=CURRENT_TIMESTAMP WHERE task_id=?""", (status, message, int(increment), task_id))
            self.db.commit()

    def recover_interrupted(self, tasks: list[Task]) -> None:
        with self.lock:
            for task in tasks:
                status, _, _ = self.get(task.task_id)
                if status == "提交中":
                    self.set(task.task_id, "待核查", "上次运行在提交阶段中断；请到后台核对后人工处理")
                elif status == "执行中":
                    self.set(task.task_id, "待执行", "上次运行在提交前中断")

    def retry_failed(self, tasks: list[Task]) -> int:
        count = 0
        for task in tasks:
            status, _, _ = self.get(task.task_id)
            if status == "失败":
                self.set(task.task_id, "待执行", "")
                count += 1
        return count

    def close(self) -> None:
        with self.lock:
            self.db.close()
