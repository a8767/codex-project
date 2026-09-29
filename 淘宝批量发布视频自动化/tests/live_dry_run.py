"""用指定本地视频验证光合发布器填写流程；始终不提交发布。"""

from pathlib import Path
import sys

from taobao_uploader.runner import BatchRunner
from taobao_uploader.state import StateStore
from taobao_uploader.tasks import Task


def main() -> int:
    if len(sys.argv) != 2:
        print("用法：python tests/live_dry_run.py <视频绝对路径>")
        return 2
    video = Path(sys.argv[1]).resolve()
    if not video.is_file():
        print(f"文件不存在：{video}")
        return 2
    task = Task(2, "live-dry-run", video, "路在脚下 骑在当秋", "", (), None,
                (), (), "", None, "内容无需标注")
    store = StateStore(Path(__file__).resolve().parent.parent / ".test_run" / "state.sqlite3")
    try:
        store.register([task])
        store.recover_interrupted([task])
        store.retry_failed([task])
        runner = BatchRunner([task], store, lambda ident, status, msg: print(f"{ident} {status}: {msg}", flush=True),
                             dry_run=True)
        runner.run()
        status, message, _ = store.get(task.task_id)
        print(f"最终状态：{status}，{message}", flush=True)
        return 0 if status == "验证通过" else 1
    finally:
        store.close()


if __name__ == "__main__":
    raise SystemExit(main())
