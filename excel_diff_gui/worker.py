"""バックグラウンドワーカー（スレッド + キュー）。"""
import queue
import threading
from typing import Callable, Any, Optional


class Worker:
    """1つのワーカースレッドでタスクを順番に実行する。"""

    def __init__(self) -> None:
        self._q: queue.Queue = queue.Queue()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def _loop(self) -> None:
        while True:
            fn, args, kwargs, result_q = self._q.get()
            try:
                result_q.put(("ok", fn(*args, **kwargs)))
            except Exception as exc:
                result_q.put(("err", exc))

    def submit(self, fn: Callable, *args, **kwargs) -> queue.Queue:
        """fnをバックグラウンドで実行し、結果Queueを返す。"""
        result_q: queue.Queue = queue.Queue()
        self._q.put((fn, args, kwargs, result_q))
        return result_q


_worker: Optional[Worker] = None


def get_worker() -> Worker:
    global _worker
    if _worker is None:
        _worker = Worker()
    return _worker
