"""Single-thread queue executor for COM-safe serialization (FR-6).

Callers submit callables that run on one dedicated worker thread and **block**
until completion. This module is COM-agnostic: it does not import win32com or
start Excel (FR-10).

**Limitations**

- **Shutdown:** ``shutdown(wait=True)`` drains the queue and joins the worker;
  ``wait=False`` only signals stop—work already queued may be abandoned and the
  thread may still be running briefly. Abrupt process exit (e.g. SIGKILL) does
  not run shutdown; in-flight work on the worker is not guaranteed to finish.
- **Reentrancy:** Do not ``submit`` from *inside* a callable running on the
  worker thread: it would deadlock (the worker would wait on itself).
"""

from __future__ import annotations

import queue
import threading
from typing import Any, Callable, TypeVar

__all__ = ["ComThreadExecutor"]

T = TypeVar("T")
_SENTINEL = object()


class ComThreadExecutor:
    """Marshals callables onto a single worker thread; ``submit`` blocks for the result."""

    def __init__(self) -> None:
        self._q: queue.Queue[Callable[[], None] | object] = queue.Queue()
        self._mutex = threading.Lock()
        self._thread: threading.Thread | None = None
        self._shutdown = False

    def submit(self, fn: Callable[..., T], /, *args: Any, **kwargs: Any) -> T:
        """Run ``fn(*args, **kwargs)`` on the worker thread; return its result or re-raise."""
        done = threading.Event()
        out: dict[str, Any] = {}

        def wrapper() -> None:
            try:
                out["result"] = fn(*args, **kwargs)
            except BaseException as exc:
                out["exc"] = exc
            finally:
                done.set()

        with self._mutex:
            if self._shutdown:
                msg = "cannot schedule new futures after shutdown"
                raise RuntimeError(msg)
            if self._thread is None or not self._thread.is_alive():
                t = threading.Thread(
                    target=self._worker_loop,
                    name="ComThreadExecutor",
                    daemon=True,
                )
                self._thread = t
                t.start()
            self._q.put(wrapper)

        done.wait()
        if "exc" in out:
            raise out["exc"]
        return out["result"]

    def _worker_loop(self) -> None:
        while True:
            item = self._q.get()
            try:
                if item is _SENTINEL:
                    break
                assert callable(item)
                item()
            finally:
                self._q.task_done()

    def shutdown(self, *, wait: bool = True) -> None:
        """Stop accepting new work; optionally join the worker after it drains.

        Safe to call more than once. After shutdown, :meth:`submit` raises
        ``RuntimeError``.
        """
        with self._mutex:
            if self._shutdown:
                return
            self._shutdown = True
            if self._thread is None or not self._thread.is_alive():
                return
            self._q.put(_SENTINEL)

        if wait and self._thread is not None:
            self._thread.join()
