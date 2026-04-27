"""ComThreadExecutor: serialization and API (Story 6-2; no Excel / pywin32)."""

from __future__ import annotations

import threading
import time
from concurrent.futures import ThreadPoolExecutor

import pytest

from excel_mcp.com_executor import ComThreadExecutor


def test_submit_runs_callable_and_returns_result():
    ex = ComThreadExecutor()
    try:
        assert ex.submit(lambda x, y: x + y, 2, y=3) == 5
    finally:
        ex.shutdown(wait=True)


def test_submit_propagates_exception():
    ex = ComThreadExecutor()

    def boom() -> None:
        raise ValueError("expected")

    try:
        with pytest.raises(ValueError, match="expected"):
            ex.submit(boom)
    finally:
        ex.shutdown(wait=True)


def test_concurrent_submits_use_one_worker_thread():
    ex = ComThreadExecutor()
    idents: list[int] = []

    def record(_: int) -> None:
        idents.append(threading.get_ident())
        time.sleep(0.01)

    try:

        def one(i: int) -> None:
            ex.submit(record, i)

        with ThreadPoolExecutor(max_workers=8) as pool:
            list(pool.map(one, range(8)))
    finally:
        ex.shutdown(wait=True)

    assert len(idents) == 8
    assert len(set(idents)) == 1
    assert idents[0] != threading.get_ident()


def test_concurrent_submits_do_not_interleave_mid_callable():
    """If two worker threads existed, sleep in task A would let B mutate shared state."""
    ex = ComThreadExecutor()
    counter = [0]

    def task_a() -> None:
        counter[0] += 1
        local = counter[0]
        time.sleep(0.08)
        assert counter[0] == local, "another task ran during task_a"

    def task_b() -> None:
        counter[0] += 100

    barrier = threading.Barrier(2)
    err: list[BaseException] = []

    def run_a() -> None:
        try:
            barrier.wait()
            ex.submit(task_a)
        except BaseException as e:
            err.append(e)

    def run_b() -> None:
        try:
            barrier.wait()
            ex.submit(task_b)
        except BaseException as e:
            err.append(e)

    t1 = threading.Thread(target=run_a)
    t2 = threading.Thread(target=run_b)
    t1.start()
    t2.start()
    t1.join()
    t2.join()
    try:
        assert not err, err
        assert counter[0] == 101  # A then B or B then A; always fully sequential
    finally:
        ex.shutdown(wait=True)


def test_shutdown_rejects_subsequent_submit():
    ex = ComThreadExecutor()
    ex.submit(lambda: None)
    ex.shutdown(wait=True)
    with pytest.raises(RuntimeError, match="shutdown"):
        ex.submit(lambda: None)


def test_double_shutdown_is_safe():
    ex = ComThreadExecutor()
    ex.shutdown(wait=True)
    ex.shutdown(wait=True)
