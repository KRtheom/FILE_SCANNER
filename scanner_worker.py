"""검색 워커 — 별도 프로세스에서 실행된다."""
from __future__ import annotations

import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from multiprocessing.connection import Connection
from typing import Any

import scanner_engine

# 메시지 타입
MSG_PROGRESS = "progress"
MSG_RESULT = "result"
MSG_FAIL = "fail"
MSG_DONE = "done"
MSG_STOP = "stop"


def run_search(
    conn: Connection,
    paths: list[str],
    keywords: list[str],
    search_mode: str,
) -> None:
    """conn을 통해 GUI 프로세스와 통신하며 검색을 수행한다."""
    start = time.perf_counter()
    found = 0
    completed = 0
    fail_count = 0
    skip_count = 0
    total_files = 0

    try:
        if search_mode == "filename":
            files = scanner_engine.scan_files(paths, extensions=None)
        else:
            files = scanner_engine.scan_files(
                paths, extensions=scanner_engine.DEFAULT_EXTENSIONS,
            )

        total_files = len(files)
        _send(conn, MSG_PROGRESS, done=0, total=total_files)

        if search_mode == "filename":
            _search_filename(conn, files, keywords)
        else:
            _search_content(conn, files, keywords)

    except Exception:
        fail_count += 1
    finally:
        scanner_engine.clear_all_search_file_meta()
        elapsed = time.perf_counter() - start
        _send(conn, MSG_DONE, elapsed=elapsed)
        conn.close()


def _search_filename(
    conn: Connection, files: list[str], keywords: list[str],
) -> None:
    total = len(files)
    for i, fp in enumerate(files, 1):
        if _check_stop(conn):
            break
        try:
            matches = scanner_engine.search_file_by_name(fp, keywords)
        except Exception:
            _send(conn, MSG_FAIL, file_path=fp, error="검색 처리 중 예외")
            continue
        for m in matches:
            _send(
                conn, MSG_RESULT,
                keyword=m.get("keyword", ""),
                file_path=m.get("file", fp),
                location=m.get("location", ""),
                context="",
            )
        if i % 500 == 0 or i == total:
            _send(conn, MSG_PROGRESS, done=i, total=total)


def _search_content(
    conn: Connection, files: list[str], keywords: list[str],
) -> None:
    total = len(files)
    completed = 0
    worker_count = min(
        scanner_engine.MAX_WORKERS,
        max(scanner_engine.MIN_WORKERS, (os.cpu_count() or 4) // scanner_engine.CPU_DIVISOR),
    )

    with ThreadPoolExecutor(max_workers=worker_count) as executor:
        future_to_file: dict = {}
        pending = iter(files)
        batch = worker_count * 4

        # 초기 배치 제출
        for _ in range(batch):
            fp = next(pending, None)
            if fp is None:
                break
            future_to_file[executor.submit(scanner_engine.search_file, fp, keywords)] = fp

        while future_to_file:
            if _check_stop(conn):
                executor.shutdown(wait=False, cancel_futures=True)
                break

            # as_completed timeout 예외 처리
            done_futures = []
            try:
                for future in as_completed(future_to_file, timeout=1.0):
                    done_futures.append(future)
                    break  # 1개씩 처리 후 stop 체크
            except TimeoutError:
                continue  # 아직 완료된 게 없으면 다시 루프

            for future in done_futures:
                fp = future_to_file.pop(future)
                completed += 1

                try:
                    matches = future.result(timeout=scanner_engine.FILE_PROCESS_TIMEOUT)
                    err = ""
                except TimeoutError:
                    matches = []
                    err = f"시간 초과 ({scanner_engine.FILE_PROCESS_TIMEOUT}s)"
                except Exception:
                    matches = []
                    err = "검색 처리 중 예외"

                meta = scanner_engine.consume_search_file_meta(fp)

                if err:
                    _send(conn, MSG_FAIL, file_path=fp, error=err)
                elif meta.get("failed"):
                    _send(conn, MSG_FAIL, file_path=fp, error="텍스트 추출 실패")

                for m in matches:
                    _send(
                        conn, MSG_RESULT,
                        keyword=m.get("keyword", ""),
                        file_path=m.get("file", fp),
                        location=m.get("location", ""),
                        context=m.get("context", ""),
                    )

                # 다음 파일 제출
                nxt = next(pending, None)
                if nxt is not None:
                    future_to_file[executor.submit(scanner_engine.search_file, nxt, keywords)] = nxt

                if completed % 100 == 0 or completed == total:
                    _send(conn, MSG_PROGRESS, done=completed, total=total)


def _send(conn: Connection, msg_type: str, **data: Any) -> None:
    try:
        conn.send({"type": msg_type, **data})
    except (BrokenPipeError, OSError):
        pass


def _check_stop(conn: Connection) -> bool:
    try:
        while conn.poll(0):
            msg = conn.recv()
            if msg == MSG_STOP:
                return True
    except (EOFError, OSError):
        return True
    return False
