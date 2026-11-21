"""
Process pool for the per-file step of the analysis scripts.

The per-file work (Excel parsing + pandas loops) is pure-Python-heavy, so a
ThreadPoolExecutor can't use more than ~1 core because of the GIL. This runs
each file in a separate worker process instead.

Each worker loads the analysis script by path and copies in the module-level
state the parent computed at runtime (for example T0_T1_PAIRS), so worker
functions see the same globals they would in-process.

    with FilePool(__file__, {"T0_PAIRS": T0_PAIRS}, max_workers) as pool:
        fut = pool.submit("_process_single_file", path, plate_map, max_time)
"""

import importlib.util
import os
import sys
from concurrent.futures import ProcessPoolExecutor
from concurrent.futures.process import BrokenProcessPool
from datetime import datetime

# Windows can't wait on more than 61 worker processes in one pool.
MAX_POOL_WORKERS = 61 if sys.platform == "win32" else None

WORKER_CRASHED = (
    "A worker process crashed, most often because the machine ran out of memory. "
    "Lower 'Total worker budget' in the launcher and run again."
)


class WorkerCrashedError(RuntimeError):
    pass


def cap_workers(n):
    n = max(1, int(n))
    return min(n, MAX_POOL_WORKERS) if MAX_POOL_WORKERS else n

_worker_module = None


def _init_worker(script_path, module_state):
    global _worker_module
    stem = os.path.splitext(os.path.basename(script_path))[0]
    spec = importlib.util.spec_from_file_location(f"_file_pool_{stem}", script_path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    for name, value in (module_state or {}).items():
        setattr(mod, name, value)
    _worker_module = mod


def _call(func_name, args):
    return getattr(_worker_module, func_name)(*args)


def tag_suffix(file_tag):
    """'_DMD' for file tag 'DMD', '' for none: appended to output file names so a
    plate's workbooks show which plate they came from (d3_75Hz_waveformkinetics_DMD.xlsx)."""
    file_tag = (file_tag or "").strip()
    return f"_{file_tag}" if file_tag else ""


def reserve_output_paths(paths):
    """
    Pick a unique output path for each requested path, in order, before the
    workbooks are written in parallel. A path that already exists on disk, or
    was already handed out earlier in this list, gets a timestamp suffix (the
    same naming make_timestamped_path_if_exists uses). Comparison ignores case,
    because the macOS and Windows file systems do.
    """
    claimed = set()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    result = []
    for path in paths:
        base, ext = os.path.splitext(path)
        candidate = path
        counter = 0
        while os.path.exists(candidate) or os.path.normcase(candidate).lower() in claimed:
            candidate = f"{base}_{stamp}{ext}" if counter == 0 else f"{base}_{stamp}_{counter}{ext}"
            counter += 1
        claimed.add(os.path.normcase(candidate).lower())
        result.append(candidate)
    return result


class FilePool:
    def __init__(self, script_path, module_state, max_workers):
        self._executor = ProcessPoolExecutor(
            max_workers=cap_workers(max_workers),
            initializer=_init_worker,
            initargs=(os.path.abspath(script_path), dict(module_state or {})),
        )

    def submit(self, func_name, *args):
        return self._executor.submit(_call, func_name, args)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self._executor.shutdown(wait=True, cancel_futures=exc_type is not None)
        if exc_type is not None and issubclass(exc_type, BrokenProcessPool):
            # Don't let a crash pass as an ordinary per-file error: the scripts
            # skip those and carry on, and the run would be recorded as done.
            raise WorkerCrashedError(WORKER_CRASHED) from exc
        return False


def raise_if_crashed(exc):
    """Call from a per-file `except Exception` handler before logging and skipping."""
    if isinstance(exc, BrokenProcessPool):
        raise WorkerCrashedError(WORKER_CRASHED) from exc
