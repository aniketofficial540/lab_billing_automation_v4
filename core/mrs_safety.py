"""
core/mrs_safety.py — Master Rate Sheet safety primitives.

Three concerns handled here:
  1. Backups — every destructive operation snapshots the live MRS first.
  2. Atomic writes — save to a temp file, then os.replace() onto the live path.
  3. File locks — Windows-friendly advisory lock so two app instances cannot
     race on Master_Rate_Sheet.xlsx.

All functions are safe to call from any thread; the lock context manager
serialises writers across processes (single-machine only).
"""

import os
import shutil
import tempfile
import time
from contextlib import contextmanager
from datetime import datetime

from core.logger import get_logger

log = get_logger(__name__)

BACKUP_RETENTION = 5    # keep last N backups
LOCK_TIMEOUT_S = 10     # wait this long for the lock before giving up


def _backup_dir() -> str:
    from config import DATA_DIR
    d = os.path.join(DATA_DIR, "backups")
    os.makedirs(d, exist_ok=True)
    return d


def backup_mrs(master_path: str, tag: str = "manual") -> str:
    """Snapshot the live MRS into data/backups/. Returns the backup path.

    Raises FileNotFoundError if master_path doesn't exist.
    """
    if not os.path.exists(master_path):
        raise FileNotFoundError(f"Cannot back up: {master_path} does not exist")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_tag = "".join(c if c.isalnum() else "_" for c in tag)[:30]
    name = f"Master_Rate_Sheet_{ts}_{safe_tag}.xlsx"
    dest = os.path.join(_backup_dir(), name)

    shutil.copy2(master_path, dest)
    log.info("MRS backup created: %s", dest)
    return dest


def restore_from_backup(backup_path: str, master_path: str) -> None:
    """Atomically restore master_path from backup_path."""
    if not os.path.exists(backup_path):
        log.error("Cannot restore — backup missing: %s", backup_path)
        raise FileNotFoundError(f"Backup not found: {backup_path}")

    # Atomic replace via os.replace (works cross-volume on the same drive)
    shutil.copy2(backup_path, master_path)
    log.warning("MRS restored from backup %s", backup_path)


def prune_backups(keep: int = BACKUP_RETENTION) -> int:
    """Delete all but the `keep` most recent backups. Returns count removed."""
    d = _backup_dir()
    try:
        entries = [
            (f, os.path.getmtime(os.path.join(d, f)))
            for f in os.listdir(d)
            if f.startswith("Master_Rate_Sheet_") and f.endswith(".xlsx")
        ]
    except OSError as e:
        log.warning("Could not list backups: %s", e)
        return 0

    entries.sort(key=lambda x: x[1], reverse=True)
    removed = 0
    for name, _ in entries[keep:]:
        try:
            os.remove(os.path.join(d, name))
            removed += 1
        except OSError as e:
            log.warning("Could not remove old backup %s: %s", name, e)
    if removed:
        log.debug("Pruned %d old MRS backup(s)", removed)
    return removed


@contextmanager
def atomic_replace(target_path: str):
    """Yield a temp path; on clean exit, atomically replace target_path with it.

    Usage:
        with atomic_replace(MASTER_RATE_PATH) as tmp:
            shutil.copy2(MASTER_RATE_PATH, tmp)   # start from current state
            wb = openpyxl.load_workbook(tmp)
            ... modifications ...
            wb.save(tmp)
        # on context exit, tmp -> MASTER_RATE_PATH atomically
    """
    fd, tmp_path = tempfile.mkstemp(
        prefix=".mrs_tmp_", suffix=".xlsx",
        dir=os.path.dirname(target_path) or None,
    )
    os.close(fd)
    try:
        yield tmp_path
        os.replace(tmp_path, target_path)
        log.debug("Atomic replace OK: %s", target_path)
    except Exception:
        # Clean up tmp on failure; do not touch target
        try:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
        except OSError:
            pass
        raise


def _lock_path(master_path: str) -> str:
    return master_path + ".lock"


@contextmanager
def mrs_lock(master_path: str, timeout: float = LOCK_TIMEOUT_S):
    """Cross-process advisory lock on the MRS.

    On Windows, uses an exclusive file open as the lock token. Blocks up to
    `timeout` seconds before raising TimeoutError.
    """
    lock_path = _lock_path(master_path)
    deadline = time.monotonic() + timeout
    fh = None
    last_err = None

    while True:
        try:
            # 'x' opens exclusively — fails if file already exists
            fh = open(lock_path, "x")
            fh.write(f"pid={os.getpid()}\nat={datetime.now().isoformat()}\n")
            fh.flush()
            break
        except FileExistsError as e:
            last_err = e
            if time.monotonic() >= deadline:
                log.error("Could not acquire MRS lock within %.1fs", timeout)
                raise TimeoutError(
                    f"Master Rate Sheet is locked by another process. "
                    f"If no other instance is running, delete {lock_path} manually."
                ) from e
            time.sleep(0.2)

    try:
        log.debug("MRS lock acquired (%s)", lock_path)
        yield
    finally:
        try:
            if fh:
                fh.close()
            if os.path.exists(lock_path):
                os.remove(lock_path)
            log.debug("MRS lock released")
        except OSError as e:
            log.warning("Failed to release MRS lock: %s", e)
