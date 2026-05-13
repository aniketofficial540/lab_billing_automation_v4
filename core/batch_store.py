import json
import os
from datetime import datetime
from config import BATCH_SESSION_DIR, LAST_BATCH_SESSION_PATH
from core.logger import get_logger

log = get_logger(__name__)


def create_batch_session_dir() -> tuple[str, str]:
    session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    os.makedirs(BATCH_SESSION_DIR, exist_ok=True)
    session_dir = os.path.join(BATCH_SESSION_DIR, f"batch_{session_id}")
    os.makedirs(session_dir, exist_ok=True)
    return session_dir, session_id


def save_session_manifest(session_dir: str, manifest: dict) -> str:
    path = os.path.join(session_dir, "batch_session.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2)
    _write_latest_pointer(path)
    log.info("Batch session saved: %s", path)
    return path


def load_session_manifest(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_latest_session_path() -> str | None:
    if not os.path.exists(LAST_BATCH_SESSION_PATH):
        return None
    try:
        with open(LAST_BATCH_SESSION_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        path = data.get("session_path")
        return path if path and os.path.exists(path) else None
    except Exception as e:
        log.warning("Failed to read latest batch session pointer: %s", e)
        return None


def _write_latest_pointer(session_path: str) -> None:
    os.makedirs(os.path.dirname(LAST_BATCH_SESSION_PATH), exist_ok=True)
    data = {
        "session_path": session_path,
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }
    with open(LAST_BATCH_SESSION_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
