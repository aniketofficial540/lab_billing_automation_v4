"""
core/addendum_store.py — JSON-based addendum storage.
One file per party: data/addendums/addendum_[ShortCode].json
"""

import os
import json
from datetime import date, datetime
from core.logger import get_logger

log = get_logger(__name__)


def _addendum_dir() -> str:
    from config import ADDENDUM_DIR
    os.makedirs(ADDENDUM_DIR, exist_ok=True)
    return ADDENDUM_DIR

def _path(party_short_code: str) -> str:
    return os.path.join(_addendum_dir(), f"addendum_{party_short_code}.json")


def save_addendum(party_short_code: str, effective_date: date, tests: list[dict]) -> None:
    """Save addendum for a party. Overwrites existing."""
    from config import ADDENDUM_DATE_FORMAT
    data = {
        "party_short_code": party_short_code,
        "effective_date": effective_date.strftime(ADDENDUM_DATE_FORMAT),
        "uploaded_at": datetime.now().isoformat(),
        "tests": tests,
    }
    with open(_path(party_short_code), "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def load_addendum(party_short_code: str) -> dict | None:
    """Load addendum for a party. Returns None if not found."""
    p = _path(party_short_code)
    if not os.path.exists(p):
        return None
    with open(p, "r", encoding="utf-8") as f:
        return json.load(f)


def delete_addendum(party_short_code: str) -> None:
    """Delete addendum for a party."""
    p = _path(party_short_code)
    if os.path.exists(p):
        os.remove(p)


def list_addendums() -> list[dict]:
    """List all stored addendums. Returns list of {party_short_code, effective_date, test_count, uploaded_at}."""
    d = _addendum_dir()
    results = []
    for fname in os.listdir(d):
        if not (fname.startswith("addendum_") and fname.endswith(".json")):
            continue
        full = os.path.join(d, fname)
        try:
            with open(full, "r", encoding="utf-8") as f:
                data = json.load(f)
            results.append({
                "party_short_code": data.get("party_short_code", ""),
                "effective_date": data.get("effective_date", ""),
                "test_count": len(data.get("tests", [])),
                "uploaded_at": data.get("uploaded_at", ""),
            })
        except json.JSONDecodeError as e:
            log.error("Corrupt addendum JSON: %s — %s", fname, e)
        except OSError as e:
            log.error("Could not read addendum %s: %s", fname, e)
    return results


def parse_addendum_excel(path: str) -> list[dict]:
    """Parse Annexure-B Excel into list of {test_code, test_name, discounted_price} dicts.
    Delegates to the existing parser in addendum_writer."""
    from core.addendum_writer import parse_addendum_excel as _parse
    return _parse(path)
