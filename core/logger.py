"""
core/logger.py — Application-wide logging.

Writes to output/app.log with daily rotation. Keeps last 14 days.
Console handler at WARNING level (visible to dev), file handler at DEBUG.

Usage:
    from core.logger import get_logger
    log = get_logger(__name__)
    log.info("CSV uploaded: %s", path)
    log.error("Failed to save MRS", exc_info=True)
"""

import logging
import os
import sys
from logging.handlers import TimedRotatingFileHandler


_LOGGERS_CONFIGURED = False


def _output_dir() -> str:
    from config import OUTPUT_DIR
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    return OUTPUT_DIR


def _configure_root() -> None:
    """One-time configuration of root logger. Idempotent."""
    global _LOGGERS_CONFIGURED
    if _LOGGERS_CONFIGURED:
        return

    root = logging.getLogger("lab_billing")
    root.setLevel(logging.DEBUG)
    root.propagate = False

    # File handler — verbose, rotates daily, 14-day retention
    log_path = os.path.join(_output_dir(), "app.log")
    fh = TimedRotatingFileHandler(
        log_path, when="midnight", backupCount=14, encoding="utf-8"
    )
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(
        "%(asctime)s | %(levelname)-7s | %(name)-30s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    ))
    root.addHandler(fh)

    # Console handler — warnings and above only
    ch = logging.StreamHandler(sys.stderr)
    ch.setLevel(logging.WARNING)
    ch.setFormatter(logging.Formatter("%(levelname)s: %(name)s: %(message)s"))
    root.addHandler(ch)

    _LOGGERS_CONFIGURED = True
    root.info("=" * 60)
    root.info("Logger initialised")


def get_logger(name: str) -> logging.Logger:
    """Return a logger under the lab_billing namespace.

    `name` is typically `__name__` from the calling module.
    """
    _configure_root()
    # Strip leading 'core.' or 'ui.' to keep names short
    if name.startswith(("core.", "ui.")):
        suffix = name
    else:
        suffix = name.rsplit(".", 1)[-1]
    return logging.getLogger("lab_billing." + suffix)
