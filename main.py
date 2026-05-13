import sys

# Bootstrap persistent user-data dirs and seed bundled MRS on first run.
# Must run BEFORE the logger or any UI module imports — they assume these
# paths already exist.
from config import init_user_data
init_user_data()

from core.logger import get_logger
from ui.main_window import launch


log = get_logger(__name__)

def _excepthook(exc_type, exc_value, tb):
    """Last-resort handler so uncaught exceptions land in app.log."""
    log.critical("Uncaught exception", exc_info=(exc_type, exc_value, tb))
    sys.__excepthook__(exc_type, exc_value, tb)

    
if __name__ == "__main__":
    sys.excepthook = _excepthook
    log.info("Application starting (frozen=%s)", getattr(sys, "frozen", False))
    try:
        launch()
    finally:
        log.info("Application exiting")
