"""
ui/main_window.py — Top-level shell: TitleBar · ProgressBar · TabBar (3 tabs) · MainWindow.
"""

import os
import sys
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QStackedWidget, QFrame,
    QApplication, QMessageBox,
)
from PyQt5.QtCore import Qt, pyqtSlot
from PyQt5.QtGui import QFont, QColor

from config import THEME, FONT_FAMILY, MASTER_RATE_PATH
from ui.widgets import _font


# ── TitleBar ───────────────────────────────────────────────────────────────────

class TitleBar(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(76)              # was 56 — more breathing room
        self.setStyleSheet(f"background: {THEME['primary']}; border: none;")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(28, 12, 28, 12)
        layout.setSpacing(16)

        # Logo tile
        logo_box = QFrame()
        logo_box.setFixedSize(48, 48)
        logo_box.setStyleSheet(
            f"background: {THEME['primary_mid']}; border-radius: 10px; border: none;"
        )
        logo_inner = QLabel("⚕")
        logo_inner.setAlignment(Qt.AlignCenter)
        logo_inner.setFont(_font(22, bold=True))
        logo_inner.setStyleSheet("color: white; background: transparent; border: none;")
        logo_layout = QVBoxLayout(logo_box)
        logo_layout.setContentsMargins(0, 0, 0, 0)
        logo_layout.addWidget(logo_inner)

        # App name + subtitle stack
        name_lbl = QLabel("Amrita Lab Billing Tool")
        name_lbl.setFont(_font(17, bold=True))
        name_lbl.setStyleSheet(
            "color: #FFFFFF; background: transparent; border: none; "
            "letter-spacing: 0.3px;"
        )

        sub_lbl = QLabel("Clinical Laboratory  ·  External Party Billing  ·  Faridabad")
        sub_lbl.setFont(_font(10))
        sub_lbl.setStyleSheet(
            "color: #B7D2EA; background: transparent; border: none; "
            "letter-spacing: 0.2px;"
        )

        text_col = QVBoxLayout()
        text_col.setSpacing(3)
        text_col.setContentsMargins(0, 0, 0, 0)
        text_col.addWidget(name_lbl)
        text_col.addWidget(sub_lbl)

        # MRS-loaded chip
        self._mrs_badge = QLabel("⚠  Master Rate Sheet: not loaded")
        self._mrs_badge.setFont(_font(10, bold=True))
        self._mrs_badge.setMinimumHeight(34)
        self._mrs_badge.setStyleSheet(
            "background: #7B3F00; color: #FAEEDA; border-radius: 17px; "
            "padding: 6px 16px; border: none;"
        )

        layout.addWidget(logo_box)
        layout.addLayout(text_col)
        layout.addStretch()
        layout.addWidget(self._mrs_badge)

    def set_mrs_loaded(self, filename: str):
        self._mrs_badge.setText(f"✓  {filename}")
        self._mrs_badge.setStyleSheet(
            "background: #0F4F30; color: #C4ECDB; border-radius: 17px; "
            "padding: 6px 16px; border: none;"
        )

    def set_mrs_unloaded(self):
        self._mrs_badge.setText("⚠  Master Rate Sheet: not loaded")
        self._mrs_badge.setStyleSheet(
            "background: #7B3F00; color: #FAEEDA; border-radius: 17px; "
            "padding: 6px 16px; border: none;"
        )


# ── Progress Bar ───────────────────────────────────────────────────────────────

class ProgressBar(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(3)
        self._pct = 0.10

        outer = QHBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self._rail = QFrame()
        self._rail.setFixedHeight(3)
        self._rail.setStyleSheet(f"background: {THEME['border']}; border: none;")

        self._fill = QFrame()
        self._fill.setFixedHeight(3)
        self._fill.setStyleSheet(f"background: {THEME['primary_mid']}; border: none;")

        outer.addWidget(self._rail, 0)
        outer.addWidget(self._fill, 0)

    def set_progress(self, pct: float):
        self._pct = max(0.0, min(1.0, pct))
        total = self.width()
        fill_w = int(total * self._pct)
        self._fill.setFixedWidth(fill_w)
        self._rail.setFixedWidth(total - fill_w)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.set_progress(self._pct)


# ── AppTabBar (3 tabs) ─────────────────────────────────────────────────────────

class AppTabBar(QFrame):
    """Three-tab strip below the title bar. Spacious, clear active state."""

    TAB_HEIGHT       = 64    # was 46 — clearly visible, finger-friendly
    CIRCLE_SIZE      = 30    # was 22 — readable step indicator
    ACTIVE_BORDER_PX = 4

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(self.TAB_HEIGHT)
        self.setObjectName("AppTabBar")
        self.setStyleSheet(
            f"QFrame#AppTabBar {{ background: white; "
            f"border-bottom: 1px solid {THEME['border']}; }}"
        )

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self._tabs = []
        self._circles = []
        tab_data = [
            ("D", "⚙", "Documents"),
            ("1", "📊", "Operation 1 — Working Sheet"),
            ("2", "🧾", "Operation 2 — Invoice"),
        ]
        for i, (num, icon, label) in enumerate(tab_data):
            tab = self._make_tab(i, num, icon, label)
            self._tabs.append(tab)
            layout.addWidget(tab, 1)

        self._callbacks = [None, None, None]
        self._active = 0
        self._op1_done = False
        self._op2_done = False
        self._refresh_styles()

    def _make_tab(self, idx: int, num: str, icon: str, label: str) -> QWidget:
        tab = QWidget()
        tab.setCursor(Qt.PointingHandCursor)
        tab.mousePressEvent = lambda e, i=idx: self._on_click(i)

        h = QHBoxLayout(tab)
        h.setContentsMargins(20, 0, 20, 0)
        h.setSpacing(12)
        h.setAlignment(Qt.AlignCenter)

        circle = QLabel(num)
        circle.setFixedSize(self.CIRCLE_SIZE, self.CIRCLE_SIZE)
        circle.setAlignment(Qt.AlignCenter)
        circle.setFont(_font(11, bold=True))
        self._circles.append(circle)

        lbl = QLabel(f"{icon}  {label}")
        lbl.setFont(_font(12, bold=True))   # was 10pt regular — much more readable
        h.addWidget(circle)
        h.addWidget(lbl)
        return tab

    def _on_click(self, idx: int):
        if self._callbacks[idx]:
            self._callbacks[idx]()

    def set_callback(self, idx: int, fn):
        self._callbacks[idx] = fn

    def _refresh_styles(self):
        radius = self.CIRCLE_SIZE // 2
        for i, tab in enumerate(self._tabs):
            circle = self._circles[i]
            lbl = tab.layout().itemAt(1).widget()

            is_done = (i == 1 and self._op1_done) or (i == 2 and self._op2_done)
            is_active = (i == self._active)

            if is_done and not is_active:
                circle.setText("✓")
                circle.setStyleSheet(
                    f"background: {THEME['success']}; color: white; "
                    f"border-radius: {radius}px; border: none; "
                    "font-size: 13pt; font-weight: bold;"
                )
                lbl.setStyleSheet(
                    f"color: {THEME['text_secondary']}; background: transparent; "
                    "font-size: 12pt; font-weight: 500;"
                )
                tab.setStyleSheet(
                    f"background: white; "
                    f"border-bottom: 2px solid {THEME['border']};"
                )
            elif is_active:
                circle.setText(["D", "1", "2"][i])
                circle.setStyleSheet(
                    f"background: {THEME['primary']}; color: white; "
                    f"border-radius: {radius}px; border: none; "
                    "font-size: 12pt; font-weight: bold;"
                )
                lbl.setStyleSheet(
                    f"color: {THEME['primary']}; background: transparent; "
                    "font-size: 12pt; font-weight: 700;"
                )
                tab.setStyleSheet(
                    f"background: {THEME['primary_pale']}; "
                    f"border-bottom: {self.ACTIVE_BORDER_PX}px solid {THEME['primary']};"
                )
            else:
                circle.setText(["D", "1", "2"][i])
                circle.setStyleSheet(
                    f"background: {THEME['surface']}; color: {THEME['text_secondary']}; "
                    f"border-radius: {radius}px; "
                    f"border: 1.5px solid {THEME['border']}; "
                    "font-size: 12pt; font-weight: bold;"
                )
                lbl.setStyleSheet(
                    f"color: {THEME['text_secondary']}; background: transparent; "
                    "font-size: 12pt; font-weight: 500;"
                )
                tab.setStyleSheet(
                    f"background: white; "
                    f"border-bottom: 2px solid {THEME['border']};"
                )

    def switch_to(self, idx: int):
        self._active = idx
        self._refresh_styles()

    def mark_op1_done(self):
        self._op1_done = True
        self._refresh_styles()

    def mark_op2_done(self):
        self._op2_done = True
        self._refresh_styles()

    def reset_done_marks(self):
        self._op1_done = False
        self._op2_done = False
        self._refresh_styles()


# ── MainWindow ─────────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Amrita Lab Billing Tool")
        self.setMinimumSize(1100, 780)
        self.resize(1240, 880)

        self.app_state: dict = {
            "batch_path":  None,
            "batch_session_path": None,
            "master_path": None,
            "op1_result":  None,
            "op2_result":  None,
        }

        if os.path.exists(MASTER_RATE_PATH):
            self.app_state["master_path"] = MASTER_RATE_PATH

        self._setup_ui()

    def _setup_ui(self):
        central = QWidget()
        central.setStyleSheet(f"background: {THEME['surface']};")
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        self._title_bar = TitleBar()
        root.addWidget(self._title_bar)

        self._progress = ProgressBar()
        root.addWidget(self._progress)

        self._tab_bar = AppTabBar()
        self._tab_bar.set_callback(0, lambda: self._switch_tab(0))
        self._tab_bar.set_callback(1, lambda: self._switch_tab(1))
        self._tab_bar.set_callback(2, lambda: self._switch_tab(2))
        root.addWidget(self._tab_bar)

        self._stack = QStackedWidget()
        self._stack.setStyleSheet(f"background: {THEME['surface']};")
        root.addWidget(self._stack, 1)

        from ui.documents_screen import DocumentsScreen
        from ui.op1_screen import Op1Screen
        from ui.op2_screen import Op2Screen

        self._docs = DocumentsScreen(self.app_state)
        self._op1 = Op1Screen(self.app_state)
        self._op2 = Op2Screen(self.app_state)

        # Stack indices: 0=Documents, 1=Op1, 2=Op2
        self._stack.addWidget(self._docs)
        self._stack.addWidget(self._op1)
        self._stack.addWidget(self._op2)

        # Signals
        self._docs.mrs_loaded.connect(self._on_mrs_loaded)
        self._docs.request_tab_switch.connect(self._switch_tab)

        self._op1.mrs_loaded.connect(self._on_mrs_loaded)
        self._op1.go_to_documents.connect(lambda: self._switch_tab(0))
        self._op1.op1_complete.connect(self._on_op1_complete)

        self._op2.op2_complete.connect(self._on_op2_complete)
        self._op2.back_to_op1.connect(lambda: self._switch_tab(1))

        self._update_mrs_badge()
        self._progress.set_progress(0.10)

    def _on_mrs_loaded(self, path: str):
        self.app_state["master_path"] = path
        self._update_mrs_badge()

    def _update_mrs_badge(self):
        mp = self.app_state.get("master_path")
        if mp and os.path.exists(mp):
            self._title_bar.set_mrs_loaded(os.path.basename(mp))
        else:
            self._title_bar.set_mrs_unloaded()

    def _switch_tab(self, idx: int):
        self._tab_bar.switch_to(idx)
        self._stack.setCurrentIndex(idx)

        # Update progress bar
        if idx == 0:
            self._progress.set_progress(0.10)
        elif idx == 1:
            op1_done = self.app_state.get("op1_result") is not None
            self._progress.set_progress(0.60 if op1_done else 0.30)
        else:
            self._progress.set_progress(0.70)

        # Trigger refresh on the shown screen
        if idx == 0:
            self._docs.refresh()
        elif idx == 2:
            self._op2.refresh()

    def _on_op1_complete(self, result: dict):
        self.app_state["op1_result"] = result
        self._tab_bar.mark_op1_done()
        self._switch_tab(2)

    def _on_op2_complete(self):
        self._tab_bar.mark_op2_done()
        self._progress.set_progress(1.0)
        QMessageBox.information(
            self,
            "Session Complete",
            "Batch invoices generated.\n"
            "Invoice numbers have been incremented.\n\n"
            "Upload a new combined Excel to start the next batch.",
        )
        self.app_state["batch_path"] = None
        self.app_state["batch_session_path"] = None
        self.app_state["op1_result"] = None
        self.app_state["op2_result"] = None
        self._tab_bar.reset_done_marks()
        self._switch_tab(1)
        self._op1.reset()
        self._progress.set_progress(0.30)


def launch():
    app = QApplication(sys.argv)
    app.setFont(QFont(FONT_FAMILY, 10))
    app.setStyle("Fusion")

    from PyQt5.QtGui import QPalette, QColor
    palette = QPalette()
    palette.setColor(QPalette.Window,          QColor(THEME["surface"]))
    palette.setColor(QPalette.WindowText,      QColor(THEME["text_primary"]))
    palette.setColor(QPalette.Base,            QColor("#FFFFFF"))
    palette.setColor(QPalette.AlternateBase,   QColor(THEME["primary_pale"]))
    palette.setColor(QPalette.Button,          QColor("#FFFFFF"))
    palette.setColor(QPalette.ButtonText,      QColor(THEME["text_primary"]))
    palette.setColor(QPalette.Highlight,       QColor(THEME["primary_pale"]))
    palette.setColor(QPalette.HighlightedText, QColor(THEME["primary"]))
    app.setPalette(palette)

    win = MainWindow()
    win.show()
    sys.exit(app.exec_())
