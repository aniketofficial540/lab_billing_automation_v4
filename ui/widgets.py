"""
ui/widgets.py — Shared reusable PyQt5 widgets for Amrita Lab Billing Tool.
All colours imported from config.THEME. Nothing hardcoded in UI code.
"""

import os
from PyQt5.QtWidgets import (
    QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
    QFrame, QLineEdit, QSizePolicy, QScrollArea,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QAbstractItemView, QTreeWidget, QTreeWidgetItem,
    QFileDialog, QSplitter,
)
from PyQt5.QtCore import Qt, pyqtSignal, QSize
from PyQt5.QtGui import QFont, QColor, QPainter, QPainterPath, QPen

from config import THEME, FONT_FAMILY


# ── Helpers ────────────────────────────────────────────────────────────────────

def _font(size: int = 10, bold: bool = False) -> QFont:
    f = QFont(FONT_FAMILY, size)
    f.setBold(bold)
    return f


def _purge_layout(layout):
    while layout.count():
        item = layout.takeAt(0)
        if item.widget():
            item.widget().deleteLater()
        elif item.layout():
            _purge_layout(item.layout())


# ── Shared Stylesheets ─────────────────────────────────────────────────────────

COMBO_STYLE = f"""
    QComboBox {{
        background: white;
        border: 1.5px solid {THEME['border']};
        border-radius: 8px;
        padding: 5px 10px;
        color: {THEME['text_primary']};
        min-height: 36px;
        font-size: 10pt;
    }}
    QComboBox:hover {{ border-color: {THEME['primary_mid']}; }}
    QComboBox:focus {{ border-color: {THEME['primary']}; outline: none; }}
    QComboBox::drop-down {{ border: none; width: 28px; }}
    QComboBox::down-arrow {{
        border-left: 5px solid transparent;
        border-right: 5px solid transparent;
        border-top: 6px solid {THEME['primary']};
        width: 0; height: 0; margin-right: 10px;
    }}
    QComboBox QAbstractItemView {{
        background: white;
        border: 1px solid {THEME['border']};
        outline: none;
        selection-background-color: {THEME['primary_pale']};
        selection-color: {THEME['primary']};
        color: {THEME['text_primary']};
    }}
    QComboBox QAbstractItemView::item {{
        padding: 7px 14px; min-height: 28px;
        color: {THEME['text_primary']}; border: none;
    }}
    QComboBox QAbstractItemView::item:hover {{
        background: {THEME['primary_pale']}; color: {THEME['primary']};
    }}
"""

INPUT_STYLE = (
    f"QLineEdit {{ background: white; border: 1.5px solid {THEME['border']}; "
    f"border-radius: 6px; padding: 4px 8px; color: {THEME['text_primary']}; font-size: 9pt; }}"
    f"QLineEdit:hover {{ border-color: {THEME['primary_mid']}; }}"
    f"QLineEdit:focus {{ border-color: {THEME['primary']}; outline: none; }}"
)
INPUT_STYLE_ERROR = (
    f"QLineEdit {{ background: {THEME['error_bg']}; border: 1.5px solid {THEME['error']}; "
    f"border-radius: 6px; padding: 4px 8px; color: {THEME['text_primary']}; font-size: 9pt; }}"
    f"QLineEdit:focus {{ border-color: {THEME['error']}; outline: none; }}"
)
INPUT_READONLY_STYLE = (
    f"QLineEdit {{ background: {THEME['primary_light']}; border: 1px solid {THEME['border']}; "
    f"border-radius: 6px; padding: 4px 8px; color: {THEME['primary']}; font-size: 9pt; }}"
)


# ── FileDropZone ───────────────────────────────────────────────────────────────

class FileDropZone(QWidget):
    """Clickable / drag-and-drop file picker. Loaded/unloaded states."""

    file_selected = pyqtSignal(str)

    def __init__(self, label: str = "Drop file here or click to browse",
                 file_filter: str = "All Files (*.*)", parent=None):
        super().__init__(parent)
        self._label_text = label
        self._filter = file_filter
        self._path = None
        self._setup_ui()
        self.setAcceptDrops(True)

    def _setup_ui(self):
        self.setMinimumHeight(78)
        self.setCursor(Qt.PointingHandCursor)
        self._apply_style("empty")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        inner = QWidget()
        inner.setStyleSheet("background: transparent; border: none;")
        inner_layout = QVBoxLayout(inner)
        inner_layout.setContentsMargins(16, 12, 16, 12)
        inner_layout.setSpacing(3)
        inner_layout.setAlignment(Qt.AlignCenter)

        self._icon_lbl = QLabel("📂")
        self._icon_lbl.setFont(_font(18))
        self._icon_lbl.setAlignment(Qt.AlignCenter)
        self._icon_lbl.setStyleSheet("background: transparent; border: none;")
        inner_layout.addWidget(self._icon_lbl)

        self._title_lbl = QLabel(self._label_text)
        self._title_lbl.setFont(_font(10, bold=True))
        self._title_lbl.setAlignment(Qt.AlignCenter)
        self._title_lbl.setStyleSheet(f"color: {THEME['text_secondary']}; background: transparent; border: none;")
        inner_layout.addWidget(self._title_lbl)

        self._sub_lbl = QLabel("Click to browse or drag and drop")
        self._sub_lbl.setFont(_font(8))
        self._sub_lbl.setAlignment(Qt.AlignCenter)
        self._sub_lbl.setStyleSheet(f"color: {THEME['text_muted']}; background: transparent; border: none;")
        inner_layout.addWidget(self._sub_lbl)

        layout.addWidget(inner)

    def _apply_style(self, state: str):
        if state == "empty":
            self.setStyleSheet(
                f"FileDropZone {{ background: {THEME['surface']}; "
                f"border: 1.5px dashed {THEME['border']}; border-radius: 8px; }}"
            )
        elif state == "hover":
            self.setStyleSheet(
                "FileDropZone { background: #F0F6FE; "
                f"border: 1.5px dashed {THEME['primary_mid']}; border-radius: 8px; }}"
            )
        elif state == "loaded":
            self.setStyleSheet(
                "FileDropZone { background: #F0FBF6; "
                "border: 1.5px solid #1D9E75; border-radius: 8px; }"
            )

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            path, _ = QFileDialog.getOpenFileName(self, "Select File", "", self._filter)
            if path:
                self._set_file(path)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            if not self._path:
                self._apply_style("hover")

    def dragLeaveEvent(self, event):
        self._apply_style("loaded" if self._path else "empty")

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            self._set_file(urls[0].toLocalFile())

    def _set_file(self, path: str):
        self._path = path
        name = os.path.basename(path)
        self._apply_style("loaded")
        self._icon_lbl.setText("✅")
        self._title_lbl.setText(f"  {name}")
        self._title_lbl.setStyleSheet("color: #085041; background: transparent; border: none; font-weight: bold;")
        self._sub_lbl.setText("Loaded successfully · click to change")
        self._sub_lbl.setStyleSheet("color: #1D9E75; background: transparent; border: none;")
        self.file_selected.emit(path)

    def force_file(self, path: str):
        """Set file without triggering dialog (used when pre-loading from app_state)."""
        self._set_file(path)

    def clear_file(self):
        self._path = None
        self._apply_style("empty")
        self._icon_lbl.setText("📂")
        self._title_lbl.setText(self._label_text)
        self._title_lbl.setStyleSheet(f"color: {THEME['text_secondary']}; background: transparent; border: none;")
        self._sub_lbl.setText("Click to browse or drag and drop")
        self._sub_lbl.setStyleSheet(f"color: {THEME['text_muted']}; background: transparent; border: none;")

    def enterEvent(self, event):
        if not self._path:
            self._apply_style("hover")

    def leaveEvent(self, event):
        if not self._path:
            self._apply_style("empty")

    def set_large_mode(self):
        """Expand the drop zone: taller, bigger icon and text (used for primary uploads)."""
        self.setMinimumHeight(120)
        self._icon_lbl.setFont(_font(26))
        self._title_lbl.setFont(_font(12, bold=True))

    @property
    def path(self):
        return self._path


# ── StatusBadge ────────────────────────────────────────────────────────────────

class StatusBadge(QLabel):
    """Coloured pill-shaped label."""

    _STYLES = {
        "success": ("#0F4F30", "#9FE1CB"),
        "warning": ("#7B3F00", "#FAEEDA"),
        "error":   (THEME["error_bg"], THEME["error"]),
        "info":    (THEME["primary_light"], THEME["primary"]),
        "muted":   (THEME["surface"], THEME["text_muted"]),
        "navy":    (THEME["primary_pale"], THEME["primary"]),
    }

    def __init__(self, text: str = "", status: str = "info", parent=None):
        super().__init__(text, parent)
        self.setFont(_font(9, bold=True))
        self.setAlignment(Qt.AlignCenter)
        self.setContentsMargins(10, 3, 10, 3)
        self.set_status(status)

    def set_status(self, status: str, text: str = None):
        if text is not None:
            self.setText(text)
        bg, fg = self._STYLES.get(status, self._STYLES["info"])
        self.setStyleSheet(
            f"background: {bg}; color: {fg}; border-radius: 10px; "
            "padding: 3px 10px; border: none;"
        )


# ── AlertBanner ────────────────────────────────────────────────────────────────

class AlertBanner(QFrame):
    """Horizontal alert with coloured left border: info / warning / success."""

    _PALETTE = {
        "info":    ("#E6F1FB", "#042C53", "#1A5FA8"),
        "warning": ("#FAEEDA", "#412402", "#EF9F27"),
        "success": ("#E1F5EE", "#04342C", "#1D9E75"),
        "error":   (THEME["error_bg"], THEME["error"], THEME["error"]),
    }

    def __init__(self, text: str = "", style: str = "info", parent=None):
        super().__init__(parent)
        bg, fg, border = self._PALETTE.get(style, self._PALETTE["info"])
        self.setObjectName("AlertBanner")
        self.setStyleSheet(
            f"QFrame#AlertBanner {{ background: {bg}; border-left: 3px solid {border}; "
            "border-top-right-radius: 8px; border-bottom-right-radius: 8px; "
            "border-top-left-radius: 0; border-bottom-left-radius: 0; }"
        )
        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 8, 12, 8)
        layout.setSpacing(8)
        self._lbl = QLabel(text)
        self._lbl.setFont(_font(10))
        self._lbl.setWordWrap(True)
        self._lbl.setStyleSheet(f"color: {fg}; background: transparent; border: none;")
        self._lbl.setTextFormat(Qt.RichText)
        layout.addWidget(self._lbl)

    def set_text(self, text: str):
        self._lbl.setText(text)


# ── CardFrame ──────────────────────────────────────────────────────────────────

class CardFrame(QFrame):
    """White card with light-blue header bar and white body (matches HTML .card pattern)."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("CardFrame")
        self.setStyleSheet(
            "QFrame#CardFrame { background: #FFFFFF; "
            f"border: 1px solid {THEME['border']}; border-radius: 10px; }}"
        )

        self._outer = QVBoxLayout(self)
        self._outer.setContentsMargins(0, 0, 0, 0)
        self._outer.setSpacing(0)

        self._head = None  # created by set_header
        self._head_title_lbl = None
        self._head_icon_lbl = None

        self._body_widget = QWidget()
        self._body_widget.setObjectName("CBody")
        self._body_widget.setStyleSheet("QWidget#CBody { background: white; border: none; }")
        self._body_layout = QVBoxLayout(self._body_widget)
        self._body_layout.setContentsMargins(14, 14, 14, 14)
        self._body_layout.setSpacing(10)
        self._outer.addWidget(self._body_widget)

    def set_header(self, icon: str, title: str, right_widget=None):
        if self._head:
            self._outer.removeWidget(self._head)
            self._head.deleteLater()

        head = QWidget()
        head.setObjectName("CHead")
        head.setStyleSheet(
            f"QWidget#CHead {{ background: {THEME['primary_pale']}; "
            "border-top-left-radius: 9px; border-top-right-radius: 9px; "
            f"border-bottom: 1px solid {THEME['border']}; }}"
        )
        head.setFixedHeight(40)
        h = QHBoxLayout(head)
        h.setContentsMargins(14, 0, 14, 0)
        h.setSpacing(8)

        if icon:
            ic = QLabel(icon)
            ic.setFont(_font(13))
            ic.setStyleSheet(f"color: {THEME['primary_mid']}; background: transparent; border: none;")
            h.addWidget(ic)
            self._head_icon_lbl = ic
        else:
            self._head_icon_lbl = None

        t = QLabel(title)
        t.setFont(_font(11, bold=True))
        t.setStyleSheet(f"color: {THEME['primary']}; background: transparent; border: none;")
        h.addWidget(t)
        h.addStretch()
        self._head_title_lbl = t

        if right_widget:
            h.addWidget(right_widget)

        self._head = head
        self._outer.insertWidget(0, head)

    @property
    def body(self) -> QVBoxLayout:
        return self._body_layout

    def add_widget(self, w, stretch: int = 0):
        self._body_layout.addWidget(w, stretch)

    def add_layout(self, layout):
        self._body_layout.addLayout(layout)

    def add_spacing(self, n: int):
        self._body_layout.addSpacing(n)

    def set_body_padding(self, left: int, top: int, right: int, bottom: int):
        self._body_layout.setContentsMargins(left, top, right, bottom)

    def set_header_accent(self, color: str):
        """Change header icon + title label to a given colour (e.g. success/warning)."""
        if self._head_title_lbl:
            self._head_title_lbl.setStyleSheet(
                f"color: {color}; background: transparent; border: none;"
            )
        if self._head_icon_lbl:
            self._head_icon_lbl.setStyleSheet(
                f"color: {color}; background: transparent; border: none;"
            )


# StepCard kept as alias so any existing references still work
StepCard = CardFrame


# ── InfoCard ───────────────────────────────────────────────────────────────────

class InfoCard(QFrame):
    """Compact card for label-value rows — used in invoice summary, Op1 summary."""

    def __init__(self, title: str = None, parent=None):
        super().__init__(parent)
        self.setObjectName("InfoCard")
        self.setStyleSheet(
            f"QFrame#InfoCard {{ background: {THEME['surface']}; "
            f"border: 1px solid {THEME['border']}; border-radius: 8px; }}"
        )
        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(14, 10, 14, 12)
        self._layout.setSpacing(5)
        self._fixed_count = 0

        if title:
            t = QLabel(title)
            t.setFont(_font(10, bold=True))
            t.setStyleSheet(f"color: {THEME['primary']}; border: none; background: transparent;")
            self._layout.addWidget(t)
            sep = QFrame()
            sep.setFrameShape(QFrame.HLine)
            sep.setFixedHeight(1)
            sep.setStyleSheet(f"background: {THEME['border']}; border: none;")
            self._layout.addWidget(sep)
            self._fixed_count = 2

    def add_row(self, label: str, value: str, style: str = "normal"):
        row = QHBoxLayout()
        row.setSpacing(8)
        lbl = QLabel(f"{label}:")
        lbl.setFont(_font(9))
        lbl.setFixedWidth(170)
        lbl.setStyleSheet(f"color: {THEME['text_secondary']}; border: none; background: transparent;")
        val = QLabel(str(value))
        val.setFont(_font(9, bold=True))
        val.setWordWrap(True)
        color = {
            "success": THEME["success"],
            "warning": THEME["warning"],
            "blue":    THEME["primary_mid"],
        }.get(style, THEME["text_primary"])
        val.setStyleSheet(f"color: {color}; border: none; background: transparent;")
        row.addWidget(lbl)
        row.addWidget(val, 1)
        self._layout.addLayout(row)

    def clear_rows(self):
        while self._layout.count() > self._fixed_count:
            item = self._layout.takeAt(self._layout.count() - 1)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():
                _purge_layout(item.layout())


# ── StatCard / StatGrid ────────────────────────────────────────────────────────

class StatCard(QFrame):
    """One stat tile: label + large value."""

    def __init__(self, label: str, value: str = "—", value_color: str = None, parent=None):
        super().__init__(parent)
        self.setObjectName("StatCard")
        self.setStyleSheet(
            f"QFrame#StatCard {{ background: {THEME['surface']}; "
            f"border: 1px solid {THEME['border']}; border-radius: 8px; }}"
        )
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 10, 12, 10)
        layout.setSpacing(4)

        self._lbl = QLabel(label)
        self._lbl.setFont(_font(9))
        self._lbl.setStyleSheet(f"color: {THEME['text_secondary']}; background: transparent; border: none;")
        layout.addWidget(self._lbl)

        self._val = QLabel(value)
        self._val.setFont(_font(18, bold=True))
        fg = value_color or THEME["primary"]
        self._val.setStyleSheet(f"color: {fg}; background: transparent; border: none;")
        layout.addWidget(self._val)

    def set_value(self, value: str, color: str = None):
        self._val.setText(value)
        if color:
            self._val.setStyleSheet(f"color: {color}; background: transparent; border: none;")


def make_stat_grid(stats: list) -> QWidget:
    """stats: list of (label, value, color). Returns a QWidget with 4 stat cards in a row."""
    w = QWidget()
    w.setStyleSheet("background: transparent;")
    layout = QHBoxLayout(w)
    layout.setContentsMargins(0, 0, 0, 0)
    layout.setSpacing(10)
    for label, value, color in stats:
        card = StatCard(label, str(value), color)
        layout.addWidget(card)
    return w


# ── DataTable ──────────────────────────────────────────────────────────────────

class DataTable(QTreeWidget):
    """QTreeWidget with navy header, alternating rows, typed-row colouring."""

    def __init__(self, columns: list, parent=None):
        super().__init__(parent)
        self.setColumnCount(len(columns))
        self.setHeaderLabels(columns)
        self.setRootIsDecorated(False)
        self.setAlternatingRowColors(True)
        self.setSelectionMode(QAbstractItemView.SingleSelection)
        self.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.header().setSectionResizeMode(QHeaderView.Interactive)
        self.header().setStretchLastSection(True)
        self.setStyleSheet(f"""
            QTreeWidget {{
                border: 1px solid {THEME['border']};
                border-radius: 6px;
                background: #FFFFFF;
                alternate-background-color: {THEME['primary_pale']};
                outline: none;
                font-size: 9pt;
            }}
            QHeaderView::section {{
                background: {THEME['primary']};
                color: white;
                font-weight: bold;
                font-size: 8pt;
                text-transform: uppercase;
                letter-spacing: 0.5px;
                padding: 6px 8px;
                border: none;
                border-right: 1px solid {THEME['primary_mid']};
            }}
            QHeaderView::section:last {{ border-right: none; }}
            QTreeWidget::item {{ padding: 4px 6px; border: none; color: {THEME['text_primary']}; }}
            QTreeWidget::item:hover {{ background: {THEME['primary_light']}; color: {THEME['text_primary']}; }}
            QTreeWidget::item:selected {{ background: {THEME['primary_pale']}; color: {THEME['primary']}; }}
        """)

    def add_row(self, cells: list, row_type: str = "normal"):
        item = QTreeWidgetItem([str(c) if c is not None else "" for c in cells])
        if row_type == "unmatched":
            for i in range(len(cells)):
                item.setBackground(i, QColor(THEME['warning_bg']))
                item.setForeground(i, QColor(THEME['warning']))
        elif row_type == "total":
            font = _font(9, bold=True)
            for i in range(len(cells)):
                item.setBackground(i, QColor(THEME['total_bg']))
                item.setForeground(i, QColor(THEME['primary']))
                item.setFont(i, font)
        elif row_type == "separator":
            for i in range(len(cells)):
                item.setBackground(i, QColor(THEME['separator_bg']))
        self.addTopLevelItem(item)

    def load_rows(self, rows: list, row_types: list = None):
        self.clear()
        for i, row in enumerate(rows):
            t = (row_types[i] if row_types and i < len(row_types) else "normal")
            self.add_row(row, t)

    def load_dataframe(self, df, max_rows: int = 0):
        self.clear()
        for i, (_, row) in enumerate(df.iterrows()):
            if max_rows and i >= max_rows:
                break
            self.add_row([str(v) if v is not None else "" for v in row.values])


# ── UnmatchedRatesForm ─────────────────────────────────────────────────────────

class _UnmatchedRateRow(QWidget):
    changed = pyqtSignal()

    def __init__(self, cpt_code: str, service_name: str, parent=None):
        super().__init__(parent)
        self.cpt_code = cpt_code
        self.setStyleSheet(f"background: white; border-bottom: 1px solid {THEME['border']};")

        layout = QHBoxLayout(self)
        layout.setContentsMargins(10, 7, 10, 7)
        layout.setSpacing(8)

        # CPT amber pill
        cpt = QLabel(cpt_code)
        cpt.setFont(_font(9, bold=True))
        cpt.setFixedWidth(114)
        cpt.setAlignment(Qt.AlignCenter)
        cpt.setStyleSheet(
            f"color: {THEME['warning']}; background: {THEME['warning_bg']}; "
            "border-radius: 4px; padding: 2px 6px; font-family: Consolas, monospace; border: none;"
        )
        layout.addWidget(cpt)

        # Service name
        svc = QLabel(service_name)
        svc.setFont(_font(9))
        svc.setWordWrap(True)
        svc.setMinimumWidth(130)
        svc.setStyleSheet(f"color: {THEME['text_primary']}; background: transparent; border: none;")
        layout.addWidget(svc, 1)

        def _inp(placeholder: str) -> QLineEdit:
            e = QLineEdit()
            e.setPlaceholderText(placeholder)
            e.setFixedWidth(88)
            e.setFont(_font(9))
            e.setStyleSheet(INPUT_STYLE)
            return e

        self._tariff = _inp("Tariff ₹")
        self._tariff.textChanged.connect(self._recalc)
        layout.addWidget(self._tariff)

        self._bill = _inp("Bill ₹")
        self._bill.textChanged.connect(self._recalc)
        layout.addWidget(self._bill)

        # Discount — read-only blue
        self._discount = QLabel("—")
        self._discount.setFixedWidth(88)
        self._discount.setFont(_font(9))
        self._discount.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self._discount.setStyleSheet(
            f"background: {THEME['primary_light']}; color: {THEME['primary']}; "
            "border-radius: 5px; padding: 2px 8px; border: none;"
        )
        layout.addWidget(self._discount)

    def _recalc(self):
        try:
            t = self._tariff.text().strip()
            b = self._bill.text().strip()
            if t and b:
                tariff = float(t)
                bill = float(b)
                self._discount.setText(f"₹{tariff - bill:.2f}")
                if bill > tariff:
                    self._bill.setStyleSheet(INPUT_STYLE_ERROR)
                else:
                    self._bill.setStyleSheet(INPUT_STYLE)
            else:
                self._discount.setText("—")
                self._bill.setStyleSheet(INPUT_STYLE)
        except ValueError:
            self._discount.setText("—")
        self.changed.emit()

    def is_filled(self) -> bool:
        try:
            float(self._tariff.text().strip())
            float(self._bill.text().strip())
            return bool(self._tariff.text().strip() and self._bill.text().strip())
        except ValueError:
            return False

    def get_rate(self) -> dict | None:
        if not self.is_filled():
            return None
        try:
            mrp = float(self._tariff.text().strip())
            bill = float(self._bill.text().strip())
            return {
                "cpt_code": self.cpt_code,
                "mrp": self._tariff.text().strip(),
                "discount": f"{mrp - bill:.2f}",
                "bill_amount": self._bill.text().strip(),
            }
        except ValueError:
            return None


class UnmatchedRatesForm(QWidget):
    """Table with CPT + Service + Tariff input + Bill input + auto-calc Discount."""

    all_filled_changed = pyqtSignal(bool)

    def __init__(self, unmatched: list, parent=None):
        # unmatched: list of (cpt_code, service_name)
        super().__init__(parent)
        self._rows: list[_UnmatchedRateRow] = []
        self._setup_ui(unmatched)

    def _setup_ui(self, unmatched: list):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        wrapper = QFrame()
        wrapper.setObjectName("UMWrap")
        wrapper.setStyleSheet(
            "QFrame#UMWrap { border: 1px solid #F0C050; border-radius: 6px; background: white; }"
        )
        inner = QVBoxLayout(wrapper)
        inner.setContentsMargins(0, 0, 0, 0)
        inner.setSpacing(0)

        # Header
        hdr = QWidget()
        hdr.setStyleSheet(
            f"background: {THEME['warning_bg']}; border-top-left-radius: 5px; "
            "border-top-right-radius: 5px; border-bottom: 1px solid #F0C050;"
        )
        hdr_layout = QHBoxLayout(hdr)
        hdr_layout.setContentsMargins(10, 6, 10, 6)
        hdr_layout.setSpacing(8)

        col_specs = [
            ("CPT Code", 114, False),
            ("Service Name", -1, True),
            ("Amrita Tariff ₹", 88, False),
            ("Bill Amount ₹", 88, False),
            ("Discount ₹ (auto)", 88, False),
        ]
        for col_name, width, stretch in col_specs:
            lbl = QLabel(col_name)
            lbl.setFont(_font(8, bold=True))
            lbl.setStyleSheet(f"color: {THEME['warning']}; background: transparent; border: none;")
            if not stretch:
                lbl.setFixedWidth(width)
            else:
                lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            hdr_layout.addWidget(lbl)

        inner.addWidget(hdr)

        # Data rows
        for cpt, svc in unmatched:
            row = _UnmatchedRateRow(cpt, svc)
            row.changed.connect(self._on_changed)
            self._rows.append(row)
            inner.addWidget(row)

        outer.addWidget(wrapper)

    def _on_changed(self):
        self.all_filled_changed.emit(self.all_filled())

    def all_filled(self) -> bool:
        return all(r.is_filled() for r in self._rows)

    def get_confirmed_rates(self) -> list:
        return [r.get_rate() for r in self._rows if r.get_rate() is not None]


# ── AddendumEntryTable ─────────────────────────────────────────────────────────

class AddendumEntryTable(QWidget):
    """Scrollable table of all party tests. Employee enters New Discounted Price for changes."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._tests: list[tuple] = []
        self._inputs: list[QLineEdit] = []
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        # Search bar row
        search_row = QHBoxLayout()
        self._search = QLineEdit()
        self._search.setPlaceholderText("Search by test name or code...")
        self._search.setFont(_font(9))
        self._search.setStyleSheet(INPUT_STYLE)
        self._search.textChanged.connect(self._filter_rows)
        search_row.addWidget(self._search, 1)

        self._counter_lbl = QLabel("0 tests updated")
        self._counter_lbl.setFont(_font(9, bold=True))
        self._counter_lbl.setStyleSheet(f"color: {THEME['primary_mid']}; background: transparent; border: none;")
        search_row.addWidget(self._counter_lbl)
        layout.addLayout(search_row)

        # Table
        self._table = QTableWidget(0, 4)
        self._table.setHorizontalHeaderLabels([
            "Test Code", "Test Name", "Current Bill Amount", "New Discounted Price"
        ])
        hdr = self._table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(1, QHeaderView.Stretch)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self._table.setAlternatingRowColors(True)
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._table.verticalHeader().setVisible(False)
        self._table.setMinimumHeight(200)
        self._table.setStyleSheet(f"""
            QTableWidget {{
                border: 1px solid {THEME['border']}; border-radius: 6px;
                background: white; alternate-background-color: {THEME['primary_pale']};
                gridline-color: {THEME['border']}; outline: none; font-size: 9pt;
            }}
            QHeaderView::section {{
                background: {THEME['primary_pale']}; color: {THEME['primary']};
                font-weight: bold; font-size: 8pt; padding: 6px 8px;
                border: none; border-bottom: 1px solid {THEME['border']};
                border-right: 1px solid {THEME['border']};
            }}
            QTableWidget::item {{
                padding: 4px 6px; color: {THEME['text_primary']}; border: none;
            }}
            QTableWidget::item:selected {{
                background: {THEME['primary_pale']}; color: {THEME['primary']};
            }}
        """)
        layout.addWidget(self._table)

    def load_tests(self, tests: list):
        """tests: list of (code, name, current_bill_amount)"""
        self._tests = tests
        self._inputs = []
        self._table.setRowCount(0)

        for code, name, price in tests:
            r = self._table.rowCount()
            self._table.insertRow(r)

            i0 = QTableWidgetItem(code)
            i0.setForeground(QColor(THEME['warning']))
            i0.setBackground(QColor(THEME['warning_bg']))
            self._table.setItem(r, 0, i0)

            i1 = QTableWidgetItem(name)
            self._table.setItem(r, 1, i1)

            i2 = QTableWidgetItem(str(price))
            i2.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            i2.setForeground(QColor(THEME['text_secondary']))
            self._table.setItem(r, 2, i2)

            inp = QLineEdit()
            inp.setPlaceholderText("New price ₹")
            inp.setFont(_font(9))
            inp.setStyleSheet(INPUT_STYLE)
            inp.textChanged.connect(self._update_counter)
            self._table.setCellWidget(r, 3, inp)
            self._inputs.append(inp)

        self._update_counter()

    def _filter_rows(self, text: str):
        text = text.lower().strip()
        for r in range(self._table.rowCount()):
            code = (self._table.item(r, 0) or QTableWidgetItem()).text().lower()
            name = (self._table.item(r, 1) or QTableWidgetItem()).text().lower()
            self._table.setRowHidden(r, bool(text) and text not in code and text not in name)

    def _update_counter(self):
        count = sum(1 for inp in self._inputs if inp.text().strip())
        self._counter_lbl.setText(f"{count} test{'s' if count != 1 else ''} updated")

    def get_updated_rows(self) -> list:
        results = []
        for i, (code, name, _) in enumerate(self._tests):
            if i < len(self._inputs):
                price = self._inputs[i].text().strip()
                if price:
                    results.append({
                        "test_code": code,
                        "test_name": name,
                        "discounted_price": price,
                    })
        return results


# ── ChecklistPanel ─────────────────────────────────────────────────────────────

class ChecklistPanel(QWidget):
    """Verification checklist with circle icons."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(0, 0, 0, 0)
        self._layout.setSpacing(0)

    def clear(self):
        while self._layout.count():
            item = self._layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def add_item(self, label: str, status: str = "pass", detail: str = ""):
        if self._layout.count() > 0:
            sep = QFrame()
            sep.setFrameShape(QFrame.HLine)
            sep.setFixedHeight(1)
            sep.setStyleSheet(f"background: {THEME['border']}; border: none;")
            self._layout.addWidget(sep)

        item_w = QWidget()
        item_w.setStyleSheet("background: transparent;")
        h = QHBoxLayout(item_w)
        h.setContentsMargins(4, 9, 4, 9)
        h.setSpacing(12)

        if status == "pass":
            cir_bg, cir_fg, icon = THEME['success_bg'], THEME['success'], "✓"
        elif status == "warn":
            cir_bg, cir_fg, icon = THEME['warning_bg'], THEME['warning'], "!"
        else:
            cir_bg, cir_fg, icon = THEME['error_bg'], THEME['error'], "✕"

        circle = QLabel(icon)
        circle.setFixedSize(22, 22)
        circle.setAlignment(Qt.AlignCenter)
        circle.setFont(_font(9, bold=True))
        circle.setStyleSheet(
            f"background: {cir_bg}; color: {cir_fg}; border-radius: 11px; border: none;"
        )
        h.addWidget(circle)

        text_col = QWidget()
        text_col.setStyleSheet("background: transparent;")
        tv = QVBoxLayout(text_col)
        tv.setContentsMargins(0, 0, 0, 0)
        tv.setSpacing(2)

        main_lbl = QLabel(label)
        main_lbl.setFont(_font(10))
        main_lbl.setStyleSheet(f"color: {THEME['text_primary']}; background: transparent; border: none;")
        main_lbl.setWordWrap(True)
        tv.addWidget(main_lbl)

        if detail:
            det_lbl = QLabel(detail)
            det_lbl.setFont(_font(9))
            det_lbl.setStyleSheet(f"color: {THEME['text_muted']}; background: transparent; border: none;")
            det_lbl.setWordWrap(True)
            tv.addWidget(det_lbl)

        h.addWidget(text_col, 1)
        self._layout.addWidget(item_w)


# ── InvoicePreview ─────────────────────────────────────────────────────────────

class InvoicePreview(QFrame):
    """Read-only invoice preview panel."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("InvPreview")
        self.setStyleSheet(
            "QFrame#InvPreview { background: white; "
            f"border: 1px solid {THEME['border']}; border-radius: 10px; }}"
        )
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(12)

        # Hospital header (centered)
        hdr = QWidget()
        hdr.setStyleSheet("background: transparent;")
        hdr_v = QVBoxLayout(hdr)
        hdr_v.setContentsMargins(0, 0, 0, 8)
        hdr_v.setSpacing(3)
        hdr_v.setAlignment(Qt.AlignCenter)

        hosp = QLabel("AMRITA INSTITUTE OF MEDICAL SCIENCES & RESEARCH CENTRE")
        hosp.setFont(_font(11, bold=True))
        hosp.setAlignment(Qt.AlignCenter)
        hosp.setWordWrap(True)
        hosp.setStyleSheet(f"color: {THEME['primary']}; background: transparent; border: none;")
        hdr_v.addWidget(hosp)

        addr = QLabel("Mata Amritanandamayi Marg, Sector-88, Faridabad, Haryana – 121002")
        addr.setFont(_font(9))
        addr.setAlignment(Qt.AlignCenter)
        addr.setStyleSheet(f"color: {THEME['text_muted']}; background: transparent; border: none;")
        hdr_v.addWidget(addr)

        stamp = QLabel("I  N  V  O  I  C  E")
        stamp.setFont(_font(10, bold=True))
        stamp.setAlignment(Qt.AlignCenter)
        stamp.setStyleSheet(f"color: {THEME['primary_mid']}; background: transparent; border: none;")
        hdr_v.addWidget(stamp)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFixedHeight(1)
        sep.setStyleSheet(f"background: {THEME['border']}; border: none;")
        hdr_v.addWidget(sep)

        layout.addWidget(hdr)

        # 2×3 fields grid
        from PyQt5.QtWidgets import QGridLayout
        fields_w = QWidget()
        fields_w.setStyleSheet("background: transparent;")
        grid = QGridLayout(fields_w)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(20)
        grid.setVerticalSpacing(12)

        field_names = ["Invoice No.", "Invoice Date", "To", "Address", "Service", "Payment Terms"]
        self._fv: dict[str, QLabel] = {}
        for i, fname in enumerate(field_names):
            r, c = i // 2, i % 2
            fw = QWidget()
            fw.setStyleSheet("background: transparent;")
            fv = QVBoxLayout(fw)
            fv.setContentsMargins(0, 0, 0, 0)
            fv.setSpacing(2)
            lbl = QLabel(fname.upper())
            lbl.setFont(_font(8))
            lbl.setStyleSheet(f"color: {THEME['text_muted']}; background: transparent; border: none;")
            val = QLabel("—")
            val.setFont(_font(11, bold=True))
            val.setWordWrap(True)
            val.setStyleSheet(f"color: {THEME['text_primary']}; background: transparent; border: none;")
            fv.addWidget(lbl)
            fv.addWidget(val)
            grid.addWidget(fw, r, c)
            self._fv[fname] = val

        # Invoice No. stands out in navy
        self._fv["Invoice No."].setStyleSheet(
            f"color: {THEME['primary']}; font-weight: bold; font-size: 13pt; "
            "background: transparent; border: none;"
        )
        layout.addWidget(fields_w)

        # Amount box
        amt_box = QFrame()
        amt_box.setObjectName("AmtBox")
        amt_box.setStyleSheet(
            f"QFrame#AmtBox {{ background: {THEME['primary_light']}; border-radius: 8px; }}"
        )
        ab = QHBoxLayout(amt_box)
        ab.setContentsMargins(14, 12, 14, 12)
        ab.addWidget(QLabel("Total Amount Due"))
        ab.itemAt(0).widget().setFont(_font(11, bold=True))
        ab.itemAt(0).widget().setStyleSheet(
            f"color: {THEME['primary']}; background: transparent; border: none;"
        )
        ab.addStretch()
        self._amt_val = QLabel("₹0.00")
        self._amt_val.setFont(_font(18, bold=True))
        self._amt_val.setStyleSheet(
            f"color: {THEME['primary']}; background: transparent; border: none;"
        )
        ab.addWidget(self._amt_val)
        layout.addWidget(amt_box)

        # Amount in words
        self._words = QLabel("")
        self._words.setFont(_font(9))
        self._words.setWordWrap(True)
        self._words.setStyleSheet(
            f"color: {THEME['text_muted']}; font-style: italic; background: transparent; border: none;"
        )
        layout.addWidget(self._words)

        # Bank details
        bank_box = QFrame()
        bank_box.setObjectName("BankBox")
        bank_box.setStyleSheet(
            f"QFrame#BankBox {{ background: {THEME['surface']}; border-radius: 6px; }}"
        )
        bb = QVBoxLayout(bank_box)
        bb.setContentsMargins(10, 8, 10, 8)
        from config import BANK_NAME, BANK_ACCOUNT, BANK_IFSC, BANK_MICR, BANK_FAVOUR
        bank_lbl = QLabel(
            f"Bank: {BANK_NAME}  ·  A/C: {BANK_ACCOUNT}  ·  "
            f"IFSC: {BANK_IFSC}  ·  MICR: {BANK_MICR}\n"
            f"In favour of: {BANK_FAVOUR}"
        )
        bank_lbl.setFont(_font(9))
        bank_lbl.setWordWrap(True)
        bank_lbl.setStyleSheet(
            f"color: {THEME['text_muted']}; background: transparent; border: none;"
        )
        bb.addWidget(bank_lbl)
        layout.addWidget(bank_box)

    def refresh(self, data: dict):
        self._fv["Invoice No."].setText(data.get("invoice_no", "—"))
        self._fv["Invoice Date"].setText(data.get("invoice_date_str", "—"))
        self._fv["To"].setText(data.get("party_full", "—"))
        self._fv["Address"].setText(data.get("party_address", "—"))
        bp = data.get("billing_period", "")
        self._fv["Service"].setText(f"Lab Test Service – {bp}")
        self._fv["Payment Terms"].setText("30 days from receipt")
        total = data.get("total", 0)
        self._amt_val.setText(f"₹{total:,.2f}")
        self._words.setText(data.get("amount_words", ""))


# ── PrimaryButton ──────────────────────────────────────────────────────────────

class PrimaryButton(QPushButton):
    """Flat filled button. variant='navy' (default) or 'success' for green."""

    _VARIANTS = {
        "navy":    (THEME["primary"],   THEME["primary_mid"], "#FFFFFF"),
        "success": (THEME["success"],   "#1a6e42",            "#FFFFFF"),
        "ghost":   ("#FFFFFF",          THEME["primary_pale"], THEME["primary_mid"]),
    }

    def __init__(self, text: str, variant: str = "navy", parent=None):
        super().__init__(text, parent)
        self._variant = variant
        self.setFont(_font(10, bold=True))
        self.setMinimumHeight(38)
        self.setCursor(Qt.PointingHandCursor)
        self._apply(enabled=True)

    def _apply(self, enabled: bool):
        bg, hover, fg = self._VARIANTS.get(self._variant, self._VARIANTS["navy"])
        if not enabled:
            self.setStyleSheet(
                f"QPushButton {{ background: {THEME['border']}; color: {THEME['text_muted']}; "
                "border: none; border-radius: 8px; padding: 5px 20px; }"
            )
        elif self._variant == "ghost":
            self.setStyleSheet(
                f"QPushButton {{ background: transparent; color: {THEME['primary_mid']}; "
                f"border: 1px solid #85B7EB; border-radius: 8px; padding: 5px 20px; }}"
                f"QPushButton:hover {{ background: {THEME['primary_pale']}; }}"
                f"QPushButton:pressed {{ background: {THEME['primary_light']}; }}"
            )
        else:
            self.setStyleSheet(
                f"QPushButton {{ background: {bg}; color: {fg}; "
                "border: none; border-radius: 8px; padding: 5px 20px; }"
                f"QPushButton:hover {{ background: {hover}; }}"
                f"QPushButton:pressed {{ background: {hover}; }}"
                f"QPushButton:disabled {{ background: {THEME['border']}; color: {THEME['text_muted']}; }}"
            )

    def setEnabled(self, enabled: bool):
        super().setEnabled(enabled)
        self._apply(enabled)


# ── SecondaryButton ────────────────────────────────────────────────────────────

class SecondaryButton(QPushButton):
    """Surface bg, bordered button."""

    def __init__(self, text: str, parent=None):
        super().__init__(text, parent)
        self.setFont(_font(10))
        self.setMinimumHeight(38)
        self.setCursor(Qt.PointingHandCursor)
        self.setStyleSheet(
            f"QPushButton {{ background: white; color: {THEME['primary']}; "
            f"border: 1px solid {THEME['primary']}; border-radius: 8px; padding: 5px 18px; }}"
            f"QPushButton:hover {{ background: {THEME['primary_pale']}; }}"
            f"QPushButton:pressed {{ background: {THEME['primary']}; color: white; }}"
        )


# ── ToggleSwitch ───────────────────────────────────────────────────────────────

class ToggleSwitch(QWidget):
    """Custom-drawn 44×24 toggle switch."""

    toggled = pyqtSignal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._on = False
        self.setFixedSize(44, 24)
        self.setCursor(Qt.PointingHandCursor)

    def paintEvent(self, event):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)

        # Track
        track_color = QColor(THEME['primary'] if self._on else THEME['border'])
        p.setBrush(track_color)
        p.setPen(Qt.NoPen)
        p.drawRoundedRect(0, 0, self.width(), self.height(), 12, 12)

        # Thumb
        thumb_size = 18
        thumb_y = (self.height() - thumb_size) // 2
        thumb_x = self.width() - thumb_size - 3 if self._on else 3
        p.setBrush(QColor("#FFFFFF"))
        p.drawEllipse(thumb_x, thumb_y, thumb_size, thumb_size)
        p.end()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._on = not self._on
            self.update()
            self.toggled.emit(self._on)

    @property
    def is_on(self) -> bool:
        return self._on

    def set_on(self, value: bool):
        if self._on != value:
            self._on = value
            self.update()
