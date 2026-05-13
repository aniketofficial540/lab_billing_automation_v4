"""
ui/documents_screen.py — Documents & Settings screen.
Manages Master Rate Sheet and per-party addendum JSON storage.
"""

import os
import shutil
from datetime import date, datetime
import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QScrollArea, QFrame, QSizePolicy, QMessageBox,
    QDialog, QTableWidget, QTableWidgetItem, QHeaderView,
    QAbstractItemView, QDateEdit, QComboBox, QFileDialog,
    QGridLayout, QLineEdit,
)
from PyQt5.QtCore import Qt, pyqtSignal, QDate
from PyQt5.QtGui import QColor

from config import (
    THEME, MASTER_RATE_PATH,
    PC_SHEET_NAME, PC_HIS_NAME, PC_FULL_NAME, PC_SHORT,
    MR_SHEET_NAME, MR_HEADER_ROW, COL_TEST_CODE, COL_TEST_NAME, COL_MRP,
)
from ui.widgets import (
    FileDropZone, AlertBanner, CardFrame,
    PrimaryButton, SecondaryButton, _font, _purge_layout, COMBO_STYLE,
)


# ── Section header ─────────────────────────────────────────────────────────────

def _section_header(title: str, subtitle: str) -> QWidget:
    w = QWidget()
    w.setStyleSheet("background: transparent;")
    v = QVBoxLayout(w)
    v.setContentsMargins(0, 0, 0, 4)
    v.setSpacing(3)
    t = QLabel(title)
    t.setFont(_font(15, bold=True))
    t.setStyleSheet(f"color: {THEME['primary']}; background: transparent; border: none;")
    s = QLabel(subtitle)
    s.setFont(_font(9))
    s.setStyleSheet(f"color: {THEME['text_secondary']}; background: transparent; border: none;")
    v.addWidget(t)
    v.addWidget(s)
    return w


def _table_style() -> str:
    return f"""
        QTableWidget {{
            border: 1px solid {THEME['border']}; border-radius: 6px;
            background: white; alternate-background-color: {THEME['primary_pale']};
            gridline-color: {THEME['border']}; font-size: 9pt; outline: none;
        }}
        QHeaderView::section {{
            background: {THEME['primary']}; color: white;
            font-weight: bold; font-size: 8pt;
            padding: 6px 8px; border: none;
            border-right: 1px solid {THEME['primary_mid']};
        }}
        QHeaderView::section:last {{ border-right: none; }}
        QTableWidget::item {{ padding: 4px 6px; color: {THEME['text_primary']}; border: none; }}
        QTableWidget::item:selected {{
            background: {THEME['primary_pale']}; color: {THEME['primary']};
        }}
    """


# ── Addendum row widget ────────────────────────────────────────────────────────

class _AddendumRow(QFrame):
    view_clicked   = pyqtSignal(str)
    delete_clicked = pyqtSignal(str)

    def __init__(self, short_code: str, full_name: str, eff_date: str,
                 test_count: int, uploaded: str, parent=None):
        super().__init__(parent)
        self._sc = short_code
        self.setStyleSheet(
            f"QFrame {{ background: white; border-bottom: 1px solid {THEME['border']}; }}"
        )
        h = QHBoxLayout(self)
        h.setContentsMargins(12, 8, 12, 8)
        h.setSpacing(12)

        name_lbl = QLabel(full_name or short_code)
        name_lbl.setFont(_font(9, bold=True))
        name_lbl.setStyleSheet(f"color: {THEME['text_primary']}; background: transparent; border: none;")
        name_lbl.setMinimumWidth(180)
        h.addWidget(name_lbl)

        sc_lbl = QLabel(short_code)
        sc_lbl.setFont(_font(8))
        sc_lbl.setFixedWidth(80)
        sc_lbl.setStyleSheet(
            f"background: {THEME['primary_pale']}; color: {THEME['primary']}; "
            "border-radius: 4px; padding: 2px 6px; border: none;"
        )
        h.addWidget(sc_lbl)

        eff_lbl = QLabel(eff_date)
        eff_lbl.setFont(_font(9))
        eff_lbl.setFixedWidth(100)
        eff_lbl.setStyleSheet(f"color: {THEME['text_secondary']}; background: transparent; border: none;")
        h.addWidget(eff_lbl)

        cnt_lbl = QLabel(f"{test_count} tests")
        cnt_lbl.setFont(_font(9))
        cnt_lbl.setFixedWidth(70)
        cnt_lbl.setStyleSheet(f"color: {THEME['text_secondary']}; background: transparent; border: none;")
        h.addWidget(cnt_lbl)

        up_lbl = QLabel(uploaded)
        up_lbl.setFont(_font(8))
        up_lbl.setFixedWidth(90)
        up_lbl.setStyleSheet(f"color: {THEME['text_muted']}; background: transparent; border: none;")
        h.addWidget(up_lbl)

        h.addStretch()

        view_btn = PrimaryButton("View", variant="ghost")
        view_btn.setFixedHeight(28)
        view_btn.clicked.connect(lambda: self.view_clicked.emit(self._sc))
        h.addWidget(view_btn)

        del_btn = SecondaryButton("Delete")
        del_btn.setFixedHeight(28)
        del_btn.clicked.connect(lambda: self.delete_clicked.emit(self._sc))
        h.addWidget(del_btn)


def _addendum_table_header() -> QWidget:
    hdr = QWidget()
    hdr.setStyleSheet(
        f"background: {THEME['primary_pale']}; border-bottom: 1px solid {THEME['border']};"
    )
    h = QHBoxLayout(hdr)
    h.setContentsMargins(12, 6, 12, 6)
    h.setSpacing(12)
    for label, width in [
        ("Party Name", 180), ("Short Code", 80), ("Effective From", 100),
        ("Tests", 70), ("Uploaded", 90),
    ]:
        lbl = QLabel(label)
        lbl.setFont(_font(8, bold=True))
        lbl.setStyleSheet(
            f"color: {THEME['primary']}; background: transparent; border: none; "
            "text-transform: uppercase;"
        )
        lbl.setFixedWidth(width)
        h.addWidget(lbl)
    h.addStretch()
    blank = QLabel("")
    blank.setFixedWidth(140)
    h.addWidget(blank)
    return hdr


# ── MRS Editor Dialog ─────────────────────────────────────────────────────────

class MRSEditorDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Master Rate Sheet")
        self.setMinimumSize(1100, 700)
        self.resize(1300, 800)
        self._changes: dict = {}
        self._loading = False
        self._data = None
        self._setup_ui()
        self._load_data()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        title_bar = QWidget()
        title_bar.setStyleSheet(f"background: {THEME['primary_pale']}; border-bottom: 1px solid {THEME['border']};")
        tb = QHBoxLayout(title_bar)
        tb.setContentsMargins(16, 10, 16, 10)
        lbl = QLabel("Edit Master Rate Sheet")
        lbl.setFont(_font(13, bold=True))
        lbl.setStyleSheet(f"color: {THEME['primary']}; background: transparent; border: none;")
        tb.addWidget(lbl)
        tb.addStretch()
        self._change_lbl = QLabel("No changes")
        self._change_lbl.setFont(_font(9))
        self._change_lbl.setStyleSheet(f"color: {THEME['text_muted']}; background: transparent; border: none;")
        tb.addWidget(self._change_lbl)
        layout.addWidget(title_bar)

        # Search row — filters by CPT code or test name (first two cols)
        search_bar = QWidget()
        search_bar.setStyleSheet(
            f"background: white; border-bottom: 1px solid {THEME['border']};"
        )
        sb = QHBoxLayout(search_bar)
        sb.setContentsMargins(16, 8, 16, 8)
        sb.setSpacing(6)
        ic = QLabel("🔍")
        ic.setFont(_font(11))
        ic.setStyleSheet("background: transparent; border: none;")
        self._search = QLineEdit()
        self._search.setPlaceholderText("Search by CPT code or test name…")
        self._search.setFont(_font(10))
        self._search.setStyleSheet(
            f"QLineEdit {{ background: white; border: 1.5px solid {THEME['border']}; "
            f"border-radius: 8px; padding: 6px 10px; color: {THEME['text_primary']}; }}"
            f"QLineEdit:focus {{ border-color: {THEME['primary']}; }}"
        )
        self._search.textChanged.connect(self._filter_table)
        self._visible_count_lbl = QLabel("")
        self._visible_count_lbl.setFont(_font(8))
        self._visible_count_lbl.setStyleSheet(
            f"color: {THEME['text_muted']}; background: transparent; border: none;"
        )
        sb.addWidget(ic)
        sb.addWidget(self._search, 1)
        sb.addWidget(self._visible_count_lbl)
        layout.addWidget(search_bar)

        self._table = QTableWidget()
        self._table.setStyleSheet(_table_style())
        self._table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self._table.horizontalHeader().setStretchLastSection(False)
        self._table.verticalHeader().setVisible(False)
        self._table.setAlternatingRowColors(True)
        self._table.itemChanged.connect(self._on_cell_changed)
        layout.addWidget(self._table, 1)

        btn_bar = QWidget()
        btn_bar.setStyleSheet(f"background: white; border-top: 1px solid {THEME['border']};")
        bb = QHBoxLayout(btn_bar)
        bb.setContentsMargins(16, 10, 16, 10)
        bb.setSpacing(10)
        bb.addStretch()
        cancel = SecondaryButton("Cancel")
        cancel.clicked.connect(self.reject)
        self._save_btn = PrimaryButton("Save Changes", variant="navy")
        self._save_btn.clicked.connect(self._save)
        self._save_btn.setEnabled(False)
        bb.addWidget(cancel)
        bb.addWidget(self._save_btn)
        layout.addWidget(btn_bar)

    def _load_data(self):
        try:
            from core.master_rate_editor import read_mrs_for_display
            self._data = read_mrs_for_display()
        except Exception as e:
            QMessageBox.critical(self, "Load Error", str(e))
            self.reject()
            return

        headers = self._data["headers"]
        rows = self._data["rows"]
        self._loading = True
        self._table.setColumnCount(len(headers))
        self._table.setRowCount(len(rows))
        self._table.setHorizontalHeaderLabels(headers)
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                item = QTableWidgetItem(val)
                self._table.setItem(r, c, item)
        self._table.resizeColumnsToContents()
        for c in range(len(headers)):
            if self._table.columnWidth(c) > 160:
                self._table.setColumnWidth(c, 160)
        self._loading = False
        self._update_visible_count()

    def _filter_table(self, text: str):
        text = text.strip().lower()
        for r in range(self._table.rowCount()):
            if not text:
                self._table.setRowHidden(r, False)
                continue
            match = False
            # Search the first two columns (Test Code + Test Name)
            for c in range(min(2, self._table.columnCount())):
                item = self._table.item(r, c)
                if item and text in item.text().lower():
                    match = True
                    break
            self._table.setRowHidden(r, not match)
        self._update_visible_count()

    def _update_visible_count(self):
        total = self._table.rowCount()
        visible = sum(
            1 for r in range(total) if not self._table.isRowHidden(r)
        )
        if visible == total:
            self._visible_count_lbl.setText(f"{total} tests")
        else:
            self._visible_count_lbl.setText(f"{visible} of {total} tests")

    def _on_cell_changed(self, item: QTableWidgetItem):
        if self._loading or self._data is None:
            return
        row_idx = item.row()
        col_idx = item.column()
        col_name = self._data["headers"][col_idx] if col_idx < len(self._data["headers"]) else ""
        original = self._data["rows"][row_idx][col_idx] if row_idx < len(self._data["rows"]) else ""
        new_val = item.text()
        if new_val != original:
            self._changes[(row_idx, col_name)] = new_val
            item.setBackground(QColor("#FFFFF0"))
        else:
            self._changes.pop((row_idx, col_name), None)
            item.setBackground(QColor("#FFFFFF"))
        n = len(self._changes)
        self._change_lbl.setText(f"{n} cell{'s' if n != 1 else ''} changed" if n else "No changes")
        self._save_btn.setEnabled(bool(self._changes))

    def _save(self):
        if not self._changes:
            self.accept()
            return
        try:
            from core.master_rate_editor import write_cells_batch
            write_cells_batch(self._changes)
            QMessageBox.information(
                self, "Saved",
                f"{len(self._changes)} cell change(s) written to Master Rate Sheet.",
            )
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Save Failed", str(e))


# ── Upload Addendum Dialog ────────────────────────────────────────────────────

class UploadAddendumDialog(QDialog):
    def __init__(self, pc_df: pd.DataFrame, master_path: str, parent=None):
        super().__init__(parent)
        self.pc_df = pc_df
        self._master_path = master_path
        self._parsed_tests: list = []
        self._selected_path: str = ""
        self.setWindowTitle("Upload Addendum")
        self.setMinimumSize(680, 560)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(14)

        title = QLabel("Upload Addendum")
        title.setFont(_font(13, bold=True))
        title.setStyleSheet(f"color: {THEME['primary']}; background: transparent; border: none;")
        layout.addWidget(title)

        # Step 1: Party selector
        p_lbl = QLabel("Select party this addendum belongs to")
        p_lbl.setFont(_font(9, bold=True))
        p_lbl.setStyleSheet(f"color: {THEME['text_secondary']}; background: transparent; border: none;")
        layout.addWidget(p_lbl)

        self._party_combo = QComboBox()
        self._party_combo.setFont(_font(10))
        self._party_combo.setMinimumHeight(36)
        self._party_combo.setStyleSheet(COMBO_STYLE)
        self._party_combo.addItem("— Select a party —", "")
        for _, row in self.pc_df.iterrows():
            full = str(row.get(PC_FULL_NAME, "")).strip()
            short = str(row.get(PC_SHORT, "")).strip()
            if full and short:
                self._party_combo.addItem(full, short)
        layout.addWidget(self._party_combo)

        # Step 2: File
        f_lbl = QLabel("Annexure-B Excel file")
        f_lbl.setFont(_font(9, bold=True))
        f_lbl.setStyleSheet(f"color: {THEME['text_secondary']}; background: transparent; border: none;")
        layout.addWidget(f_lbl)

        self._drop = FileDropZone(
            "Drop Annexure-B Excel here or click to browse",
            "Excel Files (*.xlsx *.xls)",
        )
        self._drop.file_selected.connect(self._on_file_selected)
        layout.addWidget(self._drop)

        self._preview_lbl = QLabel("")
        self._preview_lbl.setFont(_font(8))
        self._preview_lbl.setStyleSheet(f"color: {THEME['text_muted']}; background: transparent; border: none;")
        layout.addWidget(self._preview_lbl)

        # Step 3: Effective date
        d_lbl = QLabel("Effective from (first billing month this applies):")
        d_lbl.setFont(_font(9, bold=True))
        d_lbl.setStyleSheet(f"color: {THEME['text_secondary']}; background: transparent; border: none;")
        layout.addWidget(d_lbl)

        date_row = QHBoxLayout()
        self._date_edit = QDateEdit()
        self._date_edit.setCalendarPopup(True)
        self._date_edit.setDate(QDate.currentDate())
        self._date_edit.setFont(_font(10))
        self._date_edit.setFixedWidth(180)
        self._date_edit.setStyleSheet(
            f"QDateEdit {{ background: white; border: 1.5px solid {THEME['border']}; "
            f"border-radius: 8px; padding: 5px 10px; color: {THEME['text_primary']}; }}"
            f"QDateEdit:focus {{ border-color: {THEME['primary']}; }}"
        )
        date_row.addWidget(self._date_edit)
        date_row.addStretch()
        layout.addLayout(date_row)

        layout.addWidget(AlertBanner(
            "Any date you enter is rounded to the 1st of that month for billing. "
            "The actual date entered is recorded in Party_Config remarks.",
            "warning",
        ))

        layout.addWidget(AlertBanner(
            "Prices will be written directly into the Master Rate Sheet. "
            "New tests not yet in the MRS will be added as new rows.",
            "info",
        ))

        layout.addStretch()

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        btn_row.addStretch()
        cancel = SecondaryButton("Cancel")
        cancel.clicked.connect(self.reject)
        self._save_btn = PrimaryButton("Preview Changes →", variant="navy")
        self._save_btn.clicked.connect(self._preview)
        btn_row.addWidget(cancel)
        btn_row.addWidget(self._save_btn)
        layout.addLayout(btn_row)

    def _on_file_selected(self, path: str):
        self._selected_path = path
        self._parsed_tests = []
        try:
            from core.addendum_store import parse_addendum_excel
            tests = parse_addendum_excel(path)
            self._parsed_tests = tests
            self._preview_lbl.setText(f"✓  {len(tests)} tests found in file.")
            self._preview_lbl.setStyleSheet(
                f"color: {THEME['success']}; background: transparent; border: none;"
            )
        except Exception as e:
            self._preview_lbl.setText(f"Error: {e}")
            self._preview_lbl.setStyleSheet(
                f"color: {THEME['error']}; background: transparent; border: none;"
            )

    def _get_mr_column(self, short_code: str) -> tuple[str, str | None]:
        """Return (his_name, mrs_column_name) for the given short code."""
        from config import PC_SHORT, PC_HIS_NAME, MR_SHEET_NAME, MR_HEADER_ROW
        his_name = ""
        for _, row in self.pc_df.iterrows():
            if str(row.get(PC_SHORT, "")).strip() == short_code:
                his_name = str(row.get(PC_HIS_NAME, "")).strip()
                break
        if not his_name:
            return "", None
        try:
            master_df = pd.read_excel(
                self._master_path, sheet_name=MR_SHEET_NAME,
                header=MR_HEADER_ROW, dtype=str,
            )
            master_df.columns = [str(c).strip() for c in master_df.columns]
            from core.party_matcher import resolve_mr_column
            return his_name, resolve_mr_column(his_name, master_df)
        except Exception:
            return his_name, None

    def _preview(self):
        short_code = self._party_combo.currentData()
        if not short_code:
            QMessageBox.warning(self, "Select Party", "Please select a party.")
            return
        if not self._parsed_tests:
            QMessageBox.warning(self, "No File", "Please select and parse an Annexure-B Excel file first.")
            return

        qd = self._date_edit.date()
        actual_date  = date(qd.year(), qd.month(), qd.day())
        billing_date = date(qd.year(), qd.month(), 1)   # rounded to 1st for JSON

        his_name, mr_column = self._get_mr_column(short_code)
        if not mr_column:
            QMessageBox.critical(
                self, "Party Not Found",
                f"Cannot find the Master Rate Sheet column for '{short_code}'.\n"
                "Check Party_Config and PARTY_COL_MAP in config.py.",
            )
            return

        try:
            from core.addendum_writer import compute_addendum_preview, validate_addendum_against_mrp
            changes = compute_addendum_preview(self._parsed_tests, mr_column, self._master_path)
        except Exception as e:
            QMessageBox.critical(self, "Preview Failed", str(e))
            return

        # Hard-reject: discounted price must not exceed MRP for any existing test
        validation_errors = validate_addendum_against_mrp(changes)
        if validation_errors:
            QMessageBox.critical(
                self, "Addendum Rejected — Price Above MRP",
                "The following test(s) have a discounted price greater than the "
                "Master Rate Sheet MRP. Per the agreement, addendum prices must "
                "not exceed MRP. Fix the addendum file and try again:\n\n"
                + "\n".join(validation_errors[:20])
                + ("\n  …" if len(validation_errors) > 20 else ""),
            )
            return

        dlg = AddendumChangesDialog(
            short_code=short_code,
            his_name=his_name,
            actual_date=actual_date,
            billing_date=billing_date,
            changes=changes,
            tests=self._parsed_tests,
            mr_column=mr_column,
            master_path=self._master_path,
            parent=self,
        )
        if dlg.exec_() == QDialog.Accepted:
            self.accept()


# ── Addendum Changes Preview Dialog ──────────────────────────────────────────

class AddendumChangesDialog(QDialog):
    def __init__(self, short_code: str, his_name: str, actual_date: date,
                 billing_date: date, changes: list, tests: list,
                 mr_column: str, master_path: str, parent=None):
        super().__init__(parent)
        self._short_code  = short_code
        self._his_name    = his_name
        self._actual_date = actual_date
        self._billing_date = billing_date
        self._changes     = changes
        self._tests       = tests
        self._mr_column   = mr_column
        self._master_path = master_path
        self.setWindowTitle(f"Review Changes — {short_code}")
        self.setMinimumSize(860, 560)
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 16)
        layout.setSpacing(12)

        title = QLabel(f"Addendum preview — <b>{self._short_code}</b>  ·  column: {self._mr_column}")
        title.setFont(_font(12, bold=True))
        title.setTextFormat(Qt.RichText)
        title.setStyleSheet(f"color: {THEME['primary']}; background: transparent; border: none;")
        layout.addWidget(title)

        updates  = sum(1 for c in self._changes if not c["is_new_row"])
        new_rows = sum(1 for c in self._changes if c["is_new_row"])
        layout.addWidget(AlertBanner(
            f"Recorded date: <b>{self._actual_date.strftime('%d-%b-%Y')}</b>  ·  "
            f"Billing from: <b>01-{self._actual_date.strftime('%b-%Y')}</b>  ·  "
            f"<b>{updates}</b> price{'s' if updates != 1 else ''} updated  ·  "
            f"<b>{new_rows}</b> new row{'s' if new_rows != 1 else ''} added to MRS",
            "info",
        ))

        cols = ["CPT Code", "Test Name", "Current MRS Price (₹)", "New Price (₹)", "Difference (₹)", "Action"]
        tbl = QTableWidget(len(self._changes), len(cols))
        tbl.setHorizontalHeaderLabels(cols)
        tbl.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        for ci in [0, 2, 3, 4, 5]:
            tbl.horizontalHeader().setSectionResizeMode(ci, QHeaderView.ResizeToContents)
        tbl.verticalHeader().setVisible(False)
        tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        tbl.setAlternatingRowColors(True)
        tbl.setStyleSheet(_table_style())

        for r, ch in enumerate(self._changes):
            old_p = ch["old_price"]
            new_p = ch["new_price"]
            diff  = (new_p - old_p) if old_p is not None else None

            old_str  = f"{old_p:,.2f}" if old_p is not None else "—  (new test)"
            new_str  = f"{new_p:,.2f}"
            diff_str = f"{diff:+,.2f}" if diff is not None else "new"
            action   = "New row" if ch["is_new_row"] else "Update"

            row_bg = QColor(THEME["primary_pale"]) if ch["is_new_row"] else None

            for ci, val in enumerate([ch["test_code"], ch["test_name"],
                                       old_str, new_str, diff_str, action]):
                item = QTableWidgetItem(val)
                item.setToolTip(val)
                if row_bg:
                    item.setBackground(row_bg)
                if ci == 4 and diff is not None:
                    item.setForeground(
                        QColor(THEME["success"]) if diff <= 0 else QColor(THEME["warning"])
                    )
                tbl.setItem(r, ci, item)

        layout.addWidget(tbl, 1)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        btn_row.addStretch()
        cancel = SecondaryButton("Cancel")
        cancel.clicked.connect(self.reject)
        apply_btn = PrimaryButton("✓  Apply to Master Rate Sheet", variant="navy")
        apply_btn.clicked.connect(self._apply)
        btn_row.addWidget(cancel)
        btn_row.addWidget(apply_btn)
        layout.addLayout(btn_row)

    def _apply(self):
        from core.logger import get_logger
        from core.mrs_safety import backup_mrs, restore_from_backup, mrs_lock, prune_backups
        log = get_logger("ui.AddendumChangesDialog")

        # Snapshot current MRS before any write — we need this to roll back
        try:
            backup_path = backup_mrs(self._master_path, tag=f"pre_addendum_{self._short_code}")
        except Exception as e:
            log.error("Cannot back up MRS before addendum apply: %s", e)
            QMessageBox.critical(
                self, "Backup Failed",
                f"Could not create a backup before applying changes:\n{e}\n\n"
                "Aborted. The Master Rate Sheet was not modified.",
            )
            return

        json_written = False
        try:
            with mrs_lock(self._master_path):
                from core.addendum_writer import apply_addendum
                result = apply_addendum(
                    self._tests,
                    self._mr_column,
                    self._his_name,
                    self._actual_date,
                    self._master_path,
                )
                from core.addendum_store import save_addendum
                save_addendum(self._short_code, self._billing_date, self._tests)
                json_written = True

            log.info(
                "Addendum applied: party=%s, updated=%d, added=%d",
                self._short_code, result["updated"], result["added"],
            )
            prune_backups()  # housekeeping after a successful apply

            msg = f"Master Rate Sheet updated successfully.\n\n• {result['updated']} test price(s) updated"
            if result["added"]:
                msg += f"\n• {result['added']} new test(s) added as new rows"
            msg += (f"\n\nParty_Config Remarks updated to:\n"
                    f"\"Addendum applied {self._actual_date.strftime('%d-%b-%Y')}\"")
            QMessageBox.information(self, "Applied", msg)
            self.accept()
        except Exception as e:
            log.error("Addendum apply failed: %s", e, exc_info=True)
            # Roll back: restore MRS, drop partial JSON if any
            rollback_ok = True
            try:
                restore_from_backup(backup_path, self._master_path)
            except Exception as restore_err:
                rollback_ok = False
                log.critical("ROLLBACK FAILED: %s", restore_err, exc_info=True)
            if json_written:
                try:
                    from core.addendum_store import delete_addendum
                    delete_addendum(self._short_code)
                except Exception as del_err:
                    log.error("Could not delete partial addendum JSON: %s", del_err)

            if rollback_ok:
                QMessageBox.critical(
                    self, "Apply Failed — Changes Rolled Back",
                    f"The addendum could not be applied:\n\n{e}\n\n"
                    f"The Master Rate Sheet has been restored to its previous state.\n"
                    f"Backup kept at:\n{backup_path}",
                )
            else:
                QMessageBox.critical(
                    self, "Apply Failed — ROLLBACK ALSO FAILED",
                    f"The addendum apply failed:\n{e}\n\n"
                    f"AND the rollback also failed. Manual recovery needed.\n"
                    f"Restore the Master Rate Sheet by hand from:\n{backup_path}",
                )


# ── View Addendum Dialog ──────────────────────────────────────────────────────

class ViewAddendumDialog(QDialog):
    def __init__(self, data: dict, mrp_map: dict, parent=None):
        super().__init__(parent)
        sc = data.get("party_short_code", "")
        eff = data.get("effective_date", "")
        self.setWindowTitle(f"Addendum — {sc}  ·  Effective {eff}")
        self.setMinimumSize(640, 480)
        self._setup_ui(data, mrp_map)

    def _setup_ui(self, data: dict, mrp_map: dict):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(10)

        title = QLabel(
            f"<b>{data.get('party_short_code', '')}</b>  ·  Effective from {data.get('effective_date', '')}"
        )
        title.setFont(_font(11))
        title.setTextFormat(Qt.RichText)
        title.setStyleSheet(f"color: {THEME['primary']}; background: transparent; border: none;")
        layout.addWidget(title)

        tests = data.get("tests", [])
        cols = ["Test Code", "Test Name", "Discounted Price", "MRP (ref)", "Discount"]
        tbl = QTableWidget(len(tests), len(cols))
        tbl.setHorizontalHeaderLabels(cols)
        tbl.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        tbl.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        tbl.verticalHeader().setVisible(False)
        tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        tbl.setAlternatingRowColors(True)
        tbl.setStyleSheet(_table_style())

        for r, test in enumerate(tests):
            tc = str(test.get("test_code", ""))
            tn = str(test.get("test_name", ""))
            dp = float(test.get("discounted_price", 0))
            mrp_str = mrp_map.get(tc, "")
            try:
                mrp_val = float(mrp_str)
                disc = mrp_val - dp
                mrp_disp = f"{mrp_val:,.2f}"
                disc_disp = f"{disc:,.2f}"
            except (ValueError, TypeError):
                mrp_disp = mrp_str
                disc_disp = "—"
            for c, val in enumerate([tc, tn, f"{dp:,.2f}", mrp_disp, disc_disp]):
                item = QTableWidgetItem(val)
                tbl.setItem(r, c, item)

        layout.addWidget(tbl, 1)

        close = PrimaryButton("Close", variant="navy")
        close.clicked.connect(self.accept)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(close)
        layout.addLayout(btn_row)


# ── Documents Screen ──────────────────────────────────────────────────────────

class DocumentsScreen(QWidget):
    mrs_loaded = pyqtSignal(str)
    request_tab_switch = pyqtSignal(int)

    def __init__(self, app_state: dict, parent=None):
        super().__init__(parent)
        self.app_state = app_state
        self._pc_df: pd.DataFrame | None = None
        self._mrs_preview_table = None
        self._setup_ui()

    def _setup_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet(f"background: {THEME['surface']}; border: none;")

        container = QWidget()
        container.setStyleSheet(f"background: {THEME['surface']};")
        self._layout = QVBoxLayout(container)
        self._layout.setContentsMargins(28, 20, 28, 24)
        self._layout.setSpacing(16)

        self._layout.addWidget(_section_header(
            "Documents & Settings",
            "Manage the Master Rate Sheet and party addendums.",
        ))

        # MRS card placeholder (rebuilt dynamically)
        self._mrs_container = QWidget()
        self._mrs_container.setStyleSheet("background: transparent;")
        self._mrs_cl = QVBoxLayout(self._mrs_container)
        self._mrs_cl.setContentsMargins(0, 0, 0, 0)
        self._layout.addWidget(self._mrs_container)

        # Visual divider between sections
        _div = QFrame()
        _div.setFixedHeight(1)
        _div.setStyleSheet(f"background: {THEME['border']}; border: none;")
        self._layout.addWidget(_div)

        # Addendum card placeholder
        self._add_container = QWidget()
        self._add_container.setStyleSheet("background: transparent;")
        self._add_cl = QVBoxLayout(self._add_container)
        self._add_cl.setContentsMargins(0, 0, 0, 0)
        self._layout.addWidget(self._add_container)

        self._layout.addStretch()
        scroll.setWidget(container)
        outer.addWidget(scroll, 1)

        # Initial render
        self._rebuild_mrs()
        self._rebuild_addendum()

    def refresh(self):
        """Called by main_window when this tab becomes active."""
        self._load_pc_df()
        self._rebuild_mrs()
        self._rebuild_addendum()

    # ── MRS card ──────────────────────────────────────────────────────────────

    def _rebuild_mrs(self):
        self._mrs_preview_table = None
        _purge_layout(self._mrs_cl)
        card = CardFrame()
        card.set_header("📋", "Master Rate Sheet")

        mp = self.app_state.get("master_path")
        if mp and os.path.exists(mp):
            self._build_mrs_loaded(card, mp)
        else:
            drop = FileDropZone(
                "Upload Master Rate Sheet",
                "Excel Files (*.xlsx *.xls)",
            )
            drop.file_selected.connect(self._on_mrs_dropped)
            card.add_widget(drop)
            note = QLabel("Master_Rate_Sheet.xlsx — loaded once, stays persistent")
            note.setFont(_font(8))
            note.setStyleSheet(
                f"color: {THEME['text_muted']}; background: transparent; border: none;"
            )
            card.add_widget(note)

        self._mrs_cl.addWidget(card)

    def _build_mrs_loaded(self, card: CardFrame, mp: str):
        # File info row
        try:
            size = os.path.getsize(mp)
            mtime = datetime.fromtimestamp(os.path.getmtime(mp)).strftime("%d %b %Y, %H:%M")
            size_str = (f"{size/1024:.1f} KB" if size < 1024*1024
                        else f"{size/(1024*1024):.1f} MB")
        except Exception:
            size_str, mtime = "—", "—"

        info_row = QHBoxLayout()
        info_row.setSpacing(24)
        for lbl_txt, val_txt in [
            ("File", os.path.basename(mp)),
            ("Size", size_str),
            ("Last modified", mtime),
        ]:
            col = QVBoxLayout()
            col.setSpacing(2)
            l = QLabel(lbl_txt)
            l.setFont(_font(8))
            l.setStyleSheet(
                f"color: {THEME['text_muted']}; background: transparent; border: none;"
            )
            v = QLabel(val_txt)
            v.setFont(_font(9, bold=True))
            v.setStyleSheet(
                f"color: {THEME['text_primary']}; background: transparent; border: none;"
            )
            col.addWidget(l)
            col.addWidget(v)
            info_row.addLayout(col)
        info_row.addStretch()
        card.add_layout(info_row)

        # Buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        edit_btn = PrimaryButton("Edit in app →", variant="navy")
        edit_btn.clicked.connect(self._open_mrs_editor)
        repl_btn = SecondaryButton("Replace file")
        repl_btn.clicked.connect(self._replace_mrs)
        dl_btn = PrimaryButton("⬇  Download MRS", variant="success")
        dl_btn.clicked.connect(lambda: self._download_mrs(mp))
        btn_row.addWidget(edit_btn)
        btn_row.addWidget(repl_btn)
        btn_row.addWidget(dl_btn)
        btn_row.addStretch()
        card.add_layout(btn_row)

        # Search + preview table (all rows)
        try:
            df = pd.read_excel(
                mp, sheet_name=MR_SHEET_NAME,
                header=MR_HEADER_ROW, dtype=str,
            )
            df.columns = [str(c).strip() for c in df.columns]
            show_cols = list(df.columns[:6])
            df2 = df[show_cols].fillna("")

            # Search row
            search_row = QHBoxLayout()
            search_row.setSpacing(6)
            search_icon = QLabel("🔍")
            search_icon.setFont(_font(11))
            search_icon.setStyleSheet("background: transparent; border: none;")
            self._mrs_search = QLineEdit()
            self._mrs_search.setPlaceholderText("Search by CPT code or test name…")
            self._mrs_search.setFont(_font(9))
            self._mrs_search.setStyleSheet(
                f"QLineEdit {{ background: white; border: 1.5px solid {THEME['border']}; "
                f"border-radius: 8px; padding: 5px 10px; color: {THEME['text_primary']}; }}"
                f"QLineEdit:focus {{ border-color: {THEME['primary']}; }}"
            )
            count_lbl = QLabel(f"{len(df2)} tests")
            count_lbl.setFont(_font(8))
            count_lbl.setStyleSheet(
                f"color: {THEME['text_muted']}; background: transparent; border: none;"
            )
            search_row.addWidget(search_icon)
            search_row.addWidget(self._mrs_search, 1)
            search_row.addWidget(count_lbl)
            card.add_layout(search_row)

            self._mrs_preview_table = QTableWidget(len(df2), len(show_cols))
            self._mrs_preview_table.setHorizontalHeaderLabels(show_cols)
            self._mrs_preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
            self._mrs_preview_table.horizontalHeader().setStretchLastSection(True)
            self._mrs_preview_table.horizontalHeader().setMinimumSectionSize(90)
            self._mrs_preview_table.verticalHeader().setVisible(False)
            self._mrs_preview_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self._mrs_preview_table.setAlternatingRowColors(True)
            self._mrs_preview_table.setFixedHeight(260)
            self._mrs_preview_table.setStyleSheet(_table_style())
            for r, row_data in enumerate(df2.values.tolist()):
                for c, val in enumerate(row_data):
                    item = QTableWidgetItem(str(val))
                    item.setToolTip(str(val))
                    self._mrs_preview_table.setItem(r, c, item)
            card.add_widget(self._mrs_preview_table)

            self._mrs_search.textChanged.connect(self._filter_mrs_preview)
        except Exception:
            self._mrs_preview_table = None

    def _filter_mrs_preview(self, text: str):
        tbl = getattr(self, "_mrs_preview_table", None)
        if tbl is None:
            return
        for r in range(tbl.rowCount()):
            if not text:
                tbl.setRowHidden(r, False)
                continue
            match = any(
                tbl.item(r, c) and text.lower() in tbl.item(r, c).text().lower()
                for c in range(min(2, tbl.columnCount()))
            )
            tbl.setRowHidden(r, not match)

    def _download_mrs(self, mp: str):
        dest, _ = QFileDialog.getSaveFileName(
            self, "Save a copy of Master Rate Sheet",
            os.path.basename(mp),
            "Excel Files (*.xlsx)",
        )
        if not dest:
            return
        try:
            shutil.copy2(mp, dest)
            QMessageBox.information(self, "Saved", f"Master Rate Sheet saved to:\n{dest}")
        except Exception as e:
            QMessageBox.critical(self, "Save Failed", str(e))

    def _on_mrs_dropped(self, path: str):
        try:
            from core.master_rate_editor import replace_master_rate_sheet
            replace_master_rate_sheet(path)
            self.app_state["master_path"] = MASTER_RATE_PATH
            self.mrs_loaded.emit(MASTER_RATE_PATH)
            self._load_pc_df()
            self._rebuild_mrs()
            self._rebuild_addendum()
        except Exception as e:
            QMessageBox.critical(self, "Invalid File", str(e))

    def _open_mrs_editor(self):
        if not self.app_state.get("master_path"):
            return
        dlg = MRSEditorDialog(self)
        dlg.exec_()
        self._rebuild_mrs()

    def _replace_mrs(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select new Master Rate Sheet", "", "Excel Files (*.xlsx *.xls)"
        )
        if not path:
            return
        try:
            from core.master_rate_editor import replace_master_rate_sheet
            replace_master_rate_sheet(path)
            self.app_state["master_path"] = MASTER_RATE_PATH
            self.mrs_loaded.emit(MASTER_RATE_PATH)
            self._load_pc_df()
            self._rebuild_mrs()
            self._rebuild_addendum()
            QMessageBox.information(self, "Done", "Master Rate Sheet replaced successfully.")
        except Exception as e:
            QMessageBox.critical(self, "Replace Failed", str(e))

    # ── Addendum card ─────────────────────────────────────────────────────────

    def _rebuild_addendum(self):
        _purge_layout(self._add_cl)

        upload_btn = PrimaryButton("Upload New Addendum", variant="navy")
        upload_btn.setFixedHeight(32)
        upload_btn.clicked.connect(self._upload_addendum)

        card = CardFrame()
        card.set_header("📎", "Party Addendums", upload_btn)

        try:
            from core.addendum_store import list_addendums
            addendums = list_addendums()
        except Exception:
            addendums = []

        short_to_full: dict[str, str] = {}
        if self._pc_df is not None:
            for _, row in self._pc_df.iterrows():
                sc = str(row.get(PC_SHORT, "")).strip()
                fn = str(row.get(PC_FULL_NAME, "")).strip()
                if sc:
                    short_to_full[sc] = fn

        if addendums:
            card.add_widget(_addendum_table_header())
            for add in sorted(addendums, key=lambda x: x["party_short_code"]):
                sc = add["party_short_code"]
                up = add["uploaded_at"][:10] if add.get("uploaded_at") else "—"
                row_w = _AddendumRow(
                    sc,
                    short_to_full.get(sc, sc),
                    add["effective_date"],
                    add["test_count"],
                    up,
                )
                row_w.view_clicked.connect(self._view_addendum)
                row_w.delete_clicked.connect(self._delete_addendum)
                card.add_widget(row_w)
        else:
            none_lbl = QLabel(
                "No addendums stored yet. Click 'Upload New Addendum' to add one."
            )
            none_lbl.setFont(_font(9))
            none_lbl.setStyleSheet(
                f"color: {THEME['text_muted']}; background: transparent; border: none;"
            )
            card.add_widget(none_lbl)
            guide_lbl = QLabel(
                "Addendums define discounted prices that override the Master Rate Sheet "
                "for a specific party from an agreed effective date."
            )
            guide_lbl.setFont(_font(8))
            guide_lbl.setWordWrap(True)
            guide_lbl.setStyleSheet(
                f"color: {THEME['text_muted']}; background: transparent; border: none;"
            )
            card.add_widget(guide_lbl)

        if self._pc_df is not None:
            n_without = len(self._pc_df) - len(addendums)
            if n_without > 0:
                muted = QLabel(
                    f"{n_without} {'party has' if n_without == 1 else 'parties have'} no addendum stored."
                )
                muted.setFont(_font(8))
                muted.setStyleSheet(
                    f"color: {THEME['text_muted']}; background: transparent; border: none;"
                )
                card.add_widget(muted)

        self._add_cl.addWidget(card)

    def _upload_addendum(self):
        if self._pc_df is None:
            self._load_pc_df()
        if self._pc_df is None:
            QMessageBox.warning(
                self, "MRS Not Loaded",
                "Load the Master Rate Sheet first to manage addendums.",
            )
            return
        mp = self.app_state.get("master_path", "")
        dlg = UploadAddendumDialog(self._pc_df, mp, self)
        if dlg.exec_() == QDialog.Accepted:
            self._rebuild_mrs()
            self._rebuild_addendum()

    def _view_addendum(self, short_code: str):
        from core.addendum_store import load_addendum
        data = load_addendum(short_code)
        if not data:
            QMessageBox.warning(self, "Not Found", f"No addendum for {short_code}.")
            return
        mrp_map: dict[str, str] = {}
        mp = self.app_state.get("master_path")
        if mp and os.path.exists(mp):
            try:
                df = pd.read_excel(
                    mp, sheet_name=MR_SHEET_NAME,
                    header=MR_HEADER_ROW, dtype=str,
                )
                df.columns = [str(c).strip() for c in df.columns]
                for _, row in df.iterrows():
                    tc = str(row.get(COL_TEST_CODE, "")).strip()
                    mrp = str(row.get(COL_MRP, "")).strip()
                    if tc and mrp and mrp.lower() not in ("nan", "none", ""):
                        mrp_map[tc] = mrp
            except Exception:
                pass
        dlg = ViewAddendumDialog(data, mrp_map, self)
        dlg.exec_()

    def _delete_addendum(self, short_code: str):
        reply = QMessageBox.question(
            self, "Delete Addendum",
            f"Delete the stored addendum for '{short_code}'?\nThis cannot be undone.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if reply == QMessageBox.Yes:
            from core.addendum_store import delete_addendum
            delete_addendum(short_code)
            self._rebuild_addendum()

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _load_pc_df(self):
        mp = self.app_state.get("master_path")
        if not mp or not os.path.exists(mp):
            return
        try:
            df = pd.read_excel(mp, sheet_name=PC_SHEET_NAME, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]
            self._pc_df = df
        except Exception:
            self._pc_df = None
