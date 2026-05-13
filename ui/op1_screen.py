"""
ui/op1_screen.py — Operation 1: Upload Combined Export · Batch Working Sheets.
"""

import os
import pandas as pd
from datetime import datetime, date
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QScrollArea, QFrame, QSizePolicy, QMessageBox,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

from config import (
    THEME, MASTER_RATE_PATH, OUTPUT_DIR,
    MR_SHEET_NAME, PC_SHEET_NAME, MR_HEADER_ROW,
    PC_HIS_NAME, PC_FULL_NAME, PC_SHORT,
    CSV_COL_COMPANY, COL_MRP,
    UNMATCHED_COL_COMPANY, UNMATCHED_COL_SNO, UNMATCHED_COL_TEST_CODE,
    UNMATCHED_COL_TEST_NAME, UNMATCHED_COL_MRP, UNMATCHED_COL_DISCOUNTED,
)
from ui.widgets import (
    FileDropZone, StatusBadge, AlertBanner, CardFrame,
    DataTable, PrimaryButton, SecondaryButton, StatCard,
    _font,
)
from core.logger import get_logger

log = get_logger(__name__)


# ── Worker ─────────────────────────────────────────────────────────────────────

class Op1Worker(QThread):
    finished = pyqtSignal(dict)
    error    = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, input_path: str, master_path: str):
        super().__init__()
        self.input_path = input_path
        self.master_path = master_path

    def run(self):
        try:
            log.info("Op1 start: input=%s", self.input_path)
            self.progress.emit("Reading HIS export…")
            from core.csv_reader import read_his_export
            his_df = read_his_export(self.input_path)

            self.progress.emit("Loading Master Rate Sheet…")
            master_df = pd.read_excel(
                self.master_path, sheet_name=MR_SHEET_NAME,
                header=MR_HEADER_ROW, dtype=str,
            )
            master_df.columns = [str(c).strip() for c in master_df.columns]

            pc_df = pd.read_excel(self.master_path, sheet_name=PC_SHEET_NAME, dtype=str)
            pc_df.columns = [str(c).strip() for c in pc_df.columns]

            self.progress.emit("Validating export…")
            from core.validators import pre_flight
            pre_flight(his_df)

            # Clean company names and validate
            his_df[CSV_COL_COMPANY] = his_df[CSV_COL_COMPANY].astype(str).str.strip()
            blank_companies = (his_df[CSV_COL_COMPANY] == "") | his_df[CSV_COL_COMPANY].isna()
            if blank_companies.any():
                raise ValueError(
                    f"{int(blank_companies.sum())} row(s) have a blank Company Name. "
                    "Please fix the export and retry."
                )

            grouped = list(his_df.groupby(CSV_COL_COMPANY, sort=False))
            if not grouped:
                raise ValueError("No company data found after cleaning.")

            self.progress.emit("Resolving party mapping…")
            from core.party_matcher import match_party, resolve_mr_column

            company_map = []
            unmatched_companies = []
            missing_cols = []
            for company_name, cdf in grouped:
                matched_row = match_party(company_name, pc_df)
                if matched_row is None:
                    unmatched_companies.append(company_name)
                    continue
                his_name = str(matched_row[PC_HIS_NAME]).strip()
                mr_col = resolve_mr_column(his_name, master_df)
                if mr_col is None:
                    missing_cols.append(his_name)
                    continue
                company_map.append({
                    "company_name": company_name,
                    "df": cdf,
                    "matched_row": matched_row,
                    "mr_col": mr_col,
                })

            if unmatched_companies:
                bad = ", ".join(sorted(set(unmatched_companies)))
                raise ValueError(
                    "These Company Name values could not be matched in Party_Config: "
                    f"{bad}"
                )
            if missing_cols:
                bad = ", ".join(sorted(set(missing_cols)))
                raise ValueError(
                    "Missing party columns in Master Rate Sheet for: "
                    f"{bad}"
                )

            self.progress.emit("Preparing batch session…")
            from core.batch_store import create_batch_session_dir, save_session_manifest
            session_dir, session_id = create_batch_session_dir()

            for num_col in [COL_MRP]:
                if num_col in master_df.columns:
                    master_df[num_col] = pd.to_numeric(master_df[num_col], errors="coerce")

            total_matched = 0
            total_unmatched = 0
            total_amount = 0.0
            all_bill_errors = []
            unmatched_rows = []
            unmatched_seen = set()
            preview_rows = []
            companies = []

            from core.rate_engine import lookup_prices
            from core.working_sheet_writer import write_working_sheet, calculate_total
            from core.validators import verify_bill_amounts

            for item in company_map:
                company_name = item["company_name"]
                cdf = item["df"]
                matched_row = item["matched_row"]
                mr_col = item["mr_col"]

                if mr_col in master_df.columns:
                    master_df[mr_col] = pd.to_numeric(master_df[mr_col], errors="coerce")

                billing_period = _infer_billing_period(cdf)
                party_short = str(matched_row[PC_SHORT]).strip()
                party_full = str(matched_row[PC_FULL_NAME]).strip()

                matched_df, unmatched_df = lookup_prices(cdf, mr_col, master_df)

                bill_errors = verify_bill_amounts(matched_df)
                if bill_errors:
                    for err in bill_errors:
                        all_bill_errors.append(f"{party_full}: {err}")

                period_safe = billing_period.replace("'", "").replace(" ", "_")
                out_path = os.path.join(OUTPUT_DIR, f"{party_short}_{period_safe}_Working.xlsx")
                write_working_sheet(
                    matched_df, unmatched_df, party_full, billing_period, out_path
                )

                matched_file = os.path.join(session_dir, f"matched_{party_short}.csv")
                unmatched_file = os.path.join(session_dir, f"unmatched_{party_short}.csv")
                matched_df.to_csv(matched_file, index=False)
                unmatched_df.to_csv(unmatched_file, index=False)

                total_matched += len(matched_df)
                total_unmatched += len(unmatched_df)
                total_amount += calculate_total(matched_df)

                for _, r in unmatched_df.iterrows():
                    cpt_code = r.get("cpt_code", "")
                    key = (party_full, str(cpt_code).strip())
                    if key in unmatched_seen:
                        continue
                    unmatched_seen.add(key)
                    unmatched_rows.append({
                        UNMATCHED_COL_COMPANY: party_full,
                        UNMATCHED_COL_SNO: r.get("S.No.", ""),
                        UNMATCHED_COL_TEST_CODE: cpt_code,
                        UNMATCHED_COL_TEST_NAME: r.get("Service Name", ""),
                        UNMATCHED_COL_MRP: r.get("MRP", ""),
                        UNMATCHED_COL_DISCOUNTED: "",
                    })

                for _, r in matched_df.head(3).iterrows():
                    preview_rows.append([
                        party_full,
                        r.get("cpt_code", ""),
                        r.get("Service Name", ""),
                        r.get("MRP", ""),
                        r.get("Bill_Amount", ""),
                    ])

                companies.append({
                    "party_full": party_full,
                    "party_short": party_short,
                    "his_name": str(matched_row[PC_HIS_NAME]).strip(),
                    "mr_col": mr_col,
                    "billing_period": billing_period,
                    "matched_file": matched_file,
                    "unmatched_file": unmatched_file,
                    "ws_path": out_path,
                })

            unmatched_path = os.path.join(session_dir, f"Unmatched_Rates_{session_id}.xlsx")
            pd.DataFrame(unmatched_rows).to_excel(unmatched_path, index=False)

            stats = {
                "coop_removed":  his_df.attrs.get("coop_removed", 0),
                "billable_rows": his_df.attrs.get("billable_rows", 0),
                "companies":     len(companies),
                "matched":       total_matched,
                "unmatched":     total_unmatched,
                "total_amount":  total_amount,
            }

            manifest = {
                "session_id": session_id,
                "created_at": datetime.now().isoformat(timespec="seconds"),
                "master_path": self.master_path,
                "stats": stats,
                "companies": companies,
                "unmatched_template": unmatched_path,
            }
            session_path = save_session_manifest(session_dir, manifest)

            self.finished.emit({
                "session_path": session_path,
                "unmatched_path": unmatched_path,
                "stats": stats,
                "preview_rows": preview_rows,
                "bill_errors": all_bill_errors,
                "master_path": self.master_path,
            })

        except Exception as e:
            log.error("Op1 worker failed: %s", e, exc_info=True)
            self.error.emit(str(e))


def _infer_billing_period(his_df: pd.DataFrame) -> str:
    from config import CSV_COL_DATE
    try:
        dates = pd.to_datetime(his_df[CSV_COL_DATE], dayfirst=True, errors="coerce").dropna()
        if len(dates) > 0:
            most_common = dates.dt.to_period("M").mode()[0]
            return most_common.to_timestamp().strftime("%b'%y")
    except Exception:
        pass
    return date.today().strftime("%b'%y")


# ── Screen ─────────────────────────────────────────────────────────────────────

class Op1Screen(QWidget):
    op1_complete    = pyqtSignal(dict)
    mrs_loaded      = pyqtSignal(str)
    go_to_documents = pyqtSignal()

    def __init__(self, app_state: dict, parent=None):
        super().__init__(parent)
        self.app_state = app_state
        self._worker: Op1Worker | None = None
        self._setup_ui()

    # ── Layout ─────────────────────────────────────────────────────────────────

    def _setup_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        # Scrollable content
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setStyleSheet(f"background: {THEME['surface']}; border: none;")

        container = QWidget()
        container.setStyleSheet(f"background: {THEME['surface']};")
        self._layout = QVBoxLayout(container)
        self._layout.setContentsMargins(28, 20, 28, 20)
        self._layout.setSpacing(14)

        # Screen header
        self._layout.addWidget(_make_screen_header(
            1, "Upload Combined HIS Export & Generate Working Sheets",
            "Upload the monthly HIS Excel containing all companies. "
            "The tool will generate Working Sheets for every party and create a single "
            "Unmatched Rates template.",
        ))

        # Upload card
        self._upload_card = self._build_upload_card()
        self._layout.addWidget(self._upload_card)

        # Batch summary card
        self._batch_card = self._build_batch_card()
        self._layout.addWidget(self._batch_card)

        # Results section (hidden until run)
        self._results_section = self._build_results_section()
        self._results_section.setVisible(False)
        self._layout.addWidget(self._results_section)

        self._layout.addStretch()

        scroll.setWidget(container)
        outer.addWidget(scroll, 1)

        # Bottom action bar
        outer.addWidget(self._build_action_bar())

        self._refresh_mrs_state()

    # ── Card builders ──────────────────────────────────────────────────────────

    def _build_upload_card(self) -> CardFrame:
        card = CardFrame()
        card.set_header("📤", "File Upload")

        req_lbl = QLabel("REQUIRED FILES")
        req_lbl.setFont(_font(8, bold=True))
        req_lbl.setStyleSheet(
            f"color: {THEME['text_muted']}; background: transparent; border: none;"
        )
        card.add_widget(req_lbl)

        grid = QWidget()
        grid.setStyleSheet("background: transparent;")
        g_layout = QHBoxLayout(grid)
        g_layout.setContentsMargins(0, 0, 0, 0)
        g_layout.setSpacing(14)

        # HIS Excel side
        batch_col = QVBoxLayout()
        batch_label = QLabel("HIS Combined Export  *")
        batch_label.setFont(_font(9, bold=True))
        batch_label.setStyleSheet(f"color: {THEME['text_secondary']}; background: transparent;")
        self._batch_drop = FileDropZone(
            "Drop HIS Excel here or click to browse",
            "Excel Files (*.xlsx *.xls)",
        )
        self._batch_drop.set_large_mode()
        self._batch_drop.file_selected.connect(self._on_batch_selected)
        batch_col.addWidget(batch_label)
        batch_col.addWidget(self._batch_drop)

        # Master Rate Sheet side
        mrs_col = QVBoxLayout()
        mrs_head_row = QHBoxLayout()
        mrs_label = QLabel("Master Rate Sheet")
        mrs_label.setFont(_font(9, bold=True))
        mrs_label.setStyleSheet(f"color: {THEME['text_secondary']}; background: transparent;")
        self._mrs_persist_badge = QLabel("persistent · loaded once")
        self._mrs_persist_badge.setFont(_font(7, bold=True))
        self._mrs_persist_badge.setStyleSheet(
            f"color: {THEME['primary_mid']}; background: {THEME['primary_light']}; "
            "border-radius: 8px; padding: 2px 8px; border: none;"
        )
        mrs_head_row.addWidget(mrs_label)
        mrs_head_row.addWidget(self._mrs_persist_badge)
        mrs_head_row.addStretch()

        self._mrs_status_container = QWidget()
        self._mrs_status_container.setStyleSheet("background: transparent;")
        self._mrs_status_layout = QVBoxLayout(self._mrs_status_container)
        self._mrs_status_layout.setContentsMargins(0, 0, 0, 0)
        self._mrs_status_layout.setSpacing(6)

        mrs_col.addLayout(mrs_head_row)
        mrs_col.addWidget(self._mrs_status_container)

        g_layout.addLayout(batch_col, 1)
        g_layout.addLayout(mrs_col, 1)
        card.add_widget(grid)
        return card

    def _build_batch_card(self) -> CardFrame:
        card = CardFrame()
        card.set_header("📦", "Batch Summary")

        self._batch_hint = QLabel(
            "This run processes all companies in the uploaded file. "
            "Party matching happens automatically using Party_Config."
        )
        self._batch_hint.setFont(_font(9))
        self._batch_hint.setStyleSheet(
            f"color: {THEME['text_muted']}; background: transparent; border: none;"
        )
        card.add_widget(self._batch_hint)

        self._batch_stats = QLabel("—")
        self._batch_stats.setFont(_font(9, bold=True))
        self._batch_stats.setStyleSheet(
            f"color: {THEME['primary']}; background: transparent; border: none;"
        )
        card.add_widget(self._batch_stats)
        return card

    def _build_results_section(self) -> QWidget:
        w = QWidget()
        w.setStyleSheet("background: transparent;")
        layout = QVBoxLayout(w)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(14)

        # Stats row
        self._stat_companies = StatCard("Companies", "—", THEME["primary"])
        self._stat_matched   = StatCard("Matched", "—",  THEME["success"])
        self._stat_unmatched = StatCard("Unmatched", "—", THEME["warning"])
        self._stat_total_rows = StatCard("Total Rows", "—", THEME["primary"])
        self._stat_total_amt  = StatCard("Provisional Total", "—", THEME["primary_mid"])

        stats_row = QHBoxLayout()
        stats_row.setSpacing(10)
        for card in [self._stat_companies, self._stat_total_rows,
                 self._stat_matched, self._stat_unmatched, self._stat_total_amt]:
            stats_row.addWidget(card)
        layout.addLayout(stats_row)

        # Data table card
        table_card = CardFrame()
        self._result_badge = StatusBadge("0 matched", "success")
        table_card.set_header("📊", "Preview — Working Sheet Data", self._result_badge)
        table_card.set_body_padding(0, 0, 0, 0)

        self._data_table = DataTable([
            "Company", "CPT Code", "Service Name", "MRP ₹", "Bill Amount ₹",
        ])
        self._data_table.setMinimumHeight(220)
        table_card.add_widget(self._data_table)
        layout.addWidget(table_card)

        return w

    def _build_action_bar(self) -> QWidget:
        bar = QFrame()
        bar.setObjectName("ActionBar1")
        bar.setFixedHeight(54)
        bar.setStyleSheet(
            f"QFrame#ActionBar1 {{ background: white; border-top: 1px solid {THEME['border']}; }}"
        )
        h = QHBoxLayout(bar)
        h.setContentsMargins(24, 0, 24, 0)
        h.setSpacing(10)

        self._status_lbl = QLabel("Upload combined HIS Excel to begin.")
        self._status_lbl.setFont(_font(12, bold=True))
        self._status_lbl.setStyleSheet(f"color: {THEME['text_primary']}; background: transparent;")
        h.addWidget(self._status_lbl)
        h.addStretch()

        self._run_btn = PrimaryButton("▶  Run Operation 1", variant="navy")
        self._run_btn.setEnabled(False)
        self._run_btn.setToolTip("Upload combined HIS Excel to begin")
        self._run_btn.clicked.connect(self._run_op1)
        h.addWidget(self._run_btn)

        self._unmatched_btn = PrimaryButton("⬇  Unmatched Template", variant="navy")
        self._unmatched_btn.setVisible(False)
        self._unmatched_btn.clicked.connect(self._open_unmatched_template)
        h.addWidget(self._unmatched_btn)

        self._output_btn = PrimaryButton("📂  Output Folder", variant="success")
        self._output_btn.setVisible(False)
        self._output_btn.clicked.connect(self._open_output_folder)
        h.addWidget(self._output_btn)

        self._continue_btn = SecondaryButton("Continue to Op 2 →")
        self._continue_btn.setVisible(False)
        self._continue_btn.clicked.connect(self._proceed_to_op2)
        h.addWidget(self._continue_btn)

        return bar

    # ── MRS state ─────────────────────────────────────────────────────────────

    def _refresh_mrs_state(self):
        # Clear existing widgets from the container
        while self._mrs_status_layout.count():
            item = self._mrs_status_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        mp = self.app_state.get("master_path")
        if mp and os.path.exists(mp):
            # Green loaded state
            loaded_frame = QFrame()
            loaded_frame.setStyleSheet(
                "QFrame { background: #F0FBF6; border: 1px solid #1D9E75; border-radius: 8px; }"
            )
            lf = QVBoxLayout(loaded_frame)
            lf.setContentsMargins(12, 10, 12, 10)
            lf.setSpacing(3)
            row1 = QLabel(f"✓  {os.path.basename(mp)}")
            row1.setFont(_font(9, bold=True))
            row1.setStyleSheet("color: #085041; background: transparent; border: none;")
            row2 = QLabel("652 tests · 33 parties")
            row2.setFont(_font(8))
            row2.setStyleSheet("color: #1D9E75; background: transparent; border: none;")
            lf.addWidget(row1)
            lf.addWidget(row2)
            self._mrs_status_layout.addWidget(loaded_frame)
        else:
            # Not-loaded state: amber alert + navigation button
            alert = AlertBanner(
                "Master Rate Sheet not loaded. Go to Documents tab to upload it first.",
                "warning",
            )
            self._mrs_status_layout.addWidget(alert)
            btn = PrimaryButton("Go to Documents →", variant="ghost")
            btn.clicked.connect(self.go_to_documents.emit)
            self._mrs_status_layout.addWidget(btn)

    # ── File handlers ─────────────────────────────────────────────────────────

    def _on_batch_selected(self, path: str):
        self.app_state["batch_path"] = path
        self._results_section.setVisible(False)
        self._unmatched_btn.setVisible(False)
        self._output_btn.setVisible(False)
        self._continue_btn.setVisible(False)
        self._batch_stats.setText("—")
        self._refresh_run_btn()

    def _refresh_run_btn(self):
        ready = bool(
            self.app_state.get("batch_path")
            and self.app_state.get("master_path")
        )
        self._run_btn.setEnabled(ready)
        if not ready:
            self._status_lbl.setText("Upload the combined Excel to run.")

    # ── Run Op1 ───────────────────────────────────────────────────────────────

    def _run_op1(self):
        self._run_btn.setEnabled(False)
        self._results_section.setVisible(False)
        self._unmatched_btn.setVisible(False)
        self._output_btn.setVisible(False)
        self._continue_btn.setVisible(False)
        self._status_lbl.setText("Processing…")

        self._worker = Op1Worker(
            self.app_state["batch_path"],
            self.app_state["master_path"],
        )
        self._worker.progress.connect(self._status_lbl.setText)
        self._worker.finished.connect(self._on_op1_done)
        self._worker.error.connect(self._on_op1_error)
        self._worker.start()

    def _on_op1_done(self, result: dict):
        self.app_state["op1_result"] = result
        self.app_state["batch_session_path"] = result.get("session_path")
        stats = result["stats"]

        self._stat_companies.set_value(str(stats["companies"]))
        self._stat_total_rows.set_value(str(stats["billable_rows"]))
        self._stat_matched.set_value(str(stats["matched"]), THEME["success"])
        self._stat_unmatched.set_value(str(stats["unmatched"]), THEME["warning"])
        self._stat_total_amt.set_value(
            f"₹{stats['total_amount']:,.0f}", THEME["primary_mid"]
        )

        # Populate table
        self._data_table.clear()
        for row in result.get("preview_rows", [])[:25]:
            self._data_table.add_row([str(v) for v in row])

        self._result_badge.set_status("success", f"{stats['matched']} matched")
        self._batch_stats.setText(
            f"{stats['companies']} companies  ·  "
            f"{stats['matched']} matched  ·  {stats['unmatched']} unmatched"
        )
        self._status_lbl.setText(
            f"Batch ready — {stats['companies']} companies  ·  "
            f"{stats['unmatched']} unmatched"
        )

        self._run_btn.setEnabled(True)
        self._results_section.setVisible(True)
        self._unmatched_btn.setVisible(True)
        self._output_btn.setVisible(True)
        self._continue_btn.setVisible(True)

    def _on_op1_error(self, msg: str):
        self._run_btn.setEnabled(True)
        self._status_lbl.setText(f"Error: {msg[:80]}")
        QMessageBox.critical(self, "Operation 1 Failed", msg)

    # ── Download / proceed ────────────────────────────────────────────────────

    def _open_unmatched_template(self):
        path = self.app_state.get("op1_result", {}).get("unmatched_path")
        if path and os.path.exists(path):
            os.startfile(path)

    def _open_output_folder(self):
        if os.path.exists(OUTPUT_DIR):
            os.startfile(OUTPUT_DIR)

    def _proceed_to_op2(self):
        result = self.app_state.get("op1_result", {})
        self.op1_complete.emit(result)

    # ── Reset (called by MainWindow after session complete) ──────────────────

    def reset(self):
        self._batch_drop.clear_file()
        self.app_state["batch_path"] = None
        self.app_state["op1_result"] = None
        self.app_state["batch_session_path"] = None
        self._batch_stats.setText("—")
        self._results_section.setVisible(False)
        self._unmatched_btn.setVisible(False)
        self._output_btn.setVisible(False)
        self._continue_btn.setVisible(False)
        self._run_btn.setEnabled(False)
        self._status_lbl.setText("Upload combined HIS Excel to begin.")
        self._refresh_mrs_state()


# ── Helpers ────────────────────────────────────────────────────────────────────

def _make_screen_header(op_num: int, title: str, desc: str) -> QWidget:
    w = QWidget()
    w.setStyleSheet("background: transparent;")
    v = QVBoxLayout(w)
    v.setContentsMargins(0, 0, 0, 6)
    v.setSpacing(4)

    badge = QLabel(f"  ● Operation {op_num}  ")
    badge.setFont(_font(9, bold=True))
    badge.setStyleSheet(
        f"background: {THEME['primary_pale']}; color: {THEME['primary']}; "
        "border-radius: 10px; padding: 3px 8px; border: none;"
    )
    badge.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
    v.addWidget(badge)

    title_lbl = QLabel(title)
    title_lbl.setFont(_font(15, bold=True))
    title_lbl.setStyleSheet(f"color: {THEME['primary']}; background: transparent; border: none;")
    v.addWidget(title_lbl)

    desc_lbl = QLabel(desc)
    desc_lbl.setFont(_font(9))
    desc_lbl.setWordWrap(True)
    desc_lbl.setStyleSheet(f"color: {THEME['text_secondary']}; background: transparent; border: none;")
    v.addWidget(desc_lbl)

    return w


