"""
ui/op2_screen.py — Operation 2: Upload Unmatched Rates · Batch Invoices.
"""

import os
import pandas as pd
from datetime import date
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QDateEdit, QScrollArea, QFrame, QSizePolicy, QMessageBox,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QDate

from config import (
    THEME, OUTPUT_DIR,
    PC_FULL_NAME, PC_SHORT, PC_ADDRESS, PC_HIS_NAME, PC_SHEET_NAME,
    UNMATCHED_COL_COMPANY, UNMATCHED_COL_TEST_CODE, UNMATCHED_COL_TEST_NAME,
    UNMATCHED_COL_MRP, UNMATCHED_COL_DISCOUNTED,
)
from ui.widgets import (
    AlertBanner, CardFrame, ChecklistPanel, InvoicePreview,
    FileDropZone,
    PrimaryButton, SecondaryButton, _font,
)
from ui.op1_screen import _make_screen_header
from core.logger import get_logger

log = get_logger(__name__)


# ── Worker ─────────────────────────────────────────────────────────────────────

class Op2Worker(QThread):
    finished = pyqtSignal(dict)
    error    = pyqtSignal(str)
    progress = pyqtSignal(str)

    def __init__(self, params: dict):
        super().__init__()
        self.params = params

    def run(self):
        try:
            log.info("Op2 start (batch)")
            p = self.params
            master_path = p["master_path"]
            session_path = p["session_path"]
            unmatched_path = p["unmatched_path"]
            invoice_date = p["invoice_date"]

            from core.batch_store import load_session_manifest
            session = load_session_manifest(session_path)
            companies = session.get("companies", [])
            if not companies:
                raise ValueError("Batch session has no companies to process.")

            self.progress.emit("Loading Party_Config…")
            pc_df = pd.read_excel(master_path, sheet_name=PC_SHEET_NAME, dtype=str)
            pc_df.columns = [str(c).strip() for c in pc_df.columns]
            pc_map = {
                str(row[PC_HIS_NAME]).strip(): row for _, row in pc_df.iterrows()
            }

            self.progress.emit("Reading unmatched rates…")
            um_df = pd.read_excel(unmatched_path, dtype=str)
            um_df.columns = [str(c).strip() for c in um_df.columns]

            required_cols = [
                UNMATCHED_COL_COMPANY,
                UNMATCHED_COL_TEST_CODE,
                UNMATCHED_COL_TEST_NAME,
                UNMATCHED_COL_MRP,
                UNMATCHED_COL_DISCOUNTED,
            ]
            missing = [c for c in required_cols if c not in um_df.columns]
            if missing:
                raise ValueError(
                    f"Unmatched file is missing columns: {', '.join(missing)}."
                )

            um_df = um_df.fillna("")
            um_df[UNMATCHED_COL_COMPANY] = um_df[UNMATCHED_COL_COMPANY].astype(str).str.strip()
            um_df[UNMATCHED_COL_TEST_CODE] = um_df[UNMATCHED_COL_TEST_CODE].astype(str).str.strip()
            um_df[UNMATCHED_COL_MRP] = um_df[UNMATCHED_COL_MRP].astype(str).str.strip()
            um_df[UNMATCHED_COL_DISCOUNTED] = um_df[UNMATCHED_COL_DISCOUNTED].astype(str).str.strip()

            blank_mask = (
                (um_df[UNMATCHED_COL_COMPANY] == "")
                & (um_df[UNMATCHED_COL_TEST_CODE] == "")
            )
            um_df = um_df[~blank_mask].copy()

            blank_company = (um_df[UNMATCHED_COL_COMPANY] == "").sum()
            blank_code = (um_df[UNMATCHED_COL_TEST_CODE] == "").sum()
            blank_price = (um_df[UNMATCHED_COL_DISCOUNTED] == "").sum()
            if blank_company or blank_code or blank_price:
                raise ValueError(
                    "Unmatched file has blank values. "
                    f"Company blanks: {blank_company}, "
                    f"Test Code blanks: {blank_code}, "
                    f"Discounted Price blanks: {blank_price}."
                )

            company_rows_map = {
                name: df for name, df in um_df.groupby(UNMATCHED_COL_COMPANY, sort=False)
            }

            self.progress.emit("Validating rates…")
            from core.rate_engine import apply_confirmed_rates
            from core.validators import verify_bill_amounts
            from core.working_sheet_writer import write_working_sheet, calculate_total
            from core.invoice_numbering import get_next_invoice_no, increment_invoice_no
            from core.invoice_writer import write_invoice
            from core.amount_words import to_indian_words

            company_inputs = []
            errors = []
            for comp in companies:
                party_full = comp["party_full"]
                party_short = comp["party_short"]
                his_name = comp["his_name"]
                billing_period = comp.get("billing_period", "")
                matched_file = comp["matched_file"]
                unmatched_file = comp["unmatched_file"]

                if his_name not in pc_map:
                    errors.append(f"Party_Config row missing for {party_full}.")
                    continue

                matched_df = pd.read_csv(matched_file, dtype=str)
                unmatched_df = pd.read_csv(unmatched_file, dtype=str)

                for df in (matched_df, unmatched_df):
                    for col in ("MRP", "Discount", "Bill_Amount"):
                        if col in df.columns:
                            df[col] = pd.to_numeric(df[col], errors="coerce")

                rows = company_rows_map.get(party_full)
                if len(unmatched_df) > 0 and (rows is None or rows.empty):
                    errors.append(f"Missing rates for {party_full} in unmatched file.")
                    continue
                if len(unmatched_df) == 0 and rows is not None and not rows.empty:
                    errors.append(f"Unexpected rows for {party_full} in unmatched file.")
                    continue

                rate_map = {}
                if rows is not None:
                    for _, r in rows.iterrows():
                        cpt = str(r[UNMATCHED_COL_TEST_CODE]).strip()
                        try:
                            mrp = float(str(r[UNMATCHED_COL_MRP]).strip())
                            bill = float(str(r[UNMATCHED_COL_DISCOUNTED]).strip())
                        except ValueError:
                            errors.append(
                                f"{party_full} CPT {cpt}: MRP and Discounted Price must be numbers."
                            )
                            continue
                        if bill > mrp:
                            errors.append(
                                f"{party_full} CPT {cpt}: Discounted Price ({bill}) exceeds MRP ({mrp})."
                            )
                            continue
                        if cpt in rate_map and abs(rate_map[cpt]["bill_amount"] - bill) > 0.01:
                            errors.append(
                                f"{party_full} CPT {cpt}: multiple different prices in unmatched file."
                            )
                            continue
                        rate_map[cpt] = {
                            "cpt_code": cpt,
                            "mrp": mrp,
                            "discount": mrp - bill,
                            "bill_amount": bill,
                        }

                confirmed_rates = list(rate_map.values())
                confirmed_unmatched = apply_confirmed_rates(unmatched_df, confirmed_rates)

                blanks = (
                    confirmed_unmatched["Bill_Amount"].isna().sum()
                    if "Bill_Amount" in confirmed_unmatched.columns
                    else 0
                )
                if blanks > 0:
                    errors.append(
                        f"{party_full}: {blanks} unmatched test(s) still blank."
                    )
                    continue

                bill_errors = verify_bill_amounts(matched_df)
                if bill_errors:
                    errors.append(f"{party_full}: {bill_errors[0]}")
                    continue

                company_inputs.append({
                    "party_row": pc_map[his_name],
                    "party_full": party_full,
                    "party_short": party_short,
                    "billing_period": billing_period,
                    "matched_df": matched_df,
                    "confirmed_unmatched": confirmed_unmatched,
                })

            if errors:
                raise ValueError("\n".join(errors))

            self.progress.emit("Generating invoices…")
            results = []
            inv_incremented = 0
            for item in company_inputs:
                matched_df = item["matched_df"]
                confirmed_unmatched = item["confirmed_unmatched"]
                party_row = item["party_row"]
                party_full = str(party_row[PC_FULL_NAME]).strip()
                party_short = str(party_row[PC_SHORT]).strip()
                party_address = str(party_row[PC_ADDRESS]).strip()
                billing_period = item["billing_period"]

                period_safe = billing_period.replace("'", "").replace(" ", "_")
                out_ws = os.path.join(OUTPUT_DIR, f"{party_short}_{period_safe}_Working.xlsx")
                write_working_sheet(
                    matched_df, confirmed_unmatched, party_full, billing_period, out_ws
                )

                total_matched = calculate_total(matched_df)
                total_unmatched = (
                    confirmed_unmatched["Bill_Amount"].dropna().astype(float).sum()
                    if not confirmed_unmatched.empty and "Bill_Amount" in confirmed_unmatched.columns
                    else 0.0
                )
                total = total_matched + float(total_unmatched)

                invoice_no = get_next_invoice_no(party_row, date.today())

                out_inv = os.path.join(OUTPUT_DIR, f"{party_short}_{period_safe}_Invoice.xlsx")
                write_invoice(
                    party_full_name=party_full,
                    party_address=party_address,
                    invoice_no=invoice_no,
                    invoice_date=invoice_date,
                    billing_month=billing_period,
                    total_amount=total,
                    out_path=out_inv,
                )

                try:
                    increment_invoice_no(str(party_row[PC_HIS_NAME]).strip(), master_path)
                    inv_incremented += 1
                except Exception as inc_err:
                    log.error("Failed to increment invoice number: %s", inc_err, exc_info=True)

                results.append({
                    "ws_path": out_ws,
                    "inv_path": out_inv,
                    "invoice_no": invoice_no,
                    "invoice_date_str": invoice_date.strftime("%d.%m.%Y"),
                    "total": total,
                    "billing_period": billing_period,
                    "party_full": party_full,
                    "party_short": party_short,
                    "party_address": party_address,
                    "amount_words": to_indian_words(total),
                })

            last = results[-1] if results else {}
            checklist = _run_batch_checklist(len(results), inv_incremented)

            self.finished.emit({
                "results": results,
                "checklist": checklist,
                "companies": len(results),
                "inv_incremented": inv_incremented,
                "last_invoice": last,
            })

        except Exception as e:
            log.error("Op2 worker failed: %s", e, exc_info=True)
            self.error.emit(str(e))


def _run_batch_checklist(total_invoices: int, inv_incremented: int) -> list:
    items = []
    items.append({
        "label": "All unmatched rates confirmed",
        "status": "pass",
        "detail": "Validated in uploaded unmatched sheet",
    })
    items.append({
        "label": "Working Sheets generated",
        "status": "pass" if total_invoices > 0 else "fail",
        "detail": f"{total_invoices} file(s) written",
    })
    items.append({
        "label": "Invoices generated",
        "status": "pass" if total_invoices > 0 else "fail",
        "detail": f"{total_invoices} invoice(s) created",
    })
    items.append({
        "label": "Invoice numbers updated",
        "status": "pass" if inv_incremented == total_invoices else "warn",
        "detail": f"{inv_incremented}/{total_invoices} updated",
    })
    return items


# ── Screen ─────────────────────────────────────────────────────────────────────

class Op2Screen(QWidget):
    op2_complete = pyqtSignal()
    back_to_op1  = pyqtSignal()

    def __init__(self, app_state: dict, parent=None):
        super().__init__(parent)
        self.app_state   = app_state
        self._worker: Op2Worker | None = None
        self._session: dict | None = None
        self._unmatched_path: str | None = None
        self._op2_result: dict = {}
        self._setup_ui()

    # ── Layout ─────────────────────────────────────────────────────────────────

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
        self._content = QVBoxLayout(container)
        self._content.setContentsMargins(28, 20, 28, 20)
        self._content.setSpacing(14)

        self._content.addWidget(_make_screen_header(
            2, "Upload Unmatched Rates · Generate Invoices",
            "Upload the filled Unmatched Rates Excel and generate Working Sheets "
            "and Invoices for all companies in the batch.",
        ))

        self._op1_missing_banner = self._build_op1_missing_banner()
        self._content.addWidget(self._op1_missing_banner)

        self._op1_info_card = self._build_op1_info_card()
        self._content.addWidget(self._op1_info_card)

        # Step 1 card with an inner dynamic container
        self._step1_card = self._build_step1_card()
        self._content.addWidget(self._step1_card)

        # Step 2 — Invoice details
        self._content.addWidget(self._build_invoice_card())

        # Results (hidden)
        self._results_section = self._build_results_section()
        self._results_section.setVisible(False)
        self._content.addWidget(self._results_section)

        self._content.addStretch()

        scroll.setWidget(container)
        outer.addWidget(scroll, 1)
        outer.addWidget(self._build_action_bar())

        # Initialize invoice fields to disabled state
        self._set_invoice_fields_empty()

    # ── Card builders ──────────────────────────────────────────────────────────

    def _build_op1_missing_banner(self) -> QFrame:
        frame = QFrame()
        frame.setObjectName("Op1MissingBanner")
        frame.setStyleSheet(
            "QFrame#Op1MissingBanner { "
            "background: #FAEEDA; border: 1px solid #EF9F27; "
            "border-left: 4px solid #EF9F27; border-radius: 8px; }"
        )
        h = QHBoxLayout(frame)
        h.setContentsMargins(16, 12, 12, 12)
        h.setSpacing(14)
        txt = QLabel(
            "Complete Operation 1 first — upload the combined Excel and generate the batch."
        )
        txt.setFont(_font(10, bold=True))
        txt.setStyleSheet("color: #412402; background: transparent; border: none;")
        txt.setWordWrap(True)
        go_btn = PrimaryButton("Go to Op 1 →", variant="ghost")
        go_btn.setFixedHeight(34)
        go_btn.setFixedWidth(130)
        go_btn.clicked.connect(self.back_to_op1.emit)
        h.addWidget(txt, 1)
        h.addWidget(go_btn)
        frame.setVisible(False)
        return frame

    def _build_op1_info_card(self) -> QFrame:
        frame = QFrame()
        frame.setObjectName("Op1InfoCard")
        frame.setStyleSheet(
            f"QFrame#Op1InfoCard {{ "
            f"background: #E6F1FB; border: 1px solid {THEME['border']}; "
            "border-radius: 8px; }"
        )
        v = QVBoxLayout(frame)
        v.setContentsMargins(16, 12, 16, 12)
        v.setSpacing(4)
        self._op1_party_lbl = QLabel("—")
        self._op1_party_lbl.setFont(_font(14, bold=True))
        self._op1_party_lbl.setStyleSheet(
            f"color: {THEME['primary']}; background: transparent; border: none;"
        )
        self._op1_stats_lbl = QLabel("—")
        self._op1_stats_lbl.setFont(_font(9))
        self._op1_stats_lbl.setStyleSheet(
            f"color: {THEME['primary_mid']}; background: transparent; border: none;"
        )
        v.addWidget(self._op1_party_lbl)
        v.addWidget(self._op1_stats_lbl)
        frame.setVisible(False)
        return frame

    def _build_step1_card(self) -> CardFrame:
        card = CardFrame()
        card.set_header("📄", "Step 1 — Upload Filled Unmatched Rates")

        hint = QLabel(
            "Fill the Discounted Price column in the Unmatched Rates template, "
            "then upload it here."
        )
        hint.setFont(_font(9))
        hint.setStyleSheet(
            f"color: {THEME['text_secondary']}; background: transparent; border: none;"
        )
        card.add_widget(hint)

        self._unmatched_drop = FileDropZone(
            "Drop filled Unmatched Rates Excel here or click to browse",
            "Excel Files (*.xlsx *.xls)",
        )
        self._unmatched_drop.set_large_mode()
        self._unmatched_drop.file_selected.connect(self._on_unmatched_selected)
        card.add_widget(self._unmatched_drop)

        return card

    def _build_invoice_card(self) -> CardFrame:
        card = CardFrame()
        card.set_header("🧾", "Step 2 — Invoice Details")

        # Row 1: Invoice No | Invoice Date
        row1 = QHBoxLayout()
        row1.setSpacing(14)

        inv_no_v = QVBoxLayout()
        self._inv_no_label = QLabel("Invoice Numbers (auto-generated)")
        self._inv_no_label.setFont(_font(9))
        self._inv_no_label.setStyleSheet(f"color: {THEME['text_muted']}; background: transparent; border: none;")
        self._inv_no_display = QLabel("—")
        self._inv_no_display.setFont(_font(12, bold=True))
        self._inv_no_display.setMinimumHeight(36)
        self._inv_no_display.setStyleSheet(
            f"background: {THEME['surface']}; color: {THEME['text_muted']}; "
            f"border: 1px solid {THEME['border']}; border-radius: 8px; padding: 5px 10px;"
        )
        inv_no_v.addWidget(self._inv_no_label)
        inv_no_v.addWidget(self._inv_no_display)

        inv_date_v = QVBoxLayout()
        lbl2 = QLabel("Invoice Date  *")
        lbl2.setFont(_font(9))
        lbl2.setStyleSheet(f"color: {THEME['text_muted']}; background: transparent; border: none;")
        self._inv_date = QDateEdit()
        self._inv_date.setCalendarPopup(True)
        self._inv_date.setDate(QDate.currentDate())
        self._inv_date.setFont(_font(10))
        self._inv_date.setStyleSheet(
            f"QDateEdit {{ background: white; border: 1.5px solid {THEME['border']}; "
            f"border-radius: 8px; padding: 6px 10px; color: {THEME['text_primary']}; }}"
            f"QDateEdit:focus {{ border-color: {THEME['primary']}; }}"
        )
        inv_date_v.addWidget(lbl2)
        inv_date_v.addWidget(self._inv_date)

        row1.addLayout(inv_no_v, 1)
        row1.addLayout(inv_date_v, 1)
        card.add_layout(row1)

        # Row 2: Party | Billing Period
        row2 = QHBoxLayout()
        row2.setSpacing(14)

        party_v = QVBoxLayout()
        lbl3 = QLabel("Batch Scope")
        lbl3.setFont(_font(9))
        lbl3.setStyleSheet(f"color: {THEME['text_muted']}; background: transparent; border: none;")
        self._party_display = _readonly_lbl("—")
        party_v.addWidget(lbl3)
        party_v.addWidget(self._party_display)

        period_v = QVBoxLayout()
        lbl4 = QLabel("Billing Period")
        lbl4.setFont(_font(9))
        lbl4.setStyleSheet(f"color: {THEME['text_muted']}; background: transparent; border: none;")
        self._period_display = _readonly_lbl("—")
        period_v.addWidget(lbl4)
        period_v.addWidget(self._period_display)

        row2.addLayout(party_v, 1)
        row2.addLayout(period_v, 1)
        card.add_layout(row2)

        card.add_widget(AlertBanner(
            "Payment terms, bank details, and party address are auto-filled from "
            "Party_Config — not editable here.",
            "info",
        ))

        return card

    def _build_results_section(self) -> QWidget:
        w = QWidget()
        w.setStyleSheet("background: transparent;")
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(14)

        cl_card = CardFrame()
        cl_card.set_header("✓", "Verification Checklist")
        self._checklist = ChecklistPanel()
        cl_card.add_widget(self._checklist)
        v.addWidget(cl_card)

        inv_card = CardFrame()
        inv_card.set_header("🧾", "Invoice Preview")
        self._inv_preview = InvoicePreview()
        inv_card.add_widget(self._inv_preview)
        v.addWidget(inv_card)

        return w

    def _build_action_bar(self) -> QWidget:
        bar = QFrame()
        bar.setObjectName("ActionBar2")
        bar.setFixedHeight(54)
        bar.setStyleSheet(
            f"QFrame#ActionBar2 {{ background: white; border-top: 1px solid {THEME['border']}; }}"
        )
        h = QHBoxLayout(bar)
        h.setContentsMargins(24, 0, 24, 0)
        h.setSpacing(10)

        back_btn = SecondaryButton("← Back to Op 1")
        back_btn.clicked.connect(self.back_to_op1.emit)
        h.addWidget(back_btn)

        self._status_lbl = QLabel("")
        self._status_lbl.setFont(_font(9))
        self._status_lbl.setStyleSheet(
            f"color: {THEME['text_secondary']}; background: transparent;"
        )
        h.addWidget(self._status_lbl)
        h.addStretch()

        self._run_btn = PrimaryButton("▶  Run Operation 2", variant="ghost")
        self._run_btn.setEnabled(False)
        self._run_btn.clicked.connect(self._run_op2)
        h.addWidget(self._run_btn)

        self._open_output_btn = PrimaryButton("📂  Output Folder", variant="success")
        self._open_output_btn.setVisible(False)
        self._open_output_btn.clicked.connect(self._open_output_folder)
        h.addWidget(self._open_output_btn)

        return bar

    # ── Refresh ────────────────────────────────────────────────────────────────

    def refresh(self):
        session_path = self.app_state.get("batch_session_path")
        if not session_path:
            from core.batch_store import load_latest_session_path
            session_path = load_latest_session_path()
            if session_path:
                self.app_state["batch_session_path"] = session_path

        if not session_path:
            self._run_btn.setEnabled(False)
            self._status_lbl.setText("Complete Operation 1 first.")
            self._op1_missing_banner.setVisible(True)
            self._op1_info_card.setVisible(False)
            self._set_invoice_fields_empty()
            return

        from core.batch_store import load_session_manifest
        self._session = load_session_manifest(session_path)

        self._op1_missing_banner.setVisible(False)
        self._op1_info_card.setVisible(True)
        self._open_output_btn.setVisible(False)
        self._results_section.setVisible(False)
        self._unmatched_path = None
        self._unmatched_drop.clear_file()

        stats = self._session.get("stats", {})
        session_id = self._session.get("session_id", "—")
        self._op1_party_lbl.setText(f"Batch Session · {session_id}")
        self._op1_stats_lbl.setText(
            f"{stats.get('companies', 0)} companies  ·  "
            f"{stats.get('unmatched', 0)} unmatched  ·  "
            f"₹{stats.get('total_amount', 0):,.0f} provisional total"
        )

        self._set_invoice_fields_ready()
        self._update_run_btn()

    def _on_unmatched_selected(self, path: str):
        self._unmatched_path = path
        self._update_run_btn()

    def _set_invoice_fields_empty(self):
        self._inv_no_label.setText("Invoice Numbers  ·  available after Op 1")
        self._inv_no_display.setText("—")
        self._inv_no_display.setStyleSheet(
            f"background: {THEME['surface']}; color: {THEME['text_muted']}; "
            f"border: 1px solid {THEME['border']}; border-radius: 8px; "
            "font-size: 12pt; padding: 5px 10px;"
        )
        self._party_display.setText("—")
        self._party_display.setStyleSheet(
            f"background: {THEME['surface']}; color: {THEME['text_muted']}; "
            f"border: 1px solid {THEME['border']}; border-radius: 8px; padding: 5px 10px;"
        )
        self._period_display.setText("—")
        self._period_display.setStyleSheet(
            f"background: {THEME['surface']}; color: {THEME['text_muted']}; "
            f"border: 1px solid {THEME['border']}; border-radius: 8px; padding: 5px 10px;"
        )

    def _set_invoice_fields_ready(self):
        self._inv_no_label.setText("Invoice Numbers (auto-generated)")
        self._inv_no_display.setText("Multiple")
        self._inv_no_display.setStyleSheet(
            f"background: {THEME['primary_light']}; color: {THEME['primary']}; "
            f"border: 1px solid {THEME['border']}; border-radius: 8px; "
            "font-size: 12pt; font-weight: bold; padding: 5px 10px;"
        )
        self._party_display.setText("Multiple companies")
        self._party_display.setStyleSheet(
            f"background: {THEME['surface']}; color: {THEME['text_primary']}; "
            f"border: 1px solid {THEME['border']}; border-radius: 8px; padding: 5px 10px;"
        )
        self._period_display.setText("Per company")
        self._period_display.setStyleSheet(
            f"background: {THEME['surface']}; color: {THEME['text_primary']}; "
            f"border: 1px solid {THEME['border']}; border-radius: 8px; padding: 5px 10px;"
        )

    def _update_run_btn(self):
        if not self._session:
            self._run_btn.setEnabled(False)
            return
        if not self._unmatched_path or not os.path.exists(self._unmatched_path):
            self._run_btn.setEnabled(False)
            self._status_lbl.setText("Upload the filled Unmatched Rates file to enable Run.")
        else:
            self._run_btn.setEnabled(True)
            self._status_lbl.setText("Ready to run.")

    # ── Run ───────────────────────────────────────────────────────────────────

    def _run_op2(self):
        if not self._session:
            QMessageBox.warning(self, "Not Ready", "Complete Operation 1 first.")
            return
        if not self._unmatched_path or not os.path.exists(self._unmatched_path):
            QMessageBox.warning(self, "Missing File", "Upload the filled Unmatched Rates file.")
            return

        qd = self._inv_date.date()
        invoice_date = date(qd.year(), qd.month(), qd.day())

        params = {
            "master_path":     self.app_state["master_path"],
            "session_path":    self.app_state.get("batch_session_path"),
            "unmatched_path":  self._unmatched_path,
            "invoice_date":    invoice_date,
        }

        self._run_btn.setEnabled(False)
        self._results_section.setVisible(False)
        self._open_output_btn.setVisible(False)
        self._status_lbl.setText("Processing…")

        self._worker = Op2Worker(params)
        self._worker.progress.connect(self._status_lbl.setText)
        self._worker.finished.connect(self._on_op2_done)
        self._worker.error.connect(self._on_op2_error)
        self._worker.start()

    def _on_op2_done(self, result: dict):
        self._op2_result = result
        self.app_state["op2_result"] = result
        self._run_btn.setEnabled(True)

        results = result.get("results", [])
        count = len(results)
        inc_ok = result.get("inv_incremented", 0)
        inv_tag = f"{inc_ok}/{count} numbers updated"
        self._status_lbl.setText(
            f"Done — {count} invoices generated  ·  {inv_tag}"
        )

        if inc_ok < count:
            QMessageBox.warning(
                self, "Invoice Numbers Not Fully Saved",
                "Some invoice numbers could not be updated in Party_Config.\n\n"
                "If you keep the generated invoices, please update the Last "
                "Invoice Number for those parties manually in the Master Rate Sheet.\n\n"
                "Details are in output/app.log.",
            )

        self._checklist.clear()
        for item in result.get("checklist", []):
            self._checklist.add_item(item["label"], item["status"], item.get("detail", ""))

        last = result.get("last_invoice", {})
        if last:
            self._inv_preview.refresh(last)

        self._results_section.setVisible(True)
        self._open_output_btn.setVisible(True)

    def _on_op2_error(self, msg: str):
        self._run_btn.setEnabled(True)
        self._status_lbl.setText(f"Error: {msg[:80]}")
        QMessageBox.critical(self, "Operation 2 Failed", msg)

    # ── Output ────────────────────────────────────────────────────────────────

    def _open_output_folder(self):
        if os.path.exists(OUTPUT_DIR):
            os.startfile(OUTPUT_DIR)
        self.op2_complete.emit()


# ── Helper ─────────────────────────────────────────────────────────────────────

def _readonly_lbl(text: str = "—") -> QLabel:
    lbl = QLabel(text)
    lbl.setFont(_font(10))
    lbl.setMinimumHeight(34)
    lbl.setStyleSheet(
        f"background: {THEME['surface']}; color: {THEME['text_primary']}; "
        f"border: 1px solid {THEME['border']}; border-radius: 8px; padding: 5px 10px;"
    )
    return lbl
