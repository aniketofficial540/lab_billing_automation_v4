# CLAUDE.md — Amrita Lab Billing Automation Tool
## Amrita Institute of Medical Sciences & Research Centre, Faridabad · Clinical Laboratory

**Spec version:** v3.2 (audit-synced 13-May-2026) | **FY:** 2025–26 onwards | **Platform:** Windows .exe (PyQt5, fully offline)

**What changed in v3.2 (May 2026):**
- **Batch mode:** upload one combined HIS Excel containing all companies; Op1 generates Working Sheets for every party in one run.
- **Single Unmatched Rates template:** Op1 exports one Excel (Company Name, S.No., Test Code, Test Name, MRP, Discounted Price) and saves a batch session snapshot for later Op2.
- **Op2 now uses the filled Unmatched Rates Excel** to generate invoices for all companies in the batch; invoices do not generate if any rate is missing.
- **Single-party flow removed** — the app always runs in batch mode.

**What changed in v3.1 (May 2026):**
- **Addendum upload now writes prices directly into the Master Rate Sheet** (party column) instead of staying in a JSON-only sidecar.
- New tests in the addendum that are not yet in the MRS are appended as new rows.
- Party_Config Remarks is updated with the **actual** date entered (e.g. "Addendum applied 15-Mar-2026"), while the JSON record stores the **billing date** (1st of that month) for billing-logic purposes.
- Hard-reject validation: addendum discounted price must not exceed MRP. Upload is blocked otherwise.
- `rate_engine.lookup_prices()` no longer does an addendum-override pass — prices are already in the MRS by the time Op1 runs.
- Documents window: Master Rate Sheet preview now loads all rows with a search box (CPT code / test name) and a Download MRS button.
- All MRS write paths (addendum apply, in-app edits, replace file, invoice number increment) go through atomic write + cross-process file lock + automatic backup. Last 5 backups kept under `data/backups/`. On any failure during addendum apply, the live MRS is rolled back.
- Application logging: writes to `output/app.log`, daily rotation, 14-day retention.
- PyInstaller `build.spec` ships with the project.

**What changed in v3.0:**
- Added Document Management window (third section of the app)
- Addendum is now uploaded once into Documents window and stored persistently — no addendum step in Op1 or Op2
- Unmatched rates form: employee enters MRP and Bill Amount only — Discount auto-calculated
- Addendum Excel format confirmed (Annexure-B: Sr. No., Test Code, Test Name, Discounted Price)
- New addendum replaces old one — no history kept

---

## 1. What This Project Does

Every month, Amrita Hospital's Clinical Lab raises invoices to external referring parties (hospitals, doctors, clinics) for lab tests processed that month. This desktop tool automates that process end-to-end — all parties in one batch.

Manual flow replaced: download HIS export → cross-reference rate agreement → build Excel billing summary → fill invoice → repeat for every party every month.

---

## 2. Three-Section App Structure

```
App
├── Documents        ← Setup area: Master Rate Sheet + Addendum management
├── Operation 1      ← Upload combined Excel → generate Working Sheets + Unmatched template
└── Operation 2      ← Upload filled unmatched Excel → generate all invoices
```

The Documents section is the one-time setup area. Once everything is stored there, Op1 and Op2 run with zero setup questions. Neither operation ever asks about addendums — that is handled silently from the stored data.

---

## 3. Core Design Principle — Zero Mid-Run Interruptions

> **Rule R-16:** All inputs are collected UPFRONT before each operation starts. Once "Run" is clicked, the operation executes completely without asking any questions.

The tool processes all parties in the uploaded batch. Op1 saves a batch session snapshot so Op2 can be run later after unmatched rates are filled.

---

## 4. Document Management Window (NEW in v3.0)

This window is always accessible via the tab bar. It is the setup and maintenance area for all persistent files. It has two sections.

### 4.1 Master Rate Sheet Section

**What it shows:**
- Current loaded filename, file size, date last updated
- A read-only preview table showing the full Master Rate Sheet (all 652 tests, first 6 columns) with a search box
- Search box (🔍 icon) — filters by CPT Code or Test Name in real time

**Actions available:**
- **Edit in app** button (navy) — opens the Master Rate Sheet data in an editable in-app table. Employee can change individual cell values. On Save, writes changes back to the xlsx file using openpyxl, atomically.
- **Replace file** button (secondary) — opens file dialog to upload a completely new Master Rate Sheet. Validates it has the correct structure before accepting. Replaces the stored file atomically.
- **⬇ Download MRS** button (success/green) — saves a copy of the live Master Rate Sheet to a user-chosen location.

**Storage:** `data/Master_Rate_Sheet.xlsx` — this is the live file. **All edits, replacements, and addendum applies are atomic** (temp file + os.replace) and protected by a file lock so two app instances cannot race.

**Backups:** every destructive operation snapshots the live MRS first into `data/backups/` (filename `Master_Rate_Sheet_YYYYMMDD_HHMMSS_<tag>.xlsx`). The 5 most recent backups are kept; older ones are pruned automatically.

### 4.2 Addendum Section

**What it shows:**
- A list of all parties that currently have an active addendum stored
- For each: Party Name, Short Code, Effective Date, number of tests in addendum, date uploaded
- Parties with no addendum show "No addendum" in muted text

**Actions available:**
- **Upload Addendum** button:
  - Opens file dialog for Excel file (Annexure-B format)
  - Expected columns (case-insensitive flexible matching): Sr. No., Test Code, Test Name, Discounted Price
  - Employee selects the party (dropdown of all 33), the Annexure-B Excel file, and the **Effective Date** (any day of the month).
  - Note shown: *"Any date you enter is rounded to the 1st of that month for billing. The actual date entered is recorded in Party_Config remarks."*
  - Clicking **Preview Changes →** opens an `AddendumChangesDialog` showing a per-test diff: CPT Code, Test Name, **Current MRS Price → New Price**, Difference, Action (Update / New row).
  - Clicking **Apply to Master Rate Sheet** in that dialog performs the actual write — see §4.3.
  - **Hard-reject validation:** if any test's discounted price exceeds its MRP (and MRP is known in MRS), the dialog refuses to open and lists every offending test.

- **View** button (per party) — shows the stored addendum JSON in a read-only table.
- **Delete** button (per party) — removes the JSON record after confirmation. Note: this does not roll back the MRS prices that were written.

**Storage path:** `data/addendums/addendum_[PartyShortCode].json` (JSON record only — the actual prices live in MRS)
**Config constant:** `ADDENDUM_DIR = os.path.join(DATA_DIR, "addendums")`

### 4.3 How Addendums Are Applied (v3.1)

When the employee clicks **Apply to Master Rate Sheet** in the preview dialog:

1. **Backup**: Live MRS is snapshotted into `data/backups/` (tag `pre_addendum_<short_code>`).
2. **Lock**: Cross-process advisory lock (`Master_Rate_Sheet.xlsx.lock`) is acquired so a second app instance cannot race.
3. **Atomic write**: A temp copy of MRS is opened with openpyxl. For each test in the addendum:
   - If the test code already exists in the party's column → cell is updated to the discounted price.
   - If not → a new row is appended (Test Code, Test Name, discounted price in party column; MRP left blank for employee to fill later).
4. **Remarks**: Party_Config Remarks cell for that party is set to *"Addendum applied DD-Mon-YYYY"* using the **actual** date the employee entered.
5. **JSON record**: A sidecar `addendum_[ShortCode].json` is written with `effective_date` set to the **1st of the entered month** (billing-logic anchor) plus the parsed test list and `uploaded_at` timestamp. This is for audit trail only — Op1 does not consult it.
6. **Atomic replace**: temp file is `os.replace`-d onto the live MRS path. If anything in steps 3–5 fails, the backup is restored and any partial JSON is removed.
7. **Prune**: oldest backups beyond the 5-most-recent are deleted.

The price lookup in `rate_engine.lookup_prices()` no longer reads addendum JSON — by the time Op1 runs, the MRS already contains the addendum prices.

**This replaces the old addendum step in Op2 entirely.** Op2 has no addendum section.

### 4.4 Storage Structure

```
data/
├── Master_Rate_Sheet.xlsx           ← live, all addendum prices already merged in
├── Master_Rate_Sheet.xlsx.lock      ← present only while a write is in flight
├── Working_Sheet.xlsx               ← read-only template (dev mode; in frozen builds lives in _MEIPASS)
├── Invoice_Bill.xlsx                ← read-only template (dev mode; in frozen builds lives in _MEIPASS)
├── last_batch_session.json          ← pointer to latest batch session manifest
├── addendums/                       ← JSON audit records (one per party that has had an addendum)
│   ├── addendum_Kalra.json
│   └── …
└── backups/                         ← rolling MRS snapshots (last 5 kept)
    ├── Master_Rate_Sheet_20260315_153012_pre_addendum_Kalra.xlsx
    └── …

output/
├── Kalra_Mar26_Working.xlsx         ← generated working sheets ([ShortCode]_[Mon][YY]_Working.xlsx)
├── Kalra_Mar26_Invoice.xlsx         ← generated invoices
├── app.log                          ← rotating app log (14-day retention)
└── batch_sessions/                  ← Op1 saves session here; Op2 reads it
    └── batch_YYYYMMDD_HHMMSS/
        ├── batch_session.json
        ├── matched_<short>.csv
        ├── unmatched_<short>.csv
        └── Unmatched_Rates_<sessionId>.xlsx
```

---

## 5. Operation 1 — Generate Working Sheets (Batch)

All inputs collected upfront. Once Run is clicked, zero interruptions.

| Step | What happens |
|------|-------------|
| 1 | Employee uploads the combined HIS Excel (all companies). Tool validates: correct columns present, not empty. If Master Rate Sheet not yet loaded, prompt to go to Documents tab first. |
| 2 | **Upfront party resolution (batch):** Tool splits by Company Name and matches each to Party_Config. If any company does not match, Op1 stops and lists them. |
| 3 | Employee clicks **Run**. No further input. For each company: (a) delete COOP0322 rows, (b) retain only 8 required columns, (c) lookup rates from the party column in MRS, (d) calculate Discount = MRP − Bill Amount, (e) verify K = I − J, (f) write the Working Sheet. A single Unmatched Rates template is also generated for all companies. |
| 4 | Working Sheets + Unmatched template are ready for download, and a batch session snapshot is saved for Op2. |

**Note (v3.1):** there is no longer a separate addendum-override step at lookup time. Addendum prices are already in the MRS by the time Op1 runs (written by the Documents → Upload Addendum flow). `rate_engine.lookup_prices()` is a single pass over the MRS party column.

---

## 6. Operation 2 — Upload Unmatched Rates · Invoices

All inputs collected upfront. Once Run is clicked, zero interruptions.
**No addendum step in Op2.** Addendum prices were written into the MRS at upload time (Documents tab); Op1 already consumed them.

| Step | What happens |
|------|-------------|
| 1 | **Upfront — unmatched rates file:** Employee uploads the filled Unmatched Rates Excel from Op1. Discounted Price must be filled for every row. |
| 2 | **Upfront — invoice date:** Employee sets invoice date (applies to all invoices in the batch). |
| 3 | Employee clicks **Run**. No further input. Tool executes: (a) merge confirmed rates per company, (b) recalculate totals, (c) generate final Working Sheets + Invoices for all companies, (d) increment invoice numbers in Party_Config. |
| 4 | All files are ready in `output/`. If any unmatched rate is missing or invalid, Op2 fails and lists the issues. |

### Unmatched Rates Template — Columns

The template generated by Op1 contains these columns:
- **Company Name** (Party_Config Full Party Name)
- **S.No.**
- **Test Code**
- **Test Name**
- **MRP** (from MRS)
- **Discounted Price** (to be filled by employee)

---

## 7. Input Files

### File 1 — HIS Monthly Export (Excel, combined)
- One Excel file per month containing **all companies** in the same format as the HIS CSV.
- 22 columns total. Only 8 used. Rest discarded.
- Columns N–T are Amrita-internal prices — NEVER read

| Col | Column Name | Used? | Reason |
|-----|------------|-------|--------|
| A | S.No. | YES | Carried to Working Sheet Col A |
| B | Counter Name | NO | Internal only |
| C | Company Name | YES | Party matching via HIS Sheet Name |
| D | Visit Type | NO | Internal only |
| E | Patient Name | YES | Working Sheet Col C |
| F | MRD Number | YES | Working Sheet Col D |
| G | Visit Code | NO | Internal only |
| H | Bill Number | YES | Working Sheet Col E |
| I | creation_date | YES | Working Sheet Col F |
| J | confirmed_date | NO | Internal only |
| K | cpt_code | YES | Lookup key + COOP0322 detection |
| L | Service Name | YES | Working Sheet Col H |
| M | ordered_date_time | NO | Duplicate |
| N–T | Tariff, Base Price, Discounts, etc. | NEVER | Amrita internal — never use |
| U | Payment Status | NO | Internal only |
| V | user_name | NO | Internal only |

### File 2 — Master Rate Sheet `data/Master_Rate_Sheet.xlsx`
Stored persistently in `data/`. The only source of pricing.

**'Master Rate Sheet' tab — matrix format:**
- Row 1: Title. Row 2: Column headers (pandas `header=1`). Rows 3–654: 652 tests.
- Col A: Test Code (lookup key). Col B: Test Name. Col C: MRP.
- Cols D–AJ: 33 party columns — each cell = agreed Bill Amount.
- Blank cell = test not in that party's agreement → unmatched.
- Discount never stored — calculated: `Discount = MRP − Bill Amount`.

**'Party_Config' tab — 33 rows:**

| Column | Constant | Used? | Purpose |
|--------|----------|-------|---------|
| HIS Sheet Name | `PC_HIS_NAME` | YES | HIS export company name matching |
| Full Party Name | `PC_FULL_NAME` | YES | Working Sheet + Invoice |
| Party Short Code | `PC_SHORT` | YES | Invoice number + filenames + addendum file naming |
| Party Address | `PC_ADDRESS` | YES | Invoice address |
| Agreement Start Date | `PC_AGR_START` | INFO | Reference only |
| Last Invoice Number | `PC_LAST_INV` | YES | Auto invoice numbering |
| Remarks | `PC_REMARKS` | INFO | Human notes |

### File 3 — Working Sheet Template `data/Working_Sheet.xlsx`
Read-only. Tool copies and fills. Never modified directly.

### File 4 — Invoice Template `data/Invoice_Bill.xlsx`
Read-only. Tool copies and fills. Never modified directly.

### File 5 — Addendum JSON files `data/addendums/addendum_[ShortCode].json`
One per party that has an active addendum. Written by Documents window. Read by rate_engine.

---

## 8. Working Sheet — 11-Column Structure

| Col | Name | Source | Notes |
|-----|------|--------|-------|
| A (1) | S.No. | HIS export | As-is |
| B (2) | Company Name | HIS export | As-is |
| C (3) | Patient Name | HIS export | As-is |
| D (4) | MRD No. | HIS export | As-is |
| E (5) | Bill Number | HIS export | As-is |
| F (6) | Date | HIS export | DD-MM-YYYY |
| G (7) | CPT Code | HIS export | As-is |
| H (8) | Service Name | HIS export | As-is |
| I (9) | Amrita Tariff (₹) | MRS Col C | MRP |
| J (10) | Discount (₹) | Calculated | MRP − Bill Amount |
| K (11) | Bill Amount (₹) | MRS party col | Feeds invoice total. K = I − J always. |

Fill colours (ARGB): WHITE=FFFFFFFF, BLUE=FFD6E4F7, GREEN1=FFE2EFDA, GREEN2=FFD5ECC9, UNMATCH=FFFFF9F5, TOTAL=FFFFF2CC, SEP=FFFCE4D6

---

## 9. Invoice Fields

| Field | Cell | Value |
|-------|------|-------|
| Invoice Number | E7 | `[ShortCode]-[Seq]/[FY]` e.g. `Kalra-01/25-26` |
| Invoice Date | E8 | DD.MM.YYYY — set by employee |
| Party Name | B11 | Full Party Name |
| Address Line 1 | B12 | Address split at last comma |
| Address Line 2 | B13 | Remainder (blank if no comma) |
| Service Description | B18 | `"Lab Test Service" – For the period Mar'26` |
| Amount | F18 | SUM of Col K |
| Total Amount | F20 | Same as F18 |
| Amount in Words | B24 | Indian format |
| Invoice Period | B27 | `1st Mar'26 to 31 Mar'26` |

Bank: Dhanlaxmi Bank | A/C: 019600100024481 | IFSC: DLXB0000196 | MICR: 110048011
In favour of: Amrita Institute of Medical Sciences & Research Centre
Payment Terms: 30 Days After Receipt of Invoice (not editable)

---

## 10. Invoice Numbering

Format: `[ShortCode]-[Seq]/[FY]` e.g. `Kalra-01/25-26`

FY calculation:
- Month >= April: `FY = YY-(YY+1)` where YY = last 2 digits of current year
- Month < April: `FY = (YY-1)-YY`
- May 2025 → 25-26 | March 2026 → 25-26 | April 2026 → 26-27 | May 2026 → 26-27

Sequence: read Last Invoice Number from Party_Config. Next = Last + 1, zero-padded 2 digits.
Incremented in Party_Config ONLY after BOTH files downloaded.

---

## 11. Amount in Words

`to_indian_words(amount)` uses `num2words(lang="en_IN")` + title-case.
- 23190 → "Rupees Twenty Three Thousand One Hundred And Ninety Only"
- 123456 → "Rupees One Lakh Twenty Three Thousand Four Hundred And Fifty Six Only"

---

## 12. Party Name Mapping

11 explicit mismatches in `PARTY_COL_MAP` (config.py):

| HIS Sheet Name | MRS Column Header |
|----------------|------------------|
| Tanwar | Tanwar Clinical Laboratory |
| Dr.Vishwajeet Ajit Jha | Dr. Vishwajeet Ajit Jha |
| Bhardwaj Diagnostics | Bhardwaj Diagnostics Center |
| Nidhi Batra | Dr.Nidhi Batra |
| Manisha Gupta(Gupta Clinic) | Manisha Gupta (Gupta Clinic) |
| S.K.G.Hospital | S.K.G Hospital |
| Heritage | HERITAGE HOSPITAL |
| Accord superspeciality Hospital | Accord Superspeciality |
| Oxyvibe Hyperbaric Clinic-FBD | Oxyvibe |
| Dr. Manvi Gupta C/O Sachin Hospital | Dr Manvi (Sachin Hospital) |
| Samarpan Hospital-FBD | Samparan Hospital |

Additional cases handled by case-insensitive / word-overlap fallback in `match_party()`: Paramhans Laboratory, Tomar hospital, Urban Lab, Doon hospital plus, etc.

`match_party()` resolution: exact → suffix strip → partial → collapse-space → word-overlap → None (dropdown shown).
`resolve_mr_column()` uses the same 5-pass strategy against the MRS party column headers, plus `PARTY_COL_MAP` overrides.

> **Note:** `Samarpan Hospital-FBD → Samparan Hospital` preserves a likely misspelling already present in the MRS column header. Renaming the MRS column would break this mapping — fix in one place only.

---

## 13. Business Rules

| Rule | Description |
|------|-------------|
| R-01 | All prices from MRS party column only. HIS cols N–T never read. |
| R-02 | Every COOP0322 row deleted before any processing. |
| R-03 | Only 8 HIS columns used. All others discarded. |
| R-04 | Discount = MRP − Bill Amount. Calculated at runtime. Never stored. |
| R-05 | K = I − J verified ±0.01 for every matched row. Error before output if mismatch. |
| R-06 | Unmatched tests: blank MRS cell. Bottom of Working Sheet, I/J/K blank. All must be resolved before Op2 runs. |
| R-07 | Invoice total = SUM(Col K). Never manually entered. |
| R-08 | Addendum upload writes prices directly into MRS party column. JSON sidecar saved as audit record only. |
| R-09 | New addendum replaces old one for same party. No history of JSON kept; MRS history preserved via `data/backups/`. |
| R-10 | Effective date entered by employee (any day of month) is rounded to **1st of that month** for billing logic; the actual entered date is written to Party_Config Remarks. |
| R-11 | Amount in words: Indian format. "Rupees ... Only". |
| R-12 | Invoice number: [ShortCode]-[Seq]/[FY]. **Implemented as:** Seq is incremented in MRS via `increment_invoice_no()` immediately after the Working Sheet + Invoice are written to `output/`. Increment failures are logged but do not abort the run — re-running Op2 against the same batch session would then duplicate that party's invoice number. See "Known issues" §22. |
| R-13 | Payment terms: 30 days. Standard. Not editable. |
| R-14 | MRS read with openpyxl data_only=True for lookups. Atomic temp+os.replace for writes, under cross-process file lock. |
| R-15 | 8 explicit + 4 case-insensitive party name mappings in config.py. |
| R-16 | Zero mid-run interruptions in Op1 or Op2. All inputs upfront. |
| R-17 | All MRS-mutating operations (addendum apply, in-app edit, replace, invoice-no increment) take a backup first; any failure rolls back. Last 5 backups retained. |
| R-18 | Unmatched Rates template: employee fills Discounted Price for every row. Op2 blocks if any missing or invalid. |
| R-19 | Addendum hard-reject: if any test's discounted price > MRP (and MRP is known), upload is blocked with the offending list. |
| R-20 | DataFrames crossing the QThread worker boundary are deep-copied to prevent in-place mutation between Op1 and Op2. |
| R-21 | Application logs to `output/app.log`, daily rotation, 14-day retention. Uncaught exceptions land there via sys.excepthook. |

---

## 14. File Architecture

```
lab_billing_project/
├── core/
│   ├── __init__.py
│   ├── logger.py               # NEW (v3.1): get_logger(), TimedRotatingFileHandler → output/app.log
│   ├── mrs_safety.py           # NEW (v3.1): backup_mrs(), restore_from_backup(), atomic_replace, mrs_lock, prune_backups
│   ├── batch_store.py          # NEW (v3.2): batch session manifest + latest pointer
│   ├── csv_reader.py           # read_his_export(path) → DataFrame (Excel/CSV)
│   ├── party_matcher.py        # match_party(), resolve_mr_column(), get_all_party_names()
│   ├── rate_engine.py          # lookup_prices(), apply_confirmed_rates() — single MRS pass, no addendum logic
│   ├── working_sheet_writer.py # write_working_sheet(), calculate_total()
│   ├── invoice_writer.py       # write_invoice()
│   ├── invoice_numbering.py    # get_financial_year(), get_next_invoice_no(), increment_invoice_no() — atomic + locked
│   ├── addendum_store.py       # save_addendum(), load_addendum(), delete_addendum(), list_addendums(), parse_addendum_excel()
│   ├── addendum_writer.py      # compute_addendum_preview(), validate_addendum_against_mrp(), apply_addendum() — atomic
│   ├── master_rate_editor.py   # read_mrs_for_display(), write_cells_batch(), replace_master_rate_sheet() — atomic
│   ├── amount_words.py         # to_indian_words()
│   └── validators.py           # pre_flight(), verify_bill_amounts(), validate_unmatched_rates()
├── ui/
│   ├── __init__.py
│   ├── main_window.py          # MainWindow, TitleBar, AppTabBar (3 tabs), launch()
│   ├── documents_screen.py     # DocumentsScreen, UploadAddendumDialog, AddendumChangesDialog, MRSEditorDialog
│   ├── op1_screen.py           # Op1Screen, Op1Worker(QThread)
│   ├── op2_screen.py           # Op2Screen, Op2Worker(QThread)
│   └── widgets.py              # All reusable PyQt5 widgets
├── data/                                  # Writable persistent data (next to .exe in frozen builds)
│   ├── Master_Rate_Sheet.xlsx             # Live pricing file (addendum prices already merged in)
│   ├── Working_Sheet.xlsx                 # Read-only template (bundled in dev; from _MEIPASS in frozen)
│   ├── Invoice_Bill.xlsx                  # Read-only template (bundled in dev; from _MEIPASS in frozen)
│   ├── last_batch_session.json            # Pointer to latest batch session manifest
│   ├── addendums/                         # JSON audit records
│   │   ├── addendum_Kalra.json
│   │   └── addendum_Tomar.json
│   └── backups/                           # Rolling MRS snapshots, last 5 kept
│       └── Master_Rate_Sheet_YYYYMMDD_HHMMSS_<tag>.xlsx
├── output/
│   ├── *.xlsx                             # Generated Working Sheets and Invoices
│   ├── app.log                            # Rotating app log (14-day retention)
│   └── batch_sessions/                    # Batch session snapshots + unmatched template
│       └── batch_YYYYMMDD_HHMMSS/
│           ├── batch_session.json
│           ├── matched_<short>.csv
│           ├── unmatched_<short>.csv
│           └── Unmatched_Rates_<sessionId>.xlsx
├── config.py                              # Now exports init_user_data() + RESOURCE_ROOT/USER_ROOT split
├── main.py                                # Calls init_user_data() first, then installs sys.excepthook → log
├── requirements.txt
├── CLAUDE.md
└── build.spec                             # PyInstaller config — bundles data/ templates via _MEIPASS
```

> **PyInstaller note:** `config.py` separates **RESOURCE_ROOT** (bundled, read-only — used for the three xlsx templates) from **USER_ROOT** (next to the .exe — used for `data/`, `output/`, addendums, backups, batch sessions). On first run, `init_user_data()` seeds `data/Master_Rate_Sheet.xlsx` from the bundled template if it doesn't yet exist. This means MRS edits survive .exe upgrades.

---

## 15. New Core Modules

### core/addendum_store.py

```python
# Save addendum for a party (overwrites any existing one)
save_addendum(party_short_code: str, effective_date: date, tests: list[dict]) → None
# tests format: [{"test_code": str, "test_name": str, "discounted_price": float}, ...]
# Writes to data/addendums/addendum_[party_short_code].json

# Load addendum for a party (returns None if no addendum stored)
load_addendum(party_short_code: str) → dict | None
# Returns: {"party_short_code": str, "effective_date": "YYYY-MM-DD", "tests": [...]}

# Delete addendum for a party
delete_addendum(party_short_code: str) → None

# List all stored addendums (for Documents window display)
list_addendums() → list[dict]
# Returns list of {"party_short_code", "effective_date", "test_count", "uploaded_at"}

# Parse addendum Excel file into list of test dicts
parse_addendum_excel(path: str) → list[dict]
# Reads Annexure-B Excel: Sr.No., Test Code, Test Name, Discounted Price columns
# Returns [{"test_code": str, "test_name": str, "discounted_price": float}, ...]
```

### core/rate_engine.py — single-pass MRS lookup (v3.1: addendum logic removed)

```python
def lookup_prices(his_df, mr_column, master_df, party_short_code=None, billing_month=None):
    # For each CPT code:
    # 1. Get MRP from master_df Col C
    # 2. Get bill_amount from master_df party column
    # 3. Discount = MRP - bill_amount
    # 4. Blank MRS cell = unmatched (handled in Op2 via Unmatched Rates template)
    # party_short_code and billing_month are accepted for backward compatibility
    # but no longer used; the addendum prices are already in the MRS at this point.

def apply_confirmed_rates(unmatched_df, confirmed_rates):
    # Merges Op2-confirmed MRP/Discount/Bill_Amount values into the unmatched DF.
```

### core/addendum_writer.py — preview + validate + atomic apply

```python
compute_addendum_preview(addendum_rows, mr_column, master_path) → list[change_dict]
# change_dict = {test_code, test_name, old_price, new_price, mrp, is_new_row}

validate_addendum_against_mrp(changes) → list[str]
# Returns list of error messages — empty list = OK. Hard-rejects discounted > MRP.

apply_addendum(addendum_rows, mr_column, his_name, effective_date_actual, master_path) → dict
# Atomic write. Updates existing test rows; appends new rows for unseen test codes.
# Returns {"updated": N, "added": M}.
```

### core/mrs_safety.py — atomic write, file lock, backups

```python
backup_mrs(master_path, tag="manual") → backup_path     # snapshot to data/backups/
restore_from_backup(backup_path, master_path)           # roll back
prune_backups(keep=5)                                   # housekeeping
atomic_replace(target_path)                             # context mgr: yields tmp; os.replace on exit
mrs_lock(master_path, timeout=10)                       # context mgr: cross-process advisory lock
```

### core/batch_store.py — batch session persistence (v3.2)

```python
create_batch_session_dir() -> tuple[session_dir, session_id]
save_session_manifest(session_dir: str, manifest: dict) -> session_path
load_session_manifest(path: str) -> dict
load_latest_session_path() -> str | None
```

### core/logger.py — application logging

```python
get_logger(__name__) → logging.Logger
# Configures root logger once (idempotent). File handler writes to output/app.log
# with daily rotation (14-day retention). Console handler at WARNING+.
```

### core/master_rate_editor.py

```python
# Read MRS into a list of dicts for in-app editing display
read_mrs_for_display() → dict with keys: headers, rows, party_columns

# Write a single cell edit back to the xlsx file
write_cell_edit(row_idx: int, col_name: str, new_value) → None

# Validate and replace entire MRS file
replace_master_rate_sheet(new_file_path: str) → None
# Validates structure before replacing
```

---

## 16. config.py — Complete Reference

```python
# Roots (PyInstaller-aware)
RESOURCE_ROOT     # bundled read-only resources; equals _MEIPASS in frozen builds, project dir in dev
USER_ROOT         # writable persistent root; next to .exe in frozen builds, project dir in dev
BASE_DIR = USER_ROOT  # legacy alias

# Bootstrap (called from main.py BEFORE any logger / UI import)
init_user_data()  # mkdir data/output/addendums/backups/batch_sessions; seed MRS from bundled template

# Runtime paths
DATA_DIR, OUTPUT_DIR, ADDENDUM_DIR, BACKUP_DIR
BATCH_SESSION_DIR = output/batch_sessions
LAST_BATCH_SESSION_PATH = data/last_batch_session.json

# File paths
MASTER_RATE_PATH = data/Master_Rate_Sheet.xlsx          # in USER_ROOT
WORKING_TEMPLATE_PATH = <RESOURCE_ROOT>/data/Working_Sheet.xlsx
INVOICE_TEMPLATE_PATH = <RESOURCE_ROOT>/data/Invoice_Bill.xlsx

# Sheet names
MR_SHEET_NAME = "Master Rate Sheet"
PC_SHEET_NAME = "Party_Config"
MR_HEADER_ROW = 1

# MRS columns
COL_TEST_CODE = "Test Code"
COL_TEST_NAME = "Test Name"
COL_MRP = "MRP"

# Party_Config columns
PC_HIS_NAME, PC_FULL_NAME, PC_SHORT, PC_ADDRESS, PC_AGR_START, PC_LAST_INV, PC_REMARKS

# HIS export columns
CSV_COLS_KEEP = ["S.No.", "Company Name", "Patient Name", "MRD Number",
                 "Bill Number", "creation_date", "cpt_code", "Service Name"]
CSV_COL_COMPANY, CSV_COL_CPT, CSV_COL_DATE
COOP_CODE = "COOP0322"

# Unmatched batch sheet columns
UNMATCHED_COL_COMPANY, UNMATCHED_COL_SNO, UNMATCHED_COL_TEST_CODE
UNMATCHED_COL_TEST_NAME, UNMATCHED_COL_MRP, UNMATCHED_COL_DISCOUNTED

# Working Sheet positions (1-based)
WS_ROW_PARTY_NAME=3, WS_ROW_HEADER=7, WS_ROW_DATA_START=8
WS_COL_PARTY_NAME=4 (col D, used in row 3)
WS_COL_BILLING_PERIOD=9 (col I, used in row 3)
WS_COL_SNO=1 through WS_COL_BILL_AMT=11

# Fill colours (ARGB)
WS_FILL_WHITE, WS_FILL_BLUE, WS_FILL_GREEN1, WS_FILL_GREEN2, WS_FILL_UNMATCH,
WS_FILL_TOTAL, WS_FILL_SEPARATOR   # note: SEPARATOR not SEP

# Invoice cells
INV_ROW_INV_NO=7, INV_ROW_INV_DATE=8, INV_ROW_PARTY=11, INV_ROW_ADDR1=12
INV_ROW_ADDR2=13, INV_ROW_SERVICE=18, INV_ROW_TOTAL=20, INV_ROW_WORDS=24, INV_ROW_PERIOD=27
INV_COL_VALUE="E", INV_COL_AMT="F"

# Invoice constants
PAYMENT_TERMS, BANK_NAME, BANK_ACCOUNT, BANK_IFSC, BANK_MICR, BANK_FAVOUR

# FY
FY_START_MONTH = 4

# Party mismatch map (8 entries)
PARTY_COL_MAP = { ... }

# Addendum
ADDENDUM_DATE_FORMAT = "%Y-%m-%d"

# UI Theme (deep navy / clinical white) — keys as they exist in config.THEME
THEME = {
    # Primary
    "primary":       "#0C3A6E",  "primary_mid":   "#1A5FA8",
    "primary_light": "#E6F1FB",  "primary_pale":  "#EBF3FB",
    # Status — note the `_bg` suffix (not `_light`)
    "success":       "#0F4F30",  "success_bg":    "#E1F5EE",
    "warning":       "#633806",  "warning_bg":    "#FAEEDA",
    "error":         "#8B1F1F",  "error_bg":      "#FDEAEA",
    "info":          "#042C53",  "info_bg":       "#E6F1FB",
    # Neutrals
    "surface": "#F7F9FC", "border": "#D8E2EE", "white": "#FFFFFF",
    # Text — note `_primary/_secondary/_muted` suffixes (not text/text_2/text_3)
    "text_primary": "#1C2833", "text_secondary": "#5D6D7E", "text_muted": "#95A5A6",
    # Special rows
    "total_bg":     "#FFF2CC", "total_text":   "#1A1A00",
    "separator_bg": "#FCE4D6",
}
FONT_FAMILY = "Segoe UI"
```

---

## 17. UI Architecture

**Framework:** PyQt5. All background processing via QThread workers.

**Window structure:**
```
MainWindow
├── TitleBar           — app name, subtitle, MRS loaded badge
├── AppTabBar          — [Documents] [Operation 1] [Operation 2]
└── QStackedWidget
    ├── DocumentsScreen  — MRS view/edit/replace + Addendum upload/view/delete
    ├── Op1Screen        — combined Excel upload → Run → results → unmatched template
    └── Op2Screen        — Upload unmatched Excel → invoice date → Run → checklist → output
```

**Tab bar (3 tabs):**
- Documents tab: always accessible, gear/folder icon
- Operation 1 tab: table icon, becomes "done" (green checkmark) after batch run completes
- Operation 2 tab: invoice icon, enabled after Op1 completes

**Shared app_state dict:**
```python
app_state = {
    "master_path":    str,    # path to data/Master_Rate_Sheet.xlsx
    "batch_path":     str,    # current combined HIS export path
    "batch_session_path": str, # latest batch session manifest
    "op1_result":     dict,   # set by Op1Worker
    "op2_result":     dict,   # set by Op2Worker
}
```

---

## 18. Technology Stack

| Component | Version | Purpose |
|-----------|---------|---------|
| Python | 3.11+ | Core language |
| PyQt5 | ≥5.15 | Desktop UI, QThread workers |
| pandas | ≥2.0 | CSV/Excel reading, DataFrames |
| openpyxl | ≥3.1 | Read MRS (data_only=True), write templates, write edits |
| num2words | ≥0.5.12 | Indian-format amount in words |
| json | stdlib | Addendum storage |
| PyInstaller | ≥6.0 | Package to .exe with datas=[("data/", "data/")] |

---

## 19. Key Data Facts

| Fact | Value |
|------|-------|
| Tests in MRS | 652 |
| Parties total | 33 |
| Active parties (March 2026) | 15 |
| PARTY_COL_MAP explicit entries | 8 |
| Case-insensitive fallback cases | 4 |
| Parties with 101 blank cells | 13 |
| HIS export columns total / used | 22 / 8 |
| COOP0322 rows per export | Up to 100+ |
| Addendum storage format | JSON per party |

---

## 20. Output File Naming

| Output | Format | Example |
|--------|--------|---------|
| Working Sheet | `[ShortCode]_[Mon][YY]_Working.xlsx` | `Kalra_Mar26_Working.xlsx` |
| Invoice | `[ShortCode]_[Mon][YY]_Invoice.xlsx` | `Kalra_Mar26_Invoice.xlsx` |

The `[Mon][YY]` token is the billing period inferred from the HIS export's `creation_date` column — modal year-month, formatted `%b'%y` (e.g. `Mar'26`) and then stripped of `'` and spaces. If the date column is unreadable, today's `%b'%y` is used as a fallback.

All generated files land in `output/` (one Working Sheet + one Invoice per party per run, overwriting any earlier file of the same name).

---

## 21. What the Tool Does NOT Replace

- Confirming rates for unmatched tests — employee calls the lab, enters before Op2 runs
- Reviewing and signing off outputs before dispatch
- Physically attaching the signed addendum document with the invoice
- Cash patient inclusion decisions for Dr. Prince Pareek — per agreement terms
- Verifying HIS export completeness against lab's own count (known gap — formal solution pending)

---

## 22. Known Issues (open as of 13-May-2026)

These are real spec ↔ code divergences observed during the v3.2 audit. None block the app; all need a decision from the product owner.

1. **`core/mrs_safety.restore_from_backup()` is not atomic.** Docstring says "atomically restore" but the body is a plain `shutil.copy2(backup_path, master_path)`. If the process dies mid-copy, the live MRS is corrupted. Fix: wrap in `atomic_replace(master_path)` and copy backup into the temp path.
2. **Invoice-number increment is not gated on user download.** R-12 in earlier specs said "incremented only after both files downloaded"; the code increments immediately after both files are written to disk. If the user discards the output, the seq is still consumed. If `increment_invoice_no()` itself throws, the error is logged and the loop continues — re-running Op2 against the same session will then duplicate that party's invoice number.
3. **No explicit deep-copy of DataFrames at the QThread boundary** (R-20). `Op1Worker` and `Op2Worker` pass DataFrames out via `pyqtSignal(dict)` without `.copy(deep=True)`. In practice each frame is only read by the main thread, but this is not what R-20 promises.
4. **`Master_Rate_Sheet.xlsx.lock` cleanup on hard crash.** The `mrs_lock` context manager removes the lock file on normal exit, but a process kill or power loss while inside the `with` block leaves the `.lock` behind, which then blocks the next launch until the user deletes it manually. Consider stamping PID+start-time in the lock file and treating a stale PID as releasable.
5. **PARTY_COL_MAP `Samarpan Hospital-FBD → Samparan Hospital`.** The MRS column appears to be misspelled at source (`Samparan` instead of `Samarpan`). The map preserves the misspelling deliberately — if anyone corrects the MRS column header, this mapping must change too.
6. **Bundled MRS in `_MEIPASS` is only used as a first-run seed.** Subsequent .exe upgrades will not refresh `data/Master_Rate_Sheet.xlsx` even if the bundled template contains schema changes. If the MRS structure ever changes, the upgrade must include a migration step or the user must Replace the MRS via the Documents window.
