import os
import sys
import shutil


def _frozen() -> bool:
    return getattr(sys, "frozen", False)


def _resource_root() -> str:
    """Where bundled, read-only resources live (templates).
    In dev: project folder. In frozen build: PyInstaller's _MEIPASS.
    """
    if _frozen():
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


def _user_root() -> str:
    """Where writable, persistent user data lives.
    In dev: project folder (same as resource root).
    In frozen build: next to the .exe — survives app updates.
    """
    if _frozen():
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# Roots
RESOURCE_ROOT = _resource_root()
USER_ROOT     = _user_root()
BASE_DIR      = USER_ROOT  # backwards-compat for callers expecting BASE_DIR

# Read-only bundled resources
_RESOURCE_DATA_DIR    = os.path.join(RESOURCE_ROOT, "data")
WORKING_TEMPLATE_PATH = os.path.join(_RESOURCE_DATA_DIR, "Working_Sheet.xlsx")
INVOICE_TEMPLATE_PATH = os.path.join(_RESOURCE_DATA_DIR, "Invoice_Bill.xlsx")
_BUNDLED_MRS_PATH     = os.path.join(_RESOURCE_DATA_DIR, "Master_Rate_Sheet.xlsx")

# Writable persistent data (lives next to the exe in frozen mode)
DATA_DIR         = os.path.join(USER_ROOT, "data")
OUTPUT_DIR       = os.path.join(USER_ROOT, "output")
ADDENDUM_DIR     = os.path.join(DATA_DIR, "addendums")
BACKUP_DIR       = os.path.join(DATA_DIR, "backups")
MASTER_RATE_PATH = os.path.join(DATA_DIR, "Master_Rate_Sheet.xlsx")
ADDENDUM_DATE_FORMAT = "%Y-%m-%d"

# Batch session storage
BATCH_SESSION_DIR = os.path.join(OUTPUT_DIR, "batch_sessions")
LAST_BATCH_SESSION_PATH = os.path.join(DATA_DIR, "last_batch_session.json")


def init_user_data() -> None:
    """Bootstrap persistent user-data dirs and seed MRS from bundled template if missing.

    Called once at app startup. Idempotent — safe to call any number of times.
    """
    for d in (DATA_DIR, OUTPUT_DIR, ADDENDUM_DIR, BACKUP_DIR, BATCH_SESSION_DIR):
        os.makedirs(d, exist_ok=True)

    # First-run seed: copy bundled MRS template into the writable location
    if not os.path.exists(MASTER_RATE_PATH) and os.path.exists(_BUNDLED_MRS_PATH):
        shutil.copy2(_BUNDLED_MRS_PATH, MASTER_RATE_PATH)

# Sheet names
MR_SHEET_NAME = "Master Rate Sheet"
PC_SHEET_NAME = "Party_Config"
MR_HEADER_ROW = 1  # 0-based pandas header row (Excel row 2)

# Master Rate Sheet columns
COL_TEST_CODE = "Test Code"
COL_TEST_NAME = "Test Name"
COL_MRP = "MRP"

# Party_Config columns
PC_HIS_NAME = "HIS Sheet Name"
PC_FULL_NAME = "Full Party Name"
PC_SHORT = "Party Short Code"
PC_ADDRESS = "Party Address"
PC_AGR_START = "Agreement Start Date"
PC_LAST_INV = "Last Invoice Number"
PC_REMARKS = "Remarks"

# HIS export columns to keep (8 of 22)
CSV_COLS_KEEP = [
    "S.No.",
    "Company Name",
    "Patient Name",
    "MRD Number",
    "Bill Number",
    "creation_date",
    "cpt_code",
    "Service Name",
]
CSV_COL_COMPANY = "Company Name"
CSV_COL_CPT = "cpt_code"
CSV_COL_DATE = "creation_date"
COOP_CODE = "COOP0322"

# Unmatched batch sheet columns
UNMATCHED_COL_COMPANY = "Company Name"
UNMATCHED_COL_SNO = "S.No."
UNMATCHED_COL_TEST_CODE = "Test Code"
UNMATCHED_COL_TEST_NAME = "Test Name"
UNMATCHED_COL_MRP = "MRP"
UNMATCHED_COL_DISCOUNTED = "Discounted Price"

# Working Sheet row/column positions (1-based)
WS_ROW_PARTY_NAME = 3
WS_COL_PARTY_NAME = 4   # column D
WS_COL_BILLING_PERIOD = 9  # column I
WS_ROW_HEADER = 7
WS_ROW_DATA_START = 8

WS_COL_SNO = 1
WS_COL_COMPANY = 2
WS_COL_PATIENT = 3
WS_COL_MRD = 4
WS_COL_BILL_NO = 5
WS_COL_DATE = 6
WS_COL_CPT = 7
WS_COL_SERVICE = 8
WS_COL_MRP = 9
WS_COL_DISCOUNT = 10
WS_COL_BILL_AMT = 11

# Working Sheet fill colours (ARGB)
WS_FILL_WHITE = "FFFFFFFF"
WS_FILL_BLUE = "FFD6E4F7"
WS_FILL_GREEN1 = "FFE2EFDA"
WS_FILL_GREEN2 = "FFD5ECC9"
WS_FILL_UNMATCH = "FFFFF9F5"
WS_FILL_TOTAL = "FFFFF2CC"
WS_FILL_SEPARATOR = "FFFCE4D6"

# Invoice template cells
INV_ROW_INV_NO = 7
INV_ROW_INV_DATE = 8
INV_ROW_PARTY = 11
INV_ROW_ADDR1 = 12
INV_ROW_ADDR2 = 13
INV_ROW_SERVICE = 18
INV_ROW_TOTAL = 20
INV_ROW_WORDS = 24
INV_ROW_PERIOD = 27
INV_COL_VALUE = "E"
INV_COL_AMT = "F"

# Invoice constants
PAYMENT_TERMS = "30 Days After Receipt of Invoice"
BANK_NAME = "Dhanlaxmi Bank"
BANK_ACCOUNT = "019600100024481"
BANK_IFSC = "DLXB0000196"
BANK_MICR = "110048011"
BANK_FAVOUR = "Amrita Institute of Medical Sciences & Research Centre"

# Financial year start month
FY_START_MONTH = 4  # April

# Party name mismatch map: HIS Sheet Name → Master Rate Sheet column header
PARTY_COL_MAP = {
    "Tanwar": "Tanwar Clinical Laboratory",
    "Dr.Vishwajeet Ajit Jha": "Dr. Vishwajeet Ajit Jha",
    "Bhardwaj Diagnostics": "Bhardwaj Diagnostics Center",
    "Nidhi Batra": "Dr.Nidhi Batra",
    "Manisha Gupta(Gupta Clinic)": "Manisha Gupta (Gupta Clinic)",
    "S.K.G.Hospital": "S.K.G Hospital",
    "Heritage": "HERITAGE HOSPITAL",
    "Accord superspeciality Hospital": "Accord Superspeciality",
    "Oxyvibe Hyperbaric Clinic-FBD": "Oxyvibe",
    "Dr. Manvi Gupta C/O Sachin Hospital": "Dr Manvi (Sachin Hospital)",
    "Samarpan Hospital-FBD": "Samparan Hospital",
}

# ── UI Theme ───────────────────────────────────────────────────────────────────
THEME = {
    # Primary palette
    "primary":       "#0C3A6E",  # title bar, tab active, table headers, primary buttons
    "primary_mid":   "#1A5FA8",  # accents, links, progress bar, hover states
    "primary_light": "#E6F1FB",  # info alert bg, invoice amount bg
    "primary_pale":  "#EBF3FB",  # alternating table rows, card headers, section bg

    # Success
    "success":    "#0F4F30",  # success buttons, download buttons, loaded badge
    "success_bg": "#E1F5EE",  # success alerts, loaded file zones, matched badges

    # Warning
    "warning":    "#633806",  # unmatched text, warning alert text
    "warning_bg": "#FAEEDA",  # unmatched row bg, warning alerts

    # Error
    "error":    "#8B1F1F",
    "error_bg": "#FDEAEA",

    # Info (aliases for alert usage)
    "info":    "#042C53",
    "info_bg": "#E6F1FB",

    # Neutrals
    "surface": "#F7F9FC",  # stat cards, read-only fields
    "border":  "#D8E2EE",  # all card and input borders
    "white":   "#FFFFFF",  # card backgrounds

    # Text
    "text_primary":   "#1C2833",
    "text_secondary": "#5D6D7E",
    "text_muted":     "#95A5A6",

    # Special rows
    "total_bg":     "#FFF2CC",
    "total_text":   "#1A1A00",
    "separator_bg": "#FCE4D6",
}

FONT_FAMILY = "Segoe UI"
