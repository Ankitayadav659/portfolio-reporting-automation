r"""
Portfolio Report Automation
============================
Daily script: reads the source files exported from Investwell and produces
the final coloured "Our Rating / SIP / Scheme-Under / Broad-Nifty / Holding Days"
report, in the exact Master Format layout.

USAGE:
    py build_report.py --master "PortfolioValuation-XXX.xlsx" \
                        --income "Income_Report-....xlsx" \
                        --scheme-bm "Scheme_benchmark_....pdf" \
                        --broad-bm "Broad_benchmark_....pdf" \
                        --sip "My SIP Report.xlsx" \   (optional - see below)
                        --out "Portfolio_Report_OUTPUT.xlsx"

SIP FILE:
  * If --sip is given, that file is used.
  * If --sip is NOT given, the script automatically looks in the same folder for a
    SIP file named in any of the usual ways: "XSIP...", "My SIP...", "MySIP...",
    "...SIP...Registration...", or any other .xlsx whose name contains "SIP".
  * If --no-sip is given, SIP detection is skipped entirely (forces a NO SIP report).
  * If no SIP file exists at all, the SIP Amount column is left blank for every row
    and a "NO SIP" tag is written next to the report title.

RULES THAT ARE NEVER BROKEN:
  * Every value is matched by Scheme Name + Folio No together
    (except Our Rating, which is scheme-name-only, because the Income Report has no folio).
  * XSIP file (XSIP_REG_RPT_...): only Status = Active SIPs are counted.
    Pause / CXL / anything else is ignored.
  * My SIP file (My SIP's-NAME-DD-MM-YYYY-...): only SIPs of the CURRENT MONTH are
    counted (month = the date in the file name). Cancelled/paused SIPs never counted.
    If the date column cannot be identified for certain, nothing is filled.
  * If a value cannot be confidently matched, the cell is left BLANK - never 0, never guessed.
"""

import argparse
import fnmatch
import os
import re
import sys

import openpyxl
from openpyxl.styles import PatternFill
import pdfplumber


# ----------------------------------------------------------------------
# COLOUR CONSTANTS (exactly as specified in the task prompt)
# ----------------------------------------------------------------------
FILL_RATING_DEEP_PINK = PatternFill("solid", fgColor="FF66A3")   # 0.75% - <1.00%
FILL_RATING_PALE_PINK = PatternFill("solid", fgColor="FFE0EC")   # < 0.75%
FILL_CAGR_BROWN        = PatternFill("solid", fgColor="A9744C")  # 6% - <10%
FILL_CAGR_LIGHT_BROWN  = PatternFill("solid", fgColor="E3C3A0")  # < 6%
FILL_SIP_GREEN         = PatternFill("solid", fgColor="00B050")  # SIP > 0
FILL_NEGATIVE_RED      = PatternFill("solid", fgColor="FF0000")  # negative Scheme/Broad value
FILL_VIOLET            = PatternFill("solid", fgColor="EE82EE")  # whole-row override
FILL_NONE              = PatternFill(fill_type=None)

LAST_LISTS_SUMMARY = None
SCRIPT_VERSION = "v8.1 - 28-Sep-2026 (XSIP first, else My SIP; member-wise, no double counting)"

COL = {  # 1-indexed column numbers in the OUTPUT sheet
    "SCHEME": 1, "FOLIO": 2, "CURRENT_VALUE": 3, "CAGR": 4,
    "OUR_RATING": 5, "SIP": 6, "SCHEME_UNDER": 7, "BROAD_NIFTY": 8, "HOLDING_DAYS": 9,
}


# ----------------------------------------------------------------------
# SMALL SHARED HELPERS
# ----------------------------------------------------------------------
def normalise_scheme(name):
    """Scheme key: collapse spaces, lowercase. Used for exact scheme matching."""
    return re.sub(r"\s+", " ", str(name)).strip().lower()


def loose_scheme(name):
    """Punctuation-insensitive scheme key (only ever used as a safe fallback,
    and only when it is unambiguous - see SipLookup)."""
    return re.sub(r"[^a-z0-9]", "", str(name).lower())


# Word clean-up so the SAME scheme written two ways gets the same key, e.g.
#   XSIP  : "INVESCO INDIA SMALL CAP FUND - REGULAR GROWTH"
#   Master: "Invesco India Smallcap Fund Reg (G)"
# Plan type is protected: "Direct" is kept, so Direct and Regular never match.
# "Regular"/"Reg" is dropped because a scheme with no plan word is the Regular plan.
_SCHEME_WORD_MAP = {
    "g": "growth", "gr": "growth", "gro": "growth",
    "dir": "direct",
    "div": "idcw", "dividend": "idcw",
    "pru": "prudential",
    "reg": "", "regular": "",
    "fund": "", "plan": "", "option": "", "opt": "", "the": "", "scheme": "",
}


def canonical_scheme(name):
    """Order-free, spelling-tolerant scheme key. Only used as a LAST fallback for SIP
    matching, always together with an exact folio match (see SipLookup)."""
    text = clean_scheme_text(name)
    words = re.findall(r"[a-z]+|[0-9]+", text)
    out = []
    for w in words:
        if w == "cap" and out:                    # "small cap" == "smallcap"
            out[-1] = out[-1] + "cap"
            continue
        if w.isdigit() and out and not out[-1].isdigit():   # "nifty 50" == "nifty50"
            out[-1] = out[-1] + w
            continue
        w = _SCHEME_WORD_MAP.get(w, w)
        if w:
            out.append(w)
    return " ".join(sorted(out))


# ---- "starting name" matching ----------------------------------------
# Compares only the START of the scheme name and ignores the shortcuts at the end:
#   "INVESCO INDIA SMALL CAP FUND - REGULAR GROWTH"  = "Invesco India Smallcap Fund Reg (G)"
#   "ICICI PRUDENTIAL AGGRESSIVE HYBRID ACTIVE FUND OF FUNDS - GROWTH"
#                                                    = "ICICI Pru Aggressive Hybrid Active FOF (G)"
# Also tolerated: "Fund of Funds" = "FOF", an "(Erstwhile ...)" / "(Formerly ...)" old
# name at the end, and shortened words ("Pru" = "Prudential", "Act" = "Active").
# Never tolerated (these make it a DIFFERENT scheme): Direct vs Regular, Growth vs IDCW,
# a different fund house, or a different word in the name.
_PLAN_DIRECT_WORDS = {"direct", "dir"}
_OPT_IDCW_WORDS = {"idcw", "dividend", "div", "payout", "reinvestment", "reinvest", "bonus"}
_OPT_GROWTH_WORDS = {"growth", "g", "gr", "gro"}
_IGNORED_END_WORDS = {"reg", "regular", "plan", "option", "opt", "fund", "scheme", "rg"}
_END_WORDS = _PLAN_DIRECT_WORDS | _OPT_IDCW_WORDS | _OPT_GROWTH_WORDS | _IGNORED_END_WORDS

# ---- RENAMED SCHEMES (edit if ever needed) ---------------------------
# If an AMC renames a scheme and the SIP file still shows the OLD name, the script cannot
# know they are the same. Add a line here:  "name in SIP file": "name in Master file",
# (the folio must still match). Leave empty if not needed.
SCHEME_NAME_ALIASES = {
    # "ICICI PRUDENTIAL ASSET ALLOCATOR FUND (FOF) - GROWTH": "ICICI Pru Dynamic Asset Allocation Active FOF (G)",
}


def clean_scheme_text(name):
    """Lower-case text with old-name notes removed and 'Fund of Funds' -> 'fof'."""
    t = str(name).lower().replace("&", " and ")
    t = re.sub(r"\((?:\s*(?:erstwhile|formerly|earlier|old name)).*?(?:\)|$)", " ", t)
    t = re.sub(r"\b(?:erstwhile|formerly known as|formerly)\b.*$", " ", t)
    t = re.sub(r"\bfunds?\s+of\s+funds?\b", " fof ", t)
    return t


def scheme_start_parts(name):
    """Returns (start_words, plan, option, other_end_words) or None.
    plan  : 'direct' or 'regular' (no plan word = Regular)
    option: 'growth', 'idcw' or None (not written)"""
    words = []
    for w in re.findall(r"[a-z]+|[0-9]+", clean_scheme_text(name)):
        if w == "cap" and words:                                  # small cap = smallcap
            words[-1] += "cap"
        elif w.isdigit() and words and not words[-1].isdigit():   # nifty 50 = nifty50
            words[-1] += w
        else:
            words.append(w)
    if "fund" in words:
        cut = len(words) - 1 - words[::-1].index("fund")          # last "fund"
        start, end = words[:cut], words[cut + 1:]
    else:
        start, end = list(words), []
    while start and start[-1] in _END_WORDS:                      # "... fof reg g"
        end.insert(0, start.pop())
    if not start:
        return None
    plan = "direct" if any(w in _PLAN_DIRECT_WORDS for w in end) else "regular"
    if any(w in _OPT_IDCW_WORDS for w in end):
        option = "idcw"
    elif any(w in _OPT_GROWTH_WORDS for w in end):
        option = "growth"
    else:
        option = None
    others = tuple(sorted(w for w in end if w not in _END_WORDS))
    return tuple(start), plan, option, others


def _start_word_same(a, b):
    if a == b:
        return True
    if any(ch.isdigit() for ch in a + b):          # numbers must be identical
        return False
    short, long_ = (a, b) if len(a) <= len(b) else (b, a)
    return len(short) >= 3 and long_.startswith(short)   # "pru" = "prudential", "act" = "active"


def start_names_match(a, b):
    """True if two scheme names are the same scheme by the starting-name rule."""
    pa, pb = scheme_start_parts(a), scheme_start_parts(b)
    if not pa or not pb:
        return False
    (sa, plan_a, opt_a, oth_a), (sb, plan_b, opt_b, oth_b) = pa, pb
    if plan_a != plan_b:
        return False                                          # Direct never = Regular
    if opt_a and opt_b and opt_a != opt_b:
        return False                                          # Growth never = IDCW
    if oth_a != oth_b:
        return False                                          # e.g. "Institutional"
    if "".join(sa) == "".join(sb):
        return True
    if len(sa) != len(sb) or sa[0] != sb[0]:                   # fund house must be identical
        return False
    return all(_start_word_same(x, y) for x, y in zip(sa, sb))


def normalise_folio(folio):
    """Folio key: trim, remove spaces, drop a trailing '.0' that Excel adds when a folio is
    read as a number, uppercase."""
    s = re.sub(r"\s+", "", str(folio))
    if s.endswith(".0"):
        s = s[:-2]
    return s.upper()


def folios_match(a, b):
    """Same folio written in different ways:
       '26439489/82' = '26439489 / 82' = '2643948982' = '26439489' (one side without /82).
    Two folios that both have a /suffix must have the SAME suffix."""
    na, nb = normalise_folio(a), normalise_folio(b)
    if not na or not nb:
        return False
    if na == nb:
        return True
    ca, cb = re.sub(r"[^A-Z0-9]", "", na), re.sub(r"[^A-Z0-9]", "", nb)
    if ca == cb:
        return True
    sa, sb = "/" in na, "/" in nb
    return sa != sb and na.split("/")[0] == nb.split("/")[0]


def normalise_name(name):
    return re.sub(r"\s+", " ", str(name)).strip().upper()


_NAME_TITLES = {"MR", "MRS", "MS", "MISS", "DR", "SMT", "SHRI", "SHREE", "KUM", "MASTER", "BABY"}


def client_names_match(a, b):
    """Same family member written two ways?
       'RITIKA VERMA' = 'Ritika  Verma' = 'VERMA RITIKA' = 'Mrs. Ritika Verma'
       = 'RITIKA S VERMA' / 'RITIKA SUNIL VERMA' (middle initial or middle name missing on one side).
    Two different first names never match ('KIRTAN VERMA' is not 'RITIKA VERMA')."""
    wa = [w for w in re.findall(r"[A-Z]+", str(a).upper()) if w not in _NAME_TITLES]
    wb = [w for w in re.findall(r"[A-Z]+", str(b).upper()) if w not in _NAME_TITLES]
    if not wa or not wb:
        return False
    if sorted(wa) == sorted(wb):
        return True
    sa = {w for w in wa if len(w) > 1}                 # ignore single-letter initials
    sb = {w for w in wb if len(w) > 1}
    small, large = (sa, sb) if len(sa) <= len(sb) else (sb, sa)
    return len(small) >= 2 and small <= large


def normalise_header(text):
    """'CAGR (%)' / 'CAGR(%)' / 'cagr %' all become 'CAGR%' so slightly different
    header text in the export still lines up."""
    return re.sub(r"[^A-Z0-9%]", "", str(text).upper())


def to_number(value):
    """Return a float for anything that is genuinely numeric, else None (-> blank cell).
    Never invents a value: text like '-', 'NA' or '' becomes None."""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace(",", "").replace("%", "").replace("\u20b9", "")
    if s in ("", "-", "--", "NA", "N/A", "nil", "NIL"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


# ----------------------------------------------------------------------
# 1. INCOME / BROKERAGE RATE FILE  ->  lowest rate per scheme
# ----------------------------------------------------------------------
def load_brokerage_rates(path):
    """Returns dict: scheme_name_normalised -> LOWEST brokerage rate (float)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rates = {}
    header_row = None
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row and row[0] and normalise_header(row[0]) == "SCHEME":
            header_row = i
            continue
        if header_row is None:
            continue
        scheme, rate = row[0], row[1]
        if scheme is None or rate is None:
            continue
        key = normalise_scheme(scheme)
        rate = to_number(rate)
        if rate is None:
            continue
        if key not in rates or rate < rates[key]:
            rates[key] = rate
    return rates


# ----------------------------------------------------------------------
# 2. SIP REGISTRATION REPORT (optional)
#    Accepts "XSIP_REG_RPT_...", "XSIP...", "My SIP...", "...SIP...Registration..." etc.
#    Accepts .xlsx, .xlsm, .xls (real Excel 97 OR the HTML/text kind many
#    websites export), and .csv.
# ----------------------------------------------------------------------
SIP_NAME_PATTERNS = [
    "xsip_reg_rpt*",      # XSIP_REG_RPT_1002222_475774273032148.xlsx
    "xsip*",
    "*xsip*",
    "my sip*", "my_sip*", "mysip*", "my-sip*",
    "*sip*registration*",
    "*sip*report*",
    "*sip*",              # last-resort: any spreadsheet with SIP in the name
]

# Files that contain "sip" by accident, or that are our own input/output files
SIP_NAME_EXCLUDES = ("portfoliovaluation", "income", "portfolio_report", "benchmark")

SIP_FILE_EXTS = (".xlsx", ".xlsm", ".xls", ".csv")


def _sip_candidates(folder):
    try:
        names = os.listdir(folder)
    except OSError:
        return []
    candidates = []
    for name in names:
        low = name.lower()
        if low.startswith("~$") or not low.endswith(SIP_FILE_EXTS) or "sip" not in low:
            continue
        if any(bad in low for bad in SIP_NAME_EXCLUDES):
            continue
        candidates.append(name)
    return sorted(candidates)


def find_all_sip_files(folder="."):
    """ALL SIP files in `folder` - XSIP and My SIP files together (e.g. one per family
    member). Each file is later read with its own rule, and the same SIP found in two
    files is counted once (see load_sip_totals).
    If there is no XSIP / My SIP file, the single best-looking SIP file (as before)."""
    cands = _sip_candidates(folder)
    known = [n for n in cands if "xsip" in n.lower() or sip_file_mode(n) == "month"]
    if known:
        return [os.path.join(folder, n) for n in known]
    one = find_sip_file(folder)
    return [one] if one else []


def find_sip_file(folder="."):
    """Look in `folder` for a SIP registration file under any of the usual names.
    Returns the full path, or None if there isn't one."""
    try:
        names = os.listdir(folder)
    except OSError:
        return None

    candidates = []
    for name in names:
        low = name.lower()
        if low.startswith("~$"):                      # Excel temp file
            continue
        if not low.endswith(SIP_FILE_EXTS):
            continue
        if "sip" not in low:
            continue
        if any(bad in low for bad in SIP_NAME_EXCLUDES):
            continue
        candidates.append(name)

    if not candidates:
        return None

    for pattern in SIP_NAME_PATTERNS:
        for name in sorted(candidates):
            if fnmatch.fnmatch(name.lower(), pattern):
                return os.path.join(folder, name)
    return os.path.join(folder, sorted(candidates)[0])


# ---- reading the file, whatever format it is in ----------------------
class _HTMLTableReader:
    """Tiny HTML table reader (no extra installs needed). Many 'Excel' downloads
    from websites are really HTML tables saved with a .xls name."""

    def __init__(self):
        from html.parser import HTMLParser

        outer = self
        self.rows, self._row, self._cell, self._in_cell = [], None, [], False

        class P(HTMLParser):
            def handle_starttag(self, tag, attrs):
                if tag == "tr":
                    outer._row = []
                elif tag in ("td", "th"):
                    outer._in_cell, outer._cell = True, []
                elif tag == "br" and outer._in_cell:
                    outer._cell.append(" ")

            def handle_endtag(self, tag):
                if tag in ("td", "th") and outer._row is not None:
                    outer._row.append(re.sub(r"\s+", " ", "".join(outer._cell)).strip() or None)
                    outer._in_cell = False
                elif tag == "tr" and outer._row is not None:
                    outer.rows.append(outer._row)
                    outer._row = None

            def handle_data(self, data):
                if outer._in_cell:
                    outer._cell.append(data)

        self.parser = P()

    def read(self, text):
        self.parser.feed(text)
        return self.rows


def _read_text(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "utf-16", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1", errors="replace")


def read_sheet_rows(path):
    """Return a list of 'sheets', each a list of rows (lists of cell values),
    for .xlsx / .xlsm / .xls / .csv files. Raises a clear error if unreadable."""
    import csv

    low = path.lower()
    with open(path, "rb") as fh:
        head = fh.read(8)

    # Real modern Excel (.xlsx / .xlsm) - it's a zip file, starts with 'PK'
    if head.startswith(b"PK"):
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        return [[list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets]

    # Real old Excel 97-2003 (.xls) binary
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        try:
            import xlrd
        except ImportError:
            raise RuntimeError(
                "The SIP file is in the OLD Excel format (.xls).\n"
                "  Fix (pick one):\n"
                "   a) Open it in Excel -> File -> Save As -> 'Excel Workbook (*.xlsx)', then run again, OR\n"
                "   b) One-time install: open Command Prompt and type:  py -m pip install xlrd")
        book = xlrd.open_workbook(path)
        return [[sh.row_values(i) for i in range(sh.nrows)] for sh in book.sheets()]

    # Otherwise it's text: an HTML table, CSV, or tab-separated file
    text = _read_text(path)
    if "<table" in text.lower():
        return [_HTMLTableReader().read(text)]

    sample = text[:5000]
    delim = "\t" if sample.count("\t") > sample.count(",") else ","
    if low.endswith(".csv") and sample.count(";") > sample.count(","):
        delim = ";"
    rows = [[(c.strip() or None) for c in r] for r in csv.reader(text.splitlines(), delimiter=delim)]
    return [rows]


# ---- header recognition ----------------------------------------------
SIP_SCHEME_HEADERS = ("SCHEMENAME", "SCHEME", "SCHEMEDESCRIPTION", "FUNDNAME", "SCHEMENAMEPLAN")
SIP_FOLIO_HEADERS = ("FOLIONO", "FOLIONUMBER", "FOLIO", "FOLIONUM")
SIP_STATUS_HEADERS = ("STATUS", "SIPSTATUS", "XSIPSTATUS", "REGISTRATIONSTATUS", "CURRENTSTATUS")
SIP_AMOUNT_HEADERS = ("INSTALLMENTAMOUNT", "INSTALMENTAMOUNT", "SIPAMOUNT", "XSIPAMOUNT",
                      "INSTALLMENTAMT", "INSTALMENTAMT", "AMOUNT")
# Family member's name (exact header only - "Scheme Name" must never be taken for it)
SIP_CLIENT_HEADERS = ("CLIENTNAME", "INVESTORNAME", "FIRSTHOLDERNAME", "PRIMARYHOLDERNAME",
                      "HOLDERNAME", "APPLICANTNAME", "INVESTOR", "CLIENT")

# A header containing any of these words is never the column we want
# (e.g. "Scheme Code" is not the scheme NAME; "Folio Type" is not the folio number).
_HEADER_BLOCKLIST = ("CODE", "TYPE", "DATE", "ISIN", "AMC", "CATEGORY", "MIN", "MAX", "TOTAL", "NO.OF", "NOOF")


def _find_header_index(header_map, wanted):
    """Exact header match first; then a careful 'contains' match as a backup."""
    for w in wanted:
        if w in header_map:
            return header_map[w]
    for w in wanted:
        for key, idx in header_map.items():
            if w in key and not any(b in key for b in _HEADER_BLOCKLIST):
                return idx
    return None


class SipLookup:
    """All counted SIPs from the SIP file, matched to Master rows by Scheme + Folio.

    .get(scheme, folio) returns the summed SIP amount, or None if there is no confident
    match (-> the cell is left blank, never 0).

    Matching steps, tried in this order - each one still needs the SAME folio
    (folio may be written '26439489/82', '26439489', '2643948982' ...):
      1. renamed-scheme list (SCHEME_NAME_ALIASES)
      2. exact name
      3. same name ignoring punctuation / spacing
      4. same name after word clean-up (Small Cap = Smallcap, Reg (G) = Regular Growth ...)
      5. same STARTING name, shortcuts at the end ignored (Direct/Regular and Growth/IDCW
         must still agree)
    A step is only accepted if it points to exactly ONE scheme in the SIP file and exactly
    ONE row in the Master file for that folio - otherwise the cell is left blank.

    FAMILY MEMBERS: when the SIP file has a Client Name column, a SIP is only counted on
    the Master rows of THAT member (Client Name must match the member heading in the
    Master file, e.g. 'RITIKA VERMA (CMGPS9214Q)'). Same scheme + folio but a different
    member's name -> left blank and listed in name_mismatch, never guessed.
    """

    LEVELS = ("alias", "exact", "loose", "clean-up", "starting name")

    def __init__(self):
        self.entries = []        # counted SIP rows: dict(name, folio, amount, id)
        self.skipped = []        # NOT counted rows: dict(name, folio, amount, reason)
        self.master_pairs = []   # every (scheme, folio) in the Master file
        self.name_fixed = []     # (SIP-file name, Master name, folio, amount, how)
        self.used = set()        # entry ids that ended up in the report
        self.last_count = 0      # how many SIPs the most recent get() matched
        self.active_rows = 0
        self.skipped_rows = 0
        self.status_counts = {}
        self.columns_found = {}
        self.mode = "active"          # "active" (XSIP) or "month" (My SIP)
        self.month_label = None       # e.g. "September 2026" (My SIP only)
        self.month_source = None
        self.date_rule_text = None
        self.all_headers = None
        self.files = []               # SIP files read (base names)
        self.file_clients = {}        # file name -> set of Client Names in it
        self.name_mismatch = []       # (SIP client, Master member, scheme, folio, amount)
        self._mismatch_ids = set()
        self.file_info = {}           # file name -> dict(kind, rule, counted, skipped, rank)
        self.errors = []              # files (or parts) that could not be used, with the reason
        self.duplicate_files = []     # (file name, same content as this earlier file)
        self.replaced = []            # (SIP entry, newer file that has the same SIP)
        self.placed = []              # (member, scheme, folio, amount, note) rows filled in the report
        self._aliases = {normalise_scheme(k): v for k, v in SCHEME_NAME_ALIASES.items()}

    # ---------------- loading ----------------
    @staticmethod
    def _client_txt(client):
        return normalise_name(client) if client not in (None, "") and str(client).strip() else None

    def add(self, scheme, folio, amount, client=None):
        self.entries.append({"name": str(scheme).strip(), "folio": str(folio).strip(),
                             "amount": amount, "id": len(self.entries),
                             "client": self._client_txt(client)})
        self.active_rows += 1

    def add_skipped(self, scheme, folio, amount, reason, client=None):
        self.skipped.append({"name": str(scheme).strip(), "folio": str(folio).strip(),
                             "amount": amount, "reason": reason, "id": -1,
                             "client": self._client_txt(client)})

    def set_master_rows(self, rows):
        """rows: (scheme, folio) or (scheme, folio, member name)"""
        self.master_pairs = [(str(t[0]), str(t[1]), t[2] if len(t) > 2 else None) for t in rows]

    @staticmethod
    def _client_ok(entry_client, member):
        """No name on either side -> folio decides (old behaviour)."""
        if not entry_client or not member:
            return True
        return client_names_match(entry_client, member)

    # ---------------- matching ----------------
    def _same(self, level, sip_name, master_name):
        if level == "alias":
            target = self._aliases.get(normalise_scheme(sip_name))
            return target is not None and normalise_scheme(target) == normalise_scheme(master_name)
        if level == "exact":
            return normalise_scheme(sip_name) == normalise_scheme(master_name)
        if level == "loose":
            return loose_scheme(sip_name) == loose_scheme(master_name)
        if level == "clean-up":
            return canonical_scheme(sip_name) == canonical_scheme(master_name)
        return start_names_match(sip_name, master_name)

    def _match(self, scheme, folio, pool, member=None):
        """Returns (hits, level) or ([], None)."""
        cands = [e for e in pool if folios_match(e["folio"], folio)
                 and self._client_ok(e.get("client"), member)]
        if not cands:
            return [], None
        masters_same_folio = [m for m in self.master_pairs if folios_match(m[1], folio)]
        for level in self.LEVELS:
            hits = [e for e in cands if self._same(level, e["name"], scheme)]
            if not hits:
                continue
            if len({normalise_scheme(e["name"]) for e in hits}) != 1:
                return [], None               # two different SIP schemes fit -> don't guess
            rival_rows = [m for m in masters_same_folio if self._same(level, hits[0]["name"], m[0])
                          and self._client_ok(hits[0].get("client"), m[2])]
            if len(rival_rows) > 1:
                return [], None               # two Master rows would claim it -> don't guess
            return hits, level
        return [], None

    def get(self, scheme, folio, member=None):
        self.last_count = 0
        hits, level = self._match(scheme, folio, self.entries, member)
        if not hits:
            if member:
                # Same scheme + folio exists, but under ANOTHER member's name -> blank, report it
                other, _ = self._match(scheme, folio, self.entries)
                other = [e for e in other if e["id"] not in self.used]
                if other:
                    self._mismatch_ids.update(e["id"] for e in other)
                    self.name_mismatch.append((other[0]["client"], normalise_name(member),
                                               str(scheme).strip(), str(folio).strip(),
                                               sum(e["amount"] for e in other)))
            return None
        total = sum(e["amount"] for e in hits)
        self.used.update(e["id"] for e in hits)
        self.last_count = len(hits)
        self.placed.append((normalise_name(member) if member else "", str(scheme).strip(),
                            str(folio).strip(), total,
                            "; ".join(sorted({e["note"] for e in hits if e.get("note")}))))
        if level not in ("exact",) or any(normalise_folio(e["folio"]) != normalise_folio(folio)
                                          for e in hits):
            self.name_fixed.append((hits[0]["name"], str(scheme).strip(), str(folio).strip(),
                                    total, level, hits[0]["folio"]))
        return total

    def explain_blank(self, scheme, folio, member=None):
        """If the SIP file HAS this scheme+folio but it was not counted, say why."""
        for sip_client, m, s, f, _amt in self.name_mismatch:
            if s == str(scheme).strip() and f == str(folio).strip() and \
                    (not member or m == normalise_name(member)):
                return f"registered under client '{sip_client}', not '{m}'"
        hits, _ = self._match(scheme, folio, self.skipped, member)
        if not hits:
            return None
        reasons = sorted({e["reason"] for e in hits})
        return "; ".join(reasons)

    def unused(self):
        """Counted SIPs from the file that did not land on any report row
        (name mismatches are reported separately in name_mismatch)."""
        out = {}
        for e in self.entries:
            if e["id"] in self.used or e["id"] in self._mismatch_ids:
                continue
            k = (normalise_scheme(e["name"]), normalise_folio(e["folio"]))
            if k in out:
                out[k] = (out[k][0], out[k][1], out[k][2] + e["amount"])
            else:
                out[k] = (e["name"], e["folio"], e["amount"])
        return list(out.values())

    def same_folio_master_rows(self, folio):
        return [m[0] for m in self.master_pairs if folios_match(m[1], folio)]

    # ---------------- same SIP in two files -> counted once ----------------
    def _same_sip(self, a, b):
        if not self._client_ok(a.get("client"), b.get("client")):
            return False
        if not folios_match(a["folio"], b["folio"]):
            return False
        return (normalise_scheme(a["name"]) == normalise_scheme(b["name"])
                or canonical_scheme(a["name"]) == canonical_scheme(b["name"])
                or start_names_match(a["name"], b["name"]))

    def resolve_duplicates(self):
        """The same member + scheme + folio found in TWO DIFFERENT files is counted ONCE -
        amounts are never added across files. Which file wins:
          1. an XSIP file always beats a My SIP file (XSIP has priority)
          2. between two files of the same type, the newest file wins
        If the XSIP file has that SIP but NOT Active (Pause / CXL ...), the My SIP file
        cannot bring it back - XSIP decides.
        A SIP that is only in the My SIP file (member / scheme not in any XSIP file) is used.
        Two SIPs on the same scheme inside ONE file are still added up (two registrations)."""
        kind = {f: info.get("kind") for f, info in self.file_info.items()}
        prio = {f: (1 if info.get("kind") == "XSIP" else 0, info.get("rank", 0))
                for f, info in self.file_info.items()}
        best_first = sorted(self.entries, key=lambda e: prio.get(e.get("file"), (0, 0)), reverse=True)
        xsip_off = [s for s in self.skipped if kind.get(s.get("file")) == "XSIP"
                    and str(s.get("reason", "")).startswith("Status")]
        dropped = set()
        kept = []
        for e in best_first:
            dup = next((k for k in kept if k.get("file") != e.get("file") and self._same_sip(k, e)),
                       None)
            off = None
            if dup is None and kind.get(e.get("file")) != "XSIP":
                off = next((s for s in xsip_off if self._same_sip(s, e)), None)
            if dup is None and off is None:
                kept.append(e)
                continue
            dropped.add(id(e))
            if off is not None:
                self.replaced.append((e, off["file"]))
                reason = (f"XSIP file '{off['file']}' says {off['reason']} - XSIP has priority, "
                          f"not counted")
            else:
                self.replaced.append((e, dup["file"]))
                why = ("XSIP file has priority" if kind.get(dup["file"]) == "XSIP"
                       and kind.get(e.get("file")) != "XSIP" else "newer file")
                reason = (f"same SIP is in '{dup['file']}' ({why}) - that amount is used, "
                          f"not added twice")
            self.skipped.append(dict(e, id=-1, reason=reason))
            self.active_rows -= 1
            if e.get("file") in self.file_info:
                self.file_info[e["file"]]["counted"] -= 1
                self.file_info[e["file"]]["replaced"] += 1
        self.entries = [e for e in self.entries if id(e) not in dropped]
        for i, e in enumerate(self.entries):
            e["id"] = i

    def summary(self):
        """After build_output: what happened to the uploaded SIPs."""
        return {
            "updated": len(self.placed),
            "updated_amount": sum(p[3] for p in self.placed),
            "unmatched": len(self.unused()) + len(self.name_mismatch),
            "errors": len(self.errors),
            "not_counted": len(self.skipped),
            "duplicates": len(self.duplicate_files) + len(self.replaced),
        }

    def __bool__(self):
        return bool(self.entries)

    def __len__(self):
        return len({(normalise_scheme(e["name"]), normalise_folio(e["folio"]))
                    for e in self.entries})


# ---- which rule applies to which SIP file ---------------------------
#   XSIP file   (XSIP_REG_RPT_...)      -> only rows with Status = Active
#   My SIP file (My SIP's-NAME-DD-MM-YYYY-...) -> only SIPs of the CURRENT MONTH
#                                          (month taken from the date in the file name)
def sip_file_mode(path):
    name = os.path.basename(str(path)).lower()
    name = re.sub(r"[_\-.']", " ", name)
    if re.search(r"\bmy\s*sip", name):
        return "month"
    return "active"


MONTH_NAMES = ["", "January", "February", "March", "April", "May", "June", "July",
               "August", "September", "October", "November", "December"]


def month_from_filename(path):
    """'My SIP's-NISREEN H BADSHAH-24-09-2026-03-37-32.xlsx' -> (2026, 9).
    Returns None if the name has no DD-MM-YYYY date in it."""
    m = re.search(r"(?<!\d)(\d{1,2})[-_.](\d{1,2})[-_.](\d{4})(?!\d)", os.path.basename(str(path)))
    if not m:
        return None
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if 1 <= d <= 31 and 1 <= mo <= 12 and 2000 <= y <= 2100:
        return (y, mo)
    return None


def parse_any_date(v):
    """Turn a cell into a date, or None. Indian day-first format is assumed
    (24-09-2026 = 24 September). Anything unclear -> None (row is not counted)."""
    import datetime as dt
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)):
        if 20000 <= v <= 80000:                      # Excel serial date number
            return dt.date(1899, 12, 30) + dt.timedelta(days=int(v))
        return None
    s = re.sub(r"\s+", " ", str(v)).strip()
    if not s:
        return None
    candidates = [s, s.split(" ")[0]]
    fmts = ("%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y", "%d-%b-%Y", "%d %b %Y", "%d-%B-%Y",
            "%d %B %Y", "%d/%b/%Y", "%Y-%m-%d", "%Y/%m/%d", "%d-%m-%y", "%d/%m/%y",
            "%d-%b-%y", "%Y-%m-%d %H:%M:%S", "%d-%m-%Y %H:%M:%S", "%d/%m/%Y %H:%M:%S")
    for c in candidates:
        for f in fmts:
            try:
                return dt.datetime.strptime(c, f).date()
            except ValueError:
                continue
    return None


# Date columns that describe WHEN an instalment falls (checked in this order)
SIP_POINT_DATE_HEADERS = ("NEXTINSTALLMENTDATE", "NEXTINSTALMENTDATE", "NEXTSIPDATE",
                          "NEXTDUEDATE", "INSTALLMENTDATE", "INSTALMENTDATE", "SIPDUEDATE",
                          "DUEDATE", "SIPDATE", "TRANSACTIONDATE", "TRXNDATE", "TXNDATE")
SIP_START_HEADERS = ("SIPSTARTDATE", "STARTDATE", "FROMDATE")
SIP_END_HEADERS = ("SIPENDDATE", "ENDDATE", "TODATE", "CEASEDATE")

# In the My SIP file, a SIP with one of these statuses is never counted,
# even if its date falls in the current month.
SIP_STOPPED_STATUSES = ("CXL", "CANCEL", "PAUSE", "TERMINAT", "CEASE", "STOP", "CLOSE",
                        "INACTIVE", "REJECT", "EXPIRE", "FAIL", "REDEEM")


def _exact_header(header_map, wanted):
    for w in wanted:
        if w in header_map:
            return header_map[w]
    return None


# ---- My SIP export with one column per month (" JUN 26 | JUL 26 | AUG 26 | SEP 26") ----
_MONTH_ABBR = {m[:3].upper(): i for i, m in enumerate(MONTH_NAMES) if m}
_MONTH_COL_RE = re.compile(r"^\s*([A-Za-z]{3})[A-Za-z]*[\s\-'.]*(\d{2}|\d{4})\s*$")


def month_column(header):
    """' SEP 26' / 'Sep-2026' / 'SEPT 26' -> (2026, 9); anything else -> None."""
    if header is None or isinstance(header, (int, float)):
        return None
    m = _MONTH_COL_RE.match(str(header))
    if not m or m.group(1).upper() not in _MONTH_ABBR:
        return None
    y = int(m.group(2))
    return (2000 + y if y < 100 else y, _MONTH_ABBR[m.group(1).upper()])


def month_label_short(ym):
    return f"{MONTH_NAMES[ym[1]][:3].upper()} {ym[0] % 100:02d}"


def _find_sip_header(rows):
    """(row index, normalised-header -> column map, raw header row) or None."""
    for i, row in enumerate(rows):
        if not row:
            continue
        cand = {}
        for idx, v in enumerate(row):
            if v not in (None, ""):
                cand.setdefault(normalise_header(v), idx)
        if cand and _find_header_index(cand, SIP_SCHEME_HEADERS) is not None and \
                _find_header_index(cand, SIP_FOLIO_HEADERS) is not None:
            return i, cand, row
    return None


def detect_sip_kind(path):
    """Which parser a SIP file needs - by file name first, then by its column headers.
      'xsip'        XSIP registration report    -> Status = ACTIVE only
      'mysip_grid'  My SIP with month columns   -> current month (else last month)
      'mysip_dates' My SIP with date columns    -> SIP date falls in the current month
    A file that is neither is read with the XSIP rule (Status = Active), as before."""
    name = re.sub(r"[_\-.']", " ", os.path.basename(str(path)).lower())
    if re.search(r"\bx\s*sip", name):
        return "xsip"
    for rows in read_sheet_rows(path):
        found = _find_sip_header(rows)
        if found and any(month_column(h) for h in found[2]):
            return "mysip_grid"
        if found and not re.search(r"\bmy\s*sip", name):
            hmap = found[1]
            if _find_header_index(hmap, SIP_STATUS_HEADERS) is not None:
                return "xsip"
    return "mysip_dates" if re.search(r"\bmy\s*sip", name) else "xsip"


def file_datetime(path):
    """For 'newest file wins': the date-time in the name
    (My SIP's-NAME-26-09-2026-06-00-38.xlsx), else the file's modified time."""
    import datetime as dt
    m = re.search(r"(\d{1,2})-(\d{1,2})-(\d{4})(?:-(\d{1,2})-(\d{1,2})-(\d{1,2}))?",
                  os.path.basename(str(path)))
    if m:
        try:
            parts = [int(x) if x else 0 for x in m.groups()]
            return dt.datetime(parts[2], parts[1], parts[0], parts[3], parts[4], parts[5]).timestamp()
        except ValueError:
            pass
    try:
        return os.path.getmtime(path)
    except OSError:
        return 0


def _load_mysip_grid(path, lookup):
    """My SIP file with one amount column per month.
    Amount used = CURRENT month column (month of the date in the file name, else the
    'as on' date in the title, else today). If that is 0 / blank -> LAST month's column
    only (never further back). Both 0 / blank -> not counted.
    Client = INVESTOR column (each family member's own name)."""
    import datetime as dt
    counted = skipped = 0
    rule = None
    for rows in read_sheet_rows(path):
        found = _find_sip_header(rows)
        if not found:
            continue
        h_idx, hmap, header_row = found
        months = {month_column(h): i for i, h in enumerate(header_row) if month_column(h)}
        if not months:
            continue

        target = month_from_filename(path)
        if target is None:
            for row in rows[:h_idx]:
                for v in row or []:
                    m = re.search(r"as\s+on\s+(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", str(v or ""), re.I)
                    if m:
                        target = (int(m.group(3)), int(m.group(2)))
        if target is None:
            today = dt.date.today()
            target = (today.year, today.month)
        prev = (target[0], target[1] - 1) if target[1] > 1 else (target[0] - 1, 12)
        i_cur, i_prev = months.get(target), months.get(prev)
        cur_lbl, prev_lbl = month_label_short(target), month_label_short(prev)
        lookup.month_label = f"{MONTH_NAMES[target[1]]} {target[0]}"
        if i_cur is None:
            lookup.errors.append(f"{os.path.basename(str(path))}: no {cur_lbl} column "
                                 f"(columns: {', '.join(month_label_short(k) for k in sorted(months))})"
                                 f" - nothing used from this file")
            continue
        rule = f"{cur_lbl} amount" + (f" (if 0 -> {prev_lbl})" if i_prev is not None else "")

        i_scheme = _find_header_index(hmap, SIP_SCHEME_HEADERS)
        i_folio = _find_header_index(hmap, SIP_FOLIO_HEADERS)
        i_status = _find_header_index(hmap, SIP_STATUS_HEADERS)
        i_client = _exact_header(hmap, SIP_CLIENT_HEADERS)
        header_client = None                     # 'Client: SHEETAL KAUSHAL PARIKH (AIUPP5990J)'
        for row in rows[:h_idx]:
            for v in row or []:
                m = re.match(r"\s*client\s*:\s*(.+?)\s*(\([A-Z0-9]+\))?\s*$", str(v or ""), re.I)
                if m:
                    header_client = m.group(1)
        lookup.columns_found = {
            "Scheme": header_row[i_scheme], "Folio": header_row[i_folio],
            "Status": header_row[i_status] if i_status is not None else "(none - My SIP)",
            "Amount": rule,
            "Client": header_row[i_client] if i_client is not None else "(from 'Client:' line)",
        }

        def cell(row, i):
            return row[i] if (i is not None and i < len(row)) else None

        for row in rows[h_idx + 1:]:
            if not row:
                continue
            first = next((str(v).strip().lower() for v in row if v not in (None, "")), "")
            if first.startswith("total"):
                continue
            scheme, folio = cell(row, i_scheme), cell(row, i_folio)
            if not scheme or not folio:
                continue
            client = cell(row, i_client) or header_client
            status = cell(row, i_status)
            status_txt = str(status).strip().upper() if status not in (None, "") else ""
            if status_txt and any(w in status_txt for w in SIP_STOPPED_STATUSES):
                lookup.add_skipped(scheme, folio, None, f"Status = {status_txt}", client)
                skipped += 1
                continue
            cur, old = to_number(cell(row, i_cur)), to_number(cell(row, i_prev))
            if cur is not None and cur > 0:
                lookup.add(scheme, folio, cur, client)
            elif old is not None and old > 0:
                lookup.add(scheme, folio, old, client)
                lookup.entries[-1]["note"] = f"{prev_lbl} amount used ({cur_lbl} was 0)"
            else:
                why = f"0 in {cur_lbl}" + (f" and {prev_lbl}" if i_prev is not None else "")
                lookup.add_skipped(scheme, folio, 0, f"{why} - no SIP this month", client)
                skipped += 1
                continue
            counted += 1
    lookup.skipped_rows += skipped
    return rule


def load_sip_totals(paths):
    """Read ONE or SEVERAL SIP files into a single SipLookup - XSIP files, My SIP files,
    or both together (e.g. one file per family member). `paths`: one path or a list.
      * each file is read with its own rule (see detect_sip_kind)
      * the same file uploaded twice (same content) is read once
      * the same member + scheme + folio in two different files -> XSIP beats My SIP,
        else the newest file; never added twice (resolve_duplicates)
      * a file that cannot be read is listed in lookup.errors; the other files still work
    """
    import hashlib
    lookup = SipLookup()
    if not paths:
        return lookup
    if isinstance(paths, (str, os.PathLike)):
        paths = [paths]
    seen = {}
    kinds = set()
    for p in paths:
        fname = os.path.basename(str(p))
        try:
            with open(p, "rb") as fh:
                digest = hashlib.sha1(fh.read()).hexdigest()
        except OSError as e:
            lookup.errors.append(f"{fname}: could not open ({e.strerror or e})")
            continue
        if digest in seen:
            lookup.duplicate_files.append((fname, seen[digest]))
            continue
        seen[digest] = fname

        before = {id(e) for e in lookup.entries + lookup.skipped}
        n_entries, n_skipped, n_errors = len(lookup.entries), len(lookup.skipped), len(lookup.errors)
        try:
            kind = detect_sip_kind(p)
            if kind == "mysip_grid":
                rule = _load_mysip_grid(p, lookup)
                label = "My SIP"
            else:
                _load_one_sip_file(p, lookup, "month" if kind == "mysip_dates" else "active")
                label = "My SIP" if kind == "mysip_dates" else "XSIP"
                rule = ("Status = ACTIVE" if kind == "xsip" else
                        (lookup.date_rule_text or "could not tell which column holds the SIP date"))
        except Exception as e:                   # this file is unusable - keep the others
            del lookup.entries[n_entries:]
            del lookup.skipped[n_skipped:]
            lookup.errors.append(f"{fname}: could not read - {str(e).splitlines()[0] if str(e) else type(e).__name__}")
            continue
        new = [e for e in lookup.entries + lookup.skipped if id(e) not in before]
        for e in new:
            e["file"] = fname
        counted = sum(1 for e in lookup.entries[n_entries:])
        if not new and len(lookup.errors) == n_errors:
            if kind == "mysip_dates" and lookup.all_headers:
                lookup.errors.append(f"{fname}: could not tell which column holds the SIP date/month "
                                     f"(columns: {' | '.join(str(h) for h in lookup.all_headers)})")
            else:
                lookup.errors.append(f"{fname}: no rows with Scheme + Folio found - is this a SIP report?")
        kinds.add(label)
        lookup.files.append(fname)
        lookup.file_clients[fname] = {e["client"] for e in new if e["client"]}
        lookup.file_info[fname] = {"kind": label, "rule": rule, "counted": counted,
                                   "skipped": len(lookup.skipped) - n_skipped, "replaced": 0,
                                   "rank": file_datetime(p)}
    lookup.mode = ("active" if kinds == {"XSIP"} else "month" if kinds == {"My SIP"}
                   else "mixed" if kinds else "active")
    lookup.resolve_duplicates()
    return lookup


def _load_one_sip_file(path, lookup, mode=None):
    """Read one SIP file INTO `lookup` (see load_sip_totals).

    XSIP file   : Status == 'Active' only (Pause / CXL / anything else ignored).
    My SIP file : only SIPs that fall in the CURRENT MONTH (the month in the file
                  name, e.g. 24-09-2026 -> September 2026). If the file also has a
                  Status column, cancelled/paused/stopped SIPs are still ignored.
    Both        : matched by Scheme Name + Folio No together; several matching SIPs
                  on the same Scheme+Folio are added up; no match -> blank (never 0).
    If the script cannot tell for sure which column holds the date, it fills
    NOTHING and says so, rather than guessing.
    """
    import calendar
    import datetime as dt

    mode = mode or sip_file_mode(path)
    lookup.mode = mode

    target = None
    if mode == "month":
        target = month_from_filename(path)
        if target is None:
            today = dt.date.today()
            target = (today.year, today.month)
            lookup.month_source = "today's date (no date found in the file name)"
        else:
            lookup.month_source = "the date in the file name"
        lookup.month_label = f"{MONTH_NAMES[target[1]]} {target[0]}"
        month_first = dt.date(target[0], target[1], 1)
        month_last = dt.date(target[0], target[1], calendar.monthrange(*target)[1])

    for rows in read_sheet_rows(path):
        header = None
        data = []
        for row in rows:
            if not row:
                continue
            if header is None:
                candidate = {}
                for idx, v in enumerate(row):
                    if v not in (None, ""):
                        candidate.setdefault(normalise_header(v), idx)
                if not candidate:
                    continue
                if (_find_header_index(candidate, SIP_SCHEME_HEADERS) is None or
                        _find_header_index(candidate, SIP_FOLIO_HEADERS) is None):
                    continue
                header = (candidate, row)
                continue
            data.append(row)

        if header is None:
            continue
        hmap, header_row = header
        i_scheme = _find_header_index(hmap, SIP_SCHEME_HEADERS)
        i_folio = _find_header_index(hmap, SIP_FOLIO_HEADERS)
        i_status = _find_header_index(hmap, SIP_STATUS_HEADERS)
        i_amount = _find_header_index(hmap, SIP_AMOUNT_HEADERS)
        i_client = _exact_header(hmap, SIP_CLIENT_HEADERS)

        def cell(row, i):
            return row[i] if (i is not None and i < len(row)) else None

        lookup.columns_found = {
            "Scheme": header_row[i_scheme],
            "Folio": header_row[i_folio],
            "Status": header_row[i_status] if i_status is not None else "(NOT FOUND)",
            "Amount": header_row[i_amount] if i_amount is not None else "(NOT FOUND)",
            "Client": header_row[i_client] if i_client is not None else "(NOT FOUND - folio only)",
        }

        # ---------- My SIP file: decide which date rule to use ----------
        date_rule = None   # ("point", idx) or ("range", start_idx, end_idx)
        if mode == "month":
            def usable(idx):
                return idx is not None and any(parse_any_date(cell(r, idx)) for r in data)

            i_point = _exact_header(hmap, SIP_POINT_DATE_HEADERS)
            if i_point is None:
                # exactly ONE other date column (not start/end) -> use it; otherwise don't guess
                others = [idx for key, idx in hmap.items()
                          if "DATE" in key and not any(w in key for w in
                             ("START", "END", "FROM", "TO", "REG", "CREAT", "CEASE", "CANCEL", "PRINT"))]
                if len(others) == 1:
                    i_point = others[0]
            i_start = _exact_header(hmap, SIP_START_HEADERS)
            i_end = _exact_header(hmap, SIP_END_HEADERS)

            if usable(i_point):
                date_rule = ("point", i_point)
                lookup.date_rule_text = (f"date in column '{header_row[i_point]}' falls in "
                                         f"{lookup.month_label}")
            elif usable(i_start):
                date_rule = ("range", i_start, i_end)
                end_txt = f" and '{header_row[i_end]}'" if i_end is not None else ""
                lookup.date_rule_text = (f"SIP running during {lookup.month_label} "
                                         f"(columns '{header_row[i_start]}'{end_txt})")
            else:
                lookup.date_rule_text = None
                lookup.all_headers = [h for h in header_row if h not in (None, "")]
                continue   # cannot tell the month -> fill nothing from this sheet

        # ---------- rows ----------
        for row in data:
            scheme, folio = cell(row, i_scheme), cell(row, i_folio)
            if not scheme or not folio:
                continue
            client = cell(row, i_client)
            status = cell(row, i_status)
            status_txt = str(status).strip().upper() if status not in (None, "") else "(BLANK)"
            lookup.status_counts[status_txt] = lookup.status_counts.get(status_txt, 0) + 1
            amount = to_number(cell(row, i_amount))

            if mode == "active":
                # XSIP: ACTIVE ONLY. Pause, CXL, blank, anything else -> ignored.
                if status_txt != "ACTIVE":
                    lookup.skipped_rows += 1
                    lookup.add_skipped(scheme, folio, amount, f"Status = {status_txt}", client)
                    continue
            else:
                # My SIP: current month only; stopped/cancelled SIPs never counted
                if i_status is not None and any(w in status_txt for w in SIP_STOPPED_STATUSES):
                    lookup.skipped_rows += 1
                    lookup.add_skipped(scheme, folio, amount, f"Status = {status_txt}", client)
                    continue
                if date_rule[0] == "point":
                    d = parse_any_date(cell(row, date_rule[1]))
                    in_month = d is not None and (d.year, d.month) == target
                else:
                    start = parse_any_date(cell(row, date_rule[1]))
                    end_raw = cell(row, date_rule[2]) if date_rule[2] is not None else None
                    end = parse_any_date(end_raw)
                    if start is None or (end_raw not in (None, "") and end is None):
                        in_month = False          # unreadable date -> not counted
                    else:
                        in_month = start <= month_last and (end is None or end >= month_first)
                if not in_month:
                    lookup.skipped_rows += 1
                    lookup.add_skipped(scheme, folio, amount,
                                       f"not in {lookup.month_label} (date outside current month)",
                                       client)
                    continue

            if amount is None or amount <= 0:
                lookup.skipped_rows += 1
                lookup.add_skipped(scheme, folio, amount, "no SIP amount in the file", client)
                continue

            lookup.add(scheme, folio, amount, client)


# ----------------------------------------------------------------------
# 3. BENCHMARK PDFs (Scheme benchmark / Broad benchmark)
# ----------------------------------------------------------------------
CLIENT_HEADER_RE = re.compile(r"([A-Z][A-Za-z .]+?)\s*\(PAN\s*:\s*([A-Z0-9]+)\)")


def extract_pdf_text(path):
    text = ""
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text += (page.extract_text() or "") + "\n"
    return text


def split_by_client(full_text):
    """dict: client_name (normalised) -> flattened text for that client.
    Keyed by NAME rather than PAN, because family members can share one PAN
    (e.g. minors under a guardian's PAN)."""
    matches = list(CLIENT_HEADER_RE.finditer(full_text))
    sections = {}
    for i, m in enumerate(matches):
        start = m.start()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(full_text)
        name = normalise_name(m.group(1))
        sections[name] = re.sub(r"\s+", " ", full_text[start:end])
    return sections


def find_xirr_diff(client_flat_text, scheme, folio):
    """Search for this exact scheme+folio inside the client's section and return XIRR - BM_XIRR."""
    if client_flat_text is None:
        return None
    scheme_pat = re.escape(scheme).replace(r"\ ", r"\s+")
    folio_pat = re.escape(str(folio))
    pattern = re.compile(
        scheme_pat + r".*?\[" + folio_pat + r"\]\s+((?:-?[\d,]+\s+){9}-?[\d.]+\s+-?[\d.]+)"
    )
    m = pattern.search(client_flat_text)
    if not m:
        return None
    nums = m.group(1).split()
    try:
        xirr, bm = float(nums[-2]), float(nums[-1])
    except ValueError:
        return None
    return round(xirr - bm, 2)


def load_benchmark_diffs(pdf_path):
    full_text = extract_pdf_text(pdf_path)
    return split_by_client(full_text)


# ----------------------------------------------------------------------
# 4. MASTER FILE  ->  parsed as an ordered list of row records
# ----------------------------------------------------------------------
def parse_master(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active

    # find header row (contains SCHEME and FOLIO NO)
    header_row_idx = None
    header = None
    for i, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if not row:
            continue
        norm = [normalise_header(v) if v is not None else "" for v in row]
        if "SCHEME" in norm and any(n.startswith("FOLIO") for n in norm):
            header_row_idx = i
            header = list(row)
            break
    if header_row_idx is None:
        raise ValueError("Could not find header row (SCHEME / FOLIO NO) in master file")

    # normalised header -> column position
    hmap = {}
    for i, name in enumerate(header):
        if name:
            hmap.setdefault(normalise_header(name), i)

    def cidx(*aliases):
        for a in aliases:
            key = normalise_header(a)
            if key in hmap:
                return hmap[key]
        return None

    i_scheme = cidx("SCHEME")
    i_folio = cidx("FOLIO NO", "FOLIO")
    i_curval = cidx("CURRENT VALUE")
    i_cagr = cidx("CAGR (%)", "CAGR")
    i_hold = cidx("HOLDING DAYS", "HOLDINGDAY")

    def val(row_vals, idx):
        if idx is None or idx >= len(row_vals):
            return None
        return row_vals[idx]

    records = []
    current_client_name = None
    pan_re = re.compile(r"\(([A-Z0-9]+)\)\s*$")
    name_re = re.compile(r"^(.*?)\s*\([A-Z0-9]+\)\s*$")

    for r in range(header_row_idx + 1, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c + 1).value for c in range(len(header))]
        scheme_cell = val(row_vals, i_scheme)
        folio_cell = val(row_vals, i_folio)

        if scheme_cell is None and all(v is None for v in row_vals):
            records.append({"type": "blank"})
            continue

        if folio_cell:  # a real scheme row
            records.append({
                "type": "scheme",
                "scheme": str(scheme_cell).strip(),
                "folio": str(folio_cell).strip(),
                "current_value": to_number(val(row_vals, i_curval)),
                "cagr": to_number(val(row_vals, i_cagr)),
                "holding_days": to_number(val(row_vals, i_hold)),
                "client_name": current_client_name,
            })
        elif scheme_cell and str(scheme_cell).strip().endswith("Total"):
            # UPDATE 1: carry Investwell's own totals across into the output.
            # Only CURRENT VALUE (C), CAGR (D) and HOLDING DAYS (I) are copied.
            # Our Rating / SIP / Scheme-Under / Broad-Nifty stay blank on totals.
            records.append({
                "type": "total",
                "label": str(scheme_cell).strip(),
                "current_value": to_number(val(row_vals, i_curval)),
                "cagr": to_number(val(row_vals, i_cagr)),
                "holding_days": to_number(val(row_vals, i_hold)),
                "client_name": current_client_name,
            })
        elif scheme_cell and pan_re.search(str(scheme_cell)):
            m = name_re.match(str(scheme_cell).strip())
            current_client_name = normalise_name(m.group(1)) if m else normalise_name(scheme_cell)
            records.append({"type": "client_header", "label": str(scheme_cell).strip()})
        elif scheme_cell:
            records.append({"type": "category_header", "label": str(scheme_cell).strip()})
        else:
            records.append({"type": "blank"})

    return records


# ----------------------------------------------------------------------
# 5. BUILD OUTPUT WORKBOOK
# ----------------------------------------------------------------------
def build_output(records, brokerage_rates, sip_totals, scheme_bm_sections, broad_bm_sections,
                 report_title, client_line, rm_line, printed_line, has_sip, out_path):

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Portfolio Report"

    bold = openpyxl.styles.Font(bold=True)
    header_fill = PatternFill("solid", fgColor="D9D9D9")
    total_fill = PatternFill("solid", fgColor="F2F2F2")

    ws.cell(row=1, column=1, value="NAYAN BALMUKUND SARAIYA").font = bold
    ws.cell(row=2, column=1, value=report_title).font = bold
    if not has_sip:
        c = ws.cell(row=2, column=2, value="NO SIP ")
        c.font = bold
        c.fill = PatternFill("solid", fgColor="FFFF00")
    ws.cell(row=3, column=1, value=client_line).font = bold
    ws.cell(row=3, column=3, value=rm_line)
    ws.cell(row=3, column=6, value=printed_line)

    headers = ["SCHEME", "FOLIO NO", "CURRENT VALUE", "CAGR (%)", "Our Rating",
               "SIP Amount", "Scheme/Under Value", "Broad/Nifty", "HOLDING DAYS"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = bold
        cell.fill = header_fill

    r = 5
    flags_report = []  # missing-data notes

    if has_sip and hasattr(sip_totals, "set_master_rows"):
        sip_totals.set_master_rows([(rec["scheme"], rec["folio"], rec.get("client_name"))
                                    for rec in records if rec["type"] == "scheme"])

    # Running SIP totals for the Total rows (worked out by the script, because the
    # Master file has no SIP column). None = nothing to add yet -> cell stays blank.
    sip_cat_sum = sip_client_sum = sip_grand_sum = None
    last_category = None
    # Client-wise NUMBER of SIPs (XSIP: Active SIPs; My SIP: this month's SIPs)
    sip_client_count = sip_grand_count = 0
    sip_count_summary = []          # (client total label, count, amount) for the printout
    if getattr(sip_totals, "mode", "active") == "month" and getattr(sip_totals, "month_label", None):
        count_word = f"SIPs in {sip_totals.month_label[:3]}: "
    elif getattr(sip_totals, "mode", "active") == "mixed":
        count_word = "SIPs: "                       # XSIP + My SIP files together
    else:
        count_word = "Active SIPs: "

    def _add(total, amount):
        return amount if total is None else total + amount

    scheme_rows, total_rows = [], []

    for rec in records:
        if rec["type"] == "blank":
            r += 1
            continue

        if rec["type"] in ("client_header", "category_header"):
            if rec["type"] == "category_header":
                last_category = rec["label"]
            else:
                sip_cat_sum = sip_client_sum = None      # new family member starts
                sip_client_count = 0
                last_category = None
            ws.cell(row=r, column=1, value=rec["label"]).font = bold
            r += 1
            continue

        if rec["type"] == "total":
            total_rows.append(r)
            # Label + the three values Investwell already calculated.
            ws.cell(row=r, column=COL["SCHEME"], value=rec["label"])
            if rec.get("current_value") is not None:
                ws.cell(row=r, column=COL["CURRENT_VALUE"], value=rec["current_value"])
            else:
                flags_report.append(
                    f"Total row '{rec['label']}': CURRENT VALUE blank (not present in master file)")
            if rec.get("cagr") is not None:
                ws.cell(row=r, column=COL["CAGR"], value=rec["cagr"])
            if rec.get("holding_days") is not None:
                ws.cell(row=r, column=COL["HOLDING_DAYS"], value=rec["holding_days"])
            # ---- SIP Amount total (sum of the SIP cells above it) ----
            label_norm = normalise_name(rec["label"])
            base = normalise_name(re.sub(r"\s*Total\s*$", "", rec["label"], flags=re.I))
            base = normalise_name(re.sub(r"\s*\([A-Z0-9]+\)\s*$", "", base))   # drop PAN if present
            if label_norm.startswith("GRAND TOTAL"):
                kind, sip_sum = "grand", sip_grand_sum
            elif rec.get("client_name") and base == rec["client_name"]:
                kind, sip_sum = "client", sip_client_sum
            elif last_category and base == normalise_name(last_category):
                kind, sip_sum = "category", sip_cat_sum
            else:
                kind, sip_sum = "unknown", None
            if has_sip:
                if kind == "unknown":
                    flags_report.append(
                        f"Total row '{rec['label']}': SIP total blank (could not tell if it is a "
                        f"category, client or grand total)")
                elif sip_sum is not None and sip_sum > 0:
                    ws.cell(row=r, column=COL["SIP"], value=sip_sum)
                # number of SIPs, shown in column B of client total and Grand Total rows
                n = {"client": sip_client_count, "grand": sip_grand_count}.get(kind, 0)
                if kind in ("client", "grand") and n > 0:
                    ws.cell(row=r, column=COL["FOLIO"], value=f"{count_word}{n}")
                if kind in ("client", "grand"):
                    sip_count_summary.append((rec["label"], n, sip_sum))
            if kind == "category":
                sip_cat_sum = None
            elif kind == "client":
                sip_cat_sum = sip_client_sum = None
                sip_client_count = 0
            # Our Rating / Scheme-Under / Broad-Nifty deliberately left blank on totals.

            for c in range(1, 10):
                ws.cell(row=r, column=c).font = bold
                ws.cell(row=r, column=c).fill = total_fill
            r += 1
            continue

        if rec["type"] == "scheme":
            scheme_rows.append(r)
            scheme, folio = rec["scheme"], rec["folio"]
            ws.cell(row=r, column=COL["SCHEME"], value=scheme)
            ws.cell(row=r, column=COL["FOLIO"], value=folio)
            ws.cell(row=r, column=COL["CURRENT_VALUE"], value=rec["current_value"])
            ws.cell(row=r, column=COL["CAGR"], value=rec["cagr"])
            ws.cell(row=r, column=COL["HOLDING_DAYS"], value=rec["holding_days"])

            # ---- Our Rating (scheme name only - Income Report has no folio) ----
            rate = brokerage_rates.get(normalise_scheme(scheme))
            if rate is not None:
                ws.cell(row=r, column=COL["OUR_RATING"], value=rate)
            else:
                flags_report.append(f"Our Rating blank: no brokerage match for '{scheme}' / {folio}")

            # ---- SIP Amount (Scheme + Folio, ACTIVE only) ----
            if has_sip:
                amt = sip_totals.get(scheme, folio, rec.get("client_name"))
                if amt is not None and amt > 0:
                    ws.cell(row=r, column=COL["SIP"], value=amt)
                    sip_cat_sum = _add(sip_cat_sum, amt)
                    sip_client_sum = _add(sip_client_sum, amt)
                    sip_grand_sum = _add(sip_grand_sum, amt)
                    sip_client_count += sip_totals.last_count
                    sip_grand_count += sip_totals.last_count
                else:
                    why = sip_totals.explain_blank(scheme, folio, rec.get("client_name"))
                    if why:
                        flags_report.append(
                            f"SIP Amount blank: '{scheme}' / {folio} IS in the SIP file but "
                            f"not counted ({why})")
                    else:
                        flags_report.append(
                            f"SIP Amount blank: no SIP in the file for '{scheme}' / {folio}")
            # if this client has no SIP file at all, leave blank (column still present)

            # ---- Scheme/Under Value ----
            client_name = rec["client_name"]
            sc_section = scheme_bm_sections.get(client_name)
            diff = find_xirr_diff(sc_section, scheme, folio)
            if diff is not None:
                ws.cell(row=r, column=COL["SCHEME_UNDER"], value=diff)
            else:
                flags_report.append(f"Scheme/Under Value blank: no XIRR match for '{scheme}' / {folio}")

            # ---- Broad/Nifty ----
            br_section = broad_bm_sections.get(client_name)
            diff2 = find_xirr_diff(br_section, scheme, folio)
            if diff2 is not None:
                ws.cell(row=r, column=COL["BROAD_NIFTY"], value=diff2)
            else:
                flags_report.append(f"Broad/Nifty blank: no XIRR match for '{scheme}' / {folio}")

            r += 1
            continue

    # ---- number formats ----
    for row in ws.iter_rows(min_row=5, max_row=max(r - 1, 5)):
        if len(row) < 9:
            continue
        row[2].number_format = "#,##0.00"  # CURRENT VALUE
        row[3].number_format = "0.00"      # CAGR
        row[4].number_format = "0.00"      # Our Rating
        row[5].number_format = "#,##0"     # SIP Amount
        row[6].number_format = "0.00"      # Scheme/Under
        row[7].number_format = "0.00"      # Broad/Nifty
        row[8].number_format = "#,##0"     # HOLDING DAYS

    apply_colour_rules(ws, r - 1)
    set_column_widths(ws)
    style_main_table(ws, 4, scheme_rows + total_rows)

    # ---- the 4 filtered lists below the Grand Total ----
    global LAST_LISTS_SUMMARY
    last_main = max(scheme_rows + total_rows) if (scheme_rows or total_rows) else 4
    LAST_LISTS_SUMMARY = add_four_lists(ws, 4, scheme_rows, last_main)

    wb.save(out_path)
    if has_sip:
        sip_totals.client_summary = sip_count_summary
        sip_totals.count_word = count_word
    return flags_report


# ----------------------------------------------------------------------
# 6. COLOURING (scheme rows only - never totals/headers)
# ----------------------------------------------------------------------
def is_scheme_data_row(ws, row_idx):
    scheme = ws.cell(row=row_idx, column=1).value
    folio = ws.cell(row=row_idx, column=2).value
    if not scheme or not folio:
        return False
    if str(scheme).strip().endswith("Total"):
        return False
    return True


def apply_colour_rules(ws, last_row):
    for r in range(5, last_row + 1):
        if not is_scheme_data_row(ws, r):
            continue  # totals, subtotals, category/client headers, grand total: never coloured

        rating = ws.cell(row=r, column=COL["OUR_RATING"]).value
        cagr = ws.cell(row=r, column=COL["CAGR"]).value
        sip = ws.cell(row=r, column=COL["SIP"]).value
        under = ws.cell(row=r, column=COL["SCHEME_UNDER"]).value
        broad = ws.cell(row=r, column=COL["BROAD_NIFTY"]).value

        rating_trig = brown_trig = under_neg = broad_neg = False

        rating_fill = None
        if is_number(rating):
            if 0.75 <= rating < 1.00:
                rating_fill = FILL_RATING_DEEP_PINK
                rating_trig = True
            elif rating < 0.75:
                rating_fill = FILL_RATING_PALE_PINK
        if rating_fill:
            ws.cell(row=r, column=COL["OUR_RATING"]).fill = rating_fill

        cagr_fill = None
        if is_number(cagr):
            if 6.00 <= cagr < 10.00:
                cagr_fill = FILL_CAGR_BROWN
                brown_trig = True
            elif cagr < 6.00:
                cagr_fill = FILL_CAGR_LIGHT_BROWN
        if cagr_fill:
            ws.cell(row=r, column=COL["CAGR"]).fill = cagr_fill

        if is_number(sip) and sip > 0:
            ws.cell(row=r, column=COL["SIP"]).fill = FILL_SIP_GREEN

        if is_number(under) and under < 0:
            ws.cell(row=r, column=COL["SCHEME_UNDER"]).fill = FILL_NEGATIVE_RED
            under_neg = True

        if is_number(broad) and broad < 0:
            ws.cell(row=r, column=COL["BROAD_NIFTY"]).fill = FILL_NEGATIVE_RED
            broad_neg = True

        # Whole-row violet override: A-H only, never column I
        if rating_trig and brown_trig and under_neg and broad_neg:
            for c in range(1, 9):
                ws.cell(row=r, column=c).fill = FILL_VIOLET


def set_column_widths(ws):
    widths = {"A": 58, "B": 16, "C": 14, "D": 10, "E": 12, "F": 12, "G": 14, "H": 12, "I": 13}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


# ======================================================================
# 7. THE FOUR FILTERED LISTS (added below the Grand Total)
# ======================================================================
# ---- SETTINGS: edit these lines to change the lists ------------------
LIST_RATING_LIMIT = 0.80        # List 1: Our Rating below this is "low"
LIST_HIGH_SIP_LIMIT = 15000     # List 2: SIP strictly above this is "high"

STOP_SIP_SCHEMES = [            # List 3: named schemes to stop SIP
    "Parag Parikh Flexi Cap Fund Reg (G)",
    "DSP Multi Asset Allocation Fund Reg (G)",
    "DSP Multi Asset Omni FoF (G)",
    "HDFC Large Cap Fund (G)",
    "Templeton India Value Fund (G)",
    "Axis Focused Fund Reg (G)",
]
STOP_SIP_NAME_CONTAINS = ["ICICI Pru"]   # List 3: every scheme whose name contains this

NEW_START_SIP_SCHEMES = [       # List 4: named schemes to start a new SIP
    "Mahindra Manulife Multi Cap Fund Reg (G)",
    "Bank of India Multi Cap Fund Reg (G)",
    "Groww Multi Cap Fund Reg (G)",
    "Bandhan Innovation Fund Reg (G)",
    "Kotak Pioneer Fund (G)",
    "Mirae Asset Gold ETF FoF Reg (G)",
    "TrustMF Multi Cap Fund Reg (G)",
]

# "E T F" (how Investwell writes it) = "ETF" (how the list writes it)?
#   True  -> "Mirae Asset Gold E T F FoF Reg (G)" counts as "Mirae Asset Gold ETF FoF Reg (G)"
#   False -> letters written with spaces must be written the same way to match
LIST_JOIN_SPACED_LETTERS = True
# ----------------------------------------------------------------------

from copy import copy as _copy
from difflib import SequenceMatcher
from openpyxl.styles import Border, Side, Font, Alignment

THIN = Side(style="thin")
BORDER_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
LIST_FONT = Font(name="Arial", size=10)
LIST_FONT_BOLD = Font(name="Arial", size=10, bold=True)
LIST_HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")

# words that don't decide WHICH scheme it is
_LIST_FILLER = {"fund", "reg", "regular", "plan", "option", "opt", "growth", "the"}


def list_name_tokens(name):
    """'Mirae Asset Gold E T F FoF Reg (G)' -> ['mirae','asset','gold','etf','fof']
    'Bank of India Multi Cap Fund Reg (G)' -> ['bank','of','india','multicap']"""
    text = str(name).lower().replace("&", " and ")
    text = re.sub(r"\(\s*(g|gr|growth)\s*\)", " growth ", text)
    raw = re.findall(r"[a-z]+|[0-9]+", text)
    merged, letters = [], ""
    for w in raw:                                   # "e t f" -> "etf"
        if len(w) == 1 and w.isalpha() and LIST_JOIN_SPACED_LETTERS:
            letters += w
            continue
        if letters:
            merged.append(letters)
            letters = ""
        merged.append(w)
    if letters:
        merged.append(letters)
    toks = []
    for w in merged:
        if w == "cap" and toks:                      # "multi cap" -> "multicap"
            toks[-1] += "cap"
        elif w not in _LIST_FILLER:
            toks.append(w)
    return toks


def _word_same(a, b):
    if a == b:
        return True
    if any(ch.isdigit() for ch in a + b):            # numbers must match exactly
        return False
    if len(a) < 5 or len(b) < 5 or a[0] != b[0]:
        return False
    return SequenceMatcher(None, a, b).ratio() >= 0.85   # tiny spelling slip only


def list_names_match(wanted, actual):
    """Fund house (first word) must be identical and EVERY other word must match,
    in the same order. 'Multi Cap' is never 'Flexi Cap'."""
    w, a = list_name_tokens(wanted), list_name_tokens(actual)
    if not w or not a or w[0] != a[0] or len(w) != len(a):
        return False
    return all(_word_same(x, y) for x, y in zip(w[1:], a[1:]))


def _find_col(ws, header_row, *names):
    wanted = [normalise_header(n) for n in names]
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        if v is not None and normalise_header(v) in wanted:
            return c
    return None


def add_four_lists(ws, header_row, scheme_rows, last_main_row):
    """Writes Lists 1-4 below the main table. Returns a summary dict for the printout."""
    ncols = 9
    for c in range(1, ws.max_column + 1):
        if ws.cell(row=header_row, column=c).value not in (None, ""):
            ncols = max(ncols, c)
    headers = [ws.cell(row=header_row, column=c).value for c in range(1, ncols + 1)]
    flag_col = ncols + 1

    c_scheme = _find_col(ws, header_row, "SCHEME")
    c_cv = _find_col(ws, header_row, "Current Value")
    c_rating = _find_col(ws, header_row, "Our Rating")
    c_sip = _find_col(ws, header_row, "SIP Amount")
    c_under = _find_col(ws, header_row, "Scheme/Under Value", "scheme(XIRR-BM XIRR)")
    c_broad = _find_col(ws, header_row, "Broad/Nifty", "nifty(XIRR-BM XIRR)")

    def v(r, c):
        val = ws.cell(row=r, column=c).value if c else None
        return val if is_number(val) else None

    rows = [{"r": r, "name": str(ws.cell(row=r, column=c_scheme).value or "").strip(),
             "cv": v(r, c_cv), "rating": v(r, c_rating), "sip": v(r, c_sip),
             "under": v(r, c_under), "broad": v(r, c_broad)} for r in scheme_rows]

    at_threshold = []

    # ---------- List 1 ----------
    l1 = []
    for x in rows:
        sn = x["under"] is not None and x["under"] < 0
        bn = x["broad"] is not None and x["broad"] < 0
        lr = x["rating"] is not None and x["rating"] < LIST_RATING_LIMIT
        parts = []
        if sn and bn:
            parts.append("Both negative")
        elif sn:
            parts.append("Scheme negative")
        elif bn:
            parts.append("Broad negative")
        if lr:
            parts.append(f"Rating <{LIST_RATING_LIMIT:.2f}")
        if parts:
            l1.append((x, " | ".join(parts)))
        else:
            if x["under"] == 0 or x["broad"] == 0:
                at_threshold.append(f"List 1: '{x['name']}' alpha exactly 0 (not negative)")
            if x["rating"] is not None and abs(x["rating"] - LIST_RATING_LIMIT) < 1e-9:
                at_threshold.append(f"List 1: '{x['name']}' rating exactly {LIST_RATING_LIMIT:.2f}")

    # ---------- List 2 ----------
    l2 = []
    for x in rows:
        if x["sip"] is not None and x["sip"] > LIST_HIGH_SIP_LIMIT:
            l2.append((x, f"SIP >{LIST_HIGH_SIP_LIMIT:,.0f}"))
        elif x["sip"] is not None and abs(x["sip"] - LIST_HIGH_SIP_LIMIT) < 1e-9:
            at_threshold.append(f"List 2: '{x['name']}' SIP exactly {LIST_HIGH_SIP_LIMIT:,.0f}")

    # ---------- named-scheme matching (Lists 3 & 4) ----------
    written_diff = []
    def match_named(named_list):
        found, not_found, near = {}, [], []
        for wanted in named_list:
            hits = [x for x in rows if list_names_match(wanted, x["name"])]
            if hits:
                for x in hits:
                    found[x["r"]] = wanted
                    if normalise_scheme(wanted) != normalise_scheme(x["name"]):
                        written_diff.append((wanted, x["name"]))
            else:
                not_found.append(wanted)
                house = list_name_tokens(wanted)[:1]
                same_house = sorted({x["name"] for x in rows
                                     if list_name_tokens(x["name"])[:1] == house})
                for s in same_house:
                    near.append((wanted, s))
        return found, not_found, near

    stop_found, stop_missing, stop_near = match_named(STOP_SIP_SCHEMES)
    l3 = []
    contains = [re.sub(r"\s+", " ", c).lower() for c in STOP_SIP_NAME_CONTAINS]
    for x in rows:
        nm = re.sub(r"\s+", " ", x["name"]).lower()
        if x["r"] in stop_found:
            l3.append((x, "Named — stop SIP"))
        elif any(c in nm for c in contains):
            l3.append((x, f"{STOP_SIP_NAME_CONTAINS[0]} — stop SIP" if len(contains) == 1
                       else "Name contains — stop SIP"))

    new_found, new_missing, new_near = match_named(NEW_START_SIP_SCHEMES)
    l4 = [(x, "New start SIP") for x in rows if x["r"] in new_found]

    # ---------- titles ----------
    def n_s(n):
        return f"{n} scheme" if n == 1 else f"{n} schemes"

    def sip_split(lst):
        has = sum(1 for x, _ in lst if x["sip"] is not None and x["sip"] > 0)
        return has, len(lst) - has

    def has_txt(n):
        return f"{n} has a SIP" if n == 1 else f"{n} have a SIP"

    def not_txt(n):
        return f"{n} does not" if n == 1 else f"{n} do not"

    l3_has, l3_no = sip_split(l3)
    l4_has, l4_no = sip_split(l4)
    n_named_found_4 = len(NEW_START_SIP_SCHEMES) - len(new_missing)
    lists = [
        (f"List 1: Negative Alpha or Low Rating (Scheme alpha < 0, or Broad alpha < 0, "
         f"or Our Rating < {LIST_RATING_LIMIT:.2f}) — {n_s(len(l1))}", l1, False),
        (f"List 2: High SIP (SIP Amount > {LIST_HIGH_SIP_LIMIT:,.0f}) — {n_s(len(l2))}", l2, True),
        (f"List 3: Stop SIP — {n_s(len(l3))} ({has_txt(l3_has)}, {not_txt(l3_no)}; "
         f"{len(stop_missing)} of the {len(STOP_SIP_SCHEMES)} named schemes were not found)", l3, True),
        (f"List 4: New Start SIP — {n_named_found_4} of the {len(NEW_START_SIP_SCHEMES)} named "
         f"schemes found, {n_s(len(l4))} included ({has_txt(l4_has)}, {not_txt(l4_no)})", l4, True),
    ]

    # ---------- write ----------
    r = last_main_row + 2
    results = []
    for title, items, sip_total in lists:
        ws.cell(row=r, column=1, value=title).font = LIST_FONT_BOLD
        r += 1
        for c in range(1, flag_col + 1):
            cell = ws.cell(row=r, column=c, value=headers[c - 1] if c <= ncols else "FLAG")
            cell.font = LIST_FONT_BOLD
            cell.fill = LIST_HEADER_FILL
            cell.border = BORDER_ALL
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        r += 1
        first = r
        if items:
            for x, flag in items:
                for c in range(1, ncols + 1):
                    src = ws.cell(row=x["r"], column=c)
                    dst = ws.cell(row=r, column=c, value=src.value)
                    dst.number_format = src.number_format
                    if src.fill is not None and src.fill.fill_type:
                        dst.fill = _copy(src.fill)
                    dst.alignment = _copy(src.alignment)
                    dst.font = LIST_FONT
                    dst.border = BORDER_ALL
                fc = ws.cell(row=r, column=flag_col, value=flag)
                fc.font = LIST_FONT
                fc.border = BORDER_ALL
                fc.alignment = Alignment(horizontal="left")
                r += 1
        else:
            ws.cell(row=r, column=1, value="No schemes matched")
            for c in range(1, flag_col + 1):
                ws.cell(row=r, column=c).font = LIST_FONT
                ws.cell(row=r, column=c).border = BORDER_ALL
            r += 1
        last = r - 1
        cvL = openpyxl.utils.get_column_letter(c_cv)
        sipL = openpyxl.utils.get_column_letter(c_sip) if c_sip else None
        tot_row = r
        ws.cell(row=r, column=1, value="Total Current Value")
        ws.cell(row=r, column=c_cv, value=f"=SUM({cvL}{first}:{cvL}{last})").number_format = "0.00"
        if sip_total and sipL:
            ws.cell(row=r, column=c_sip, value=f"=SUM({sipL}{first}:{sipL}{last})").number_format = "#,##0"
        r += 1
        ws.cell(row=r, column=1, value="Total Current Value ÷ 36")
        ws.cell(row=r, column=c_cv, value=f"={cvL}{tot_row}/36").number_format = "0.00"
        for rr in (tot_row, r):
            for c in range(1, flag_col + 1):
                cell = ws.cell(row=rr, column=c)
                cell.border = BORDER_ALL
                cell.font = LIST_FONT_BOLD if cell.value not in (None, "") else LIST_FONT
        r += 2          # one blank row before the next list

        cv_total = sum(x["cv"] or 0 for x, _ in items)
        sip_sum = sum(x["sip"] or 0 for x, _ in items)
        results.append({"title": title, "count": len(items), "cv": cv_total,
                        "cv36": cv_total / 36, "sip": sip_sum if sip_total else None,
                        "no_sip": sum(1 for x, _ in items if not (x["sip"] and x["sip"] > 0))})

    # widths: A at least 45 (or the longest name), J (FLAG) 28
    col_a = openpyxl.utils.get_column_letter(1)
    longest = max([len(x["name"]) for x in rows] + [0])
    ws.column_dimensions[col_a].width = max(ws.column_dimensions[col_a].width or 0, 45, longest + 2)
    ws.column_dimensions[openpyxl.utils.get_column_letter(flag_col)].width = 28

    return {"lists": results, "at_threshold": at_threshold, "written_diff": written_diff,
            "stop_missing": stop_missing, "stop_near": stop_near,
            "new_missing": new_missing, "new_near": new_near}


def style_main_table(ws, header_row, bordered_rows, ncols=9):
    """Match the house style of the final report: thin borders on the header, scheme
    and total rows (A-I), centred wrapped header, frozen header."""
    for c in range(1, ncols + 1):
        cell = ws.cell(row=header_row, column=c)
        cell.border = BORDER_ALL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[header_row].height = 31.5
    for r in bordered_rows:
        for c in range(1, ncols + 1):
            ws.cell(row=r, column=c).border = BORDER_ALL
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)


def print_lists_summary(s):
    if not s:
        return
    print("\n================ 4 LISTS ================")
    for i, L in enumerate(s["lists"], start=1):
        line = (f"List {i}: {L['count']} scheme(s) | Total Current Value {L['cv']:,.2f} "
                f"| ÷36 = {L['cv36']:,.2f}")
        if L["sip"] is not None:
            line += f" | Total SIP {L['sip']:,.0f}"
        if i in (3, 4):
            line += f" | without SIP: {L['no_sip']}"
        print(line)
    if s.get("written_diff"):
        print("\nNamed schemes matched although written slightly differently - please glance:")
        for wanted, actual in s["written_diff"]:
            print(f" - list: '{wanted}'  =  portfolio: '{actual}'")
    if s["at_threshold"]:
        print("\nLeft out because the value sits EXACTLY at the limit:")
        for t in s["at_threshold"]:
            print(" -", t)
    for label, missing, near in (("List 3 (Stop SIP)", s["stop_missing"], s["stop_near"]),
                                 ("List 4 (New Start SIP)", s["new_missing"], s["new_near"])):
        if missing:
            print(f"\n{label} - named schemes NOT found in this portfolio:")
            for m in missing:
                print(f" - {m}")
                for w, n in near:
                    if w == m:
                        print(f"     (same fund house in portfolio, NOT matched: {n})")


# ----------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--master", required=True)
    ap.add_argument("--income", required=True)
    ap.add_argument("--scheme-bm", required=True)
    ap.add_argument("--broad-bm", required=True)
    ap.add_argument("--sip", required=False, default=None, nargs="+",
                    help="One or more SIP files (e.g. one XSIP file per family member)")
    ap.add_argument("--no-sip", action="store_true",
                    help="Force a NO SIP report even if a SIP file is present in the folder")
    ap.add_argument("--out", required=True)
    ap.add_argument("--report-title", default="Portfolio Valuation Report")
    ap.add_argument("--client-line", default="")
    ap.add_argument("--rm-line", default="")
    ap.add_argument("--printed-line", default="")
    args = ap.parse_args()
    print(f"build_report.py {SCRIPT_VERSION}")

    # ---- work out which SIP file (if any) to use ----
    sip_path = []                       # list of SIP files (one per family member is fine)
    folder = os.path.dirname(os.path.abspath(args.master)) or "."
    if args.no_sip:
        print("--no-sip given: SIP column will be left blank and tagged NO SIP.")
    elif args.sip:
        sip_path = [p for p in args.sip if os.path.exists(p)]
        for p in args.sip:
            if not os.path.exists(p):
                print(f"WARNING: SIP file '{p}' not found - skipped.")
        if not sip_path:
            print("Looking for SIP files in this folder instead...")
            sip_path = find_all_sip_files(folder)
    else:
        sip_path = find_all_sip_files(folder)
        if not sip_path:
            print("No SIP file found in this folder - report will be tagged NO SIP.")
    if sip_path:
        print(f"SIP file(s) used ({len(sip_path)}):")
        for p in sip_path:
            print(f"  - {os.path.basename(p)}")

    print("Reading brokerage rates...")
    brokerage_rates = load_brokerage_rates(args.income)

    if sip_path:
        print("Reading SIP report(s) ...")
        try:
            sip_totals = load_sip_totals(sip_path)
        except Exception as e:
            print("\n  ERROR: could not read the SIP file.")
            print("  " + str(e).replace("\n", "\n  "))
            print("\n  The report was NOT created. Fix the SIP file and run again.")
            sys.exit(1)
        for fname, info in sip_totals.file_info.items():
            names = ", ".join(sorted(sip_totals.file_clients.get(fname, ()))) or "(no member name)"
            extra = f", {info['replaced']} already in XSIP / a newer file" if info["replaced"] else ""
            print(f"  {info['kind']:6} {fname}")
            print(f"         rule: {info['rule']} | member(s): {names}")
            print(f"         counted: {info['counted']} | not counted: {info['skipped']}{extra}")
        for fname, same in sip_totals.duplicate_files:
            print(f"  DUPLICATE {fname}: same content as {same} - read once only")
        for err in sip_totals.errors:
            print(f"  ERROR  {err}")
        if len(sip_totals) == 0:
            print("  NOTE: no usable SIPs found in the SIP file(s) - every SIP cell will stay blank.")
            if sip_totals.errors:
                print("\n  The report was NOT created. Fix the SIP file(s) and run again.")
                sys.exit(1)
    else:
        sip_totals = SipLookup()

    has_sip = bool(sip_path) and len(sip_totals) > 0

    print("Reading scheme benchmark PDF...")
    scheme_bm_sections = load_benchmark_diffs(args.scheme_bm)

    print("Reading broad benchmark PDF...")
    broad_bm_sections = load_benchmark_diffs(args.broad_bm)

    print("Reading master file...")
    records = parse_master(args.master)

    print("Building output...")
    flags = build_output(
        records, brokerage_rates, sip_totals, scheme_bm_sections, broad_bm_sections,
        report_title=args.report_title, client_line=args.client_line,
        rm_line=args.rm_line, printed_line=args.printed_line,
        has_sip=has_sip, out_path=args.out,
    )

    print(f"\nDone. Saved to {args.out}")
    print_lists_summary(LAST_LISTS_SUMMARY)

    summary = getattr(sip_totals, "client_summary", None)
    if summary:
        word = sip_totals.count_word.replace(": ", "").strip()
        print(f"\nClient-wise SIP summary ({word}):")
        for label, n, amt in summary:
            amt_txt = f"Rs {amt:,.0f}" if amt else "-"
            print(f"   {label:40} {n:>3} SIP(s)   {amt_txt}")
    if sip_totals and sip_totals.name_fixed:
        print(f"\n{len(sip_totals.name_fixed)} SIP(s) matched after cleaning up the scheme name "
              f"(same folio, same starting name, shortcuts at the end ignored) - "
              f"please glance at these:")
        for sip_name, master_name, fol, amt, how, sip_fol in sip_totals.name_fixed:
            print(f" - SIP file: '{sip_name}' / folio {sip_fol}")
            print(f"   Master  : '{master_name}' / folio {fol}  (Rs {amt:,.0f}, matched by: {how})")
    if sip_totals and sip_totals.name_mismatch:
        print(f"\n{len(sip_totals.name_mismatch)} SIP(s) have the same scheme + folio as a Master row, "
              f"but are registered under a DIFFERENT member - left blank, not guessed:")
        for sip_client, member, sch, fol, amt in sip_totals.name_mismatch:
            print(f" - '{sch}' / folio {fol}: SIP file says '{sip_client}', Master row is under "
                  f"'{member}' (Rs {amt:,.0f})")
    if sip_totals:
        leftover = sip_totals.unused()
        if leftover:
            print(f"\n{len(leftover)} counted SIP(s) in the SIP file did NOT match any Scheme+Folio "
                  f"row in the Master file (left out, not guessed):")
            for sch, fol, amt in leftover:
                print(f" - '{sch}' / folio {fol}  (Rs {amt:,.0f})")
                same = sip_totals.same_folio_master_rows(fol)
                if same:
                    print(f"     Master rows with this folio: " + " | ".join(same))
                    print(f"     -> name differs too much to be sure. If it IS one of these (e.g. the")
                    print(f"        scheme was renamed), add it to SCHEME_NAME_ALIASES in build_report.py")
                else:
                    print(f"     -> no row in the Master file has this folio (another family member,")
                    print(f"        or a folio not held in this portfolio)")
        else:
            print("\nAll counted SIPs in the SIP file were placed in the report.")
    if sip_path:
        s = sip_totals.summary()
        print(f"\nSIP UPLOAD SUMMARY: {s['updated']} updated (Rs {s['updated_amount']:,.0f}) | "
              f"{s['unmatched']} unmatched | {s['errors']} error(s) | "
              f"{s['not_counted']} not counted | {s['duplicates']} duplicate(s) not added twice")
        fallbacks = [p for p in sip_totals.placed if p[4]]
        if fallbacks:
            print("  Last month's amount used (current month was 0):")
            for member, sch, fol, amt, note in fallbacks:
                print(f"   - {member}: '{sch}' / {fol}  Rs {amt:,.0f}  ({note})")
    if flags:
        print(f"\n{len(flags)} cell(s) left BLANK (no confident match found):")
        for f in flags:
            print(" -", f)


if __name__ == "__main__":
    main()
