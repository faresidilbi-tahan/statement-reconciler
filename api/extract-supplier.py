"""
POST /api/extract-supplier
Body: { "file_base64": "<base64>", "filename": "statement.pdf" }
(also accepts the older key "pdf_base64" for backward compatibility)

Deterministic, format-agnostic supplier-statement parser - no LLM, no
external calls. Every supplier sends statements in their own layout and
their own file format, so this endpoint:
  1. sniffs the actual file type from its bytes (PDF / XLSX / CSV /
     plain text) rather than trusting the filename or content-type,
  2. routes to a format-specific reader that turns the file into plain
     rows of cells,
  3. runs the SAME column-keyword header detection and row-building
     logic against those cells regardless of source format, so adding
     a synonym to COLUMN_KEYWORDS improves every format at once.

Row id priority: an INV/C-N number inside the description (this is the
number that also appears on our statement as "Inv#"), else whichever
id-like column the sheet provides (invoice number > check/voucher number
> transaction id). Dates in "Jan 12, 2026", "12/01/2026", or native Excel
date form are all normalized to YYYY-MM-DD.
"""

from http.server import BaseHTTPRequestHandler
import base64
import csv
import io
import json
import re
import datetime as dt

import pdfplumber
import openpyxl

BUILD_TAG = "2026-08-12-multiformat"

# ------------------------------------------------------------ shared vocab

COLUMN_KEYWORDS = {
    "date": ["date", "invc", "تاريخ", "التاريخ"],
    "id": ["ref", "reference", "invoice", "inv", "voucher", "vch", "doc",
           "document", "no", "number", "num", "trx", "nb", "id", "check",
           "chk", "cheque", "رقم", "سند", "مستند", "فاتورة", "المرجع", "مرجع"],
    "description": ["description", "details", "narration", "particulars",
                    "memo", "remarks", "بيان", "البيان", "التفاصيل",
                    "الوصف", "ملاحظات"],
    "debit": ["debit", "dr", "مدين"],
    "credit": ["credit", "cr", "دائن"],
    "balance": ["balance", "رصيد", "الرصيد"],
}

# id-like columns ranked by how reliably they match OUR invoice numbers;
# used only for spreadsheet parsing, where several id columns can coexist
# on one row (e.g. "Trans Id", "Invoice number", "Check Num" all at once)
ID_COLUMN_PRIORITY = ["invoice", "voucher", "vch", "check", "chk", "cheque",
                       "ref", "reference", "doc", "document", "trx", "id",
                       "no", "number", "num", "nb"]

OPENING_WORDS = ("opening", "b/f", "b.f", "brought", "balance until",
                 "افتتاحي", "سابق", "مدور", "رصيد اول")
CLOSING_WORDS = ("closing", "c/f", "c.f", "carried", "total", "grand",
                 "ending balance", "end date", "balance as at",
                 "اجمالي", "إجمالي", "المجموع", "ختامي", "نهائي")
SKIP_WORDS = ("statement", "page ", "page:", "printed", "tel:", "fax:",
              "p.o.box", "www.", "@", "pdc")

AMOUNT_RE = re.compile(r"^\(?-?(?:[\d,]+(?:\.\d+)?|\.\d+)\)?(CR|DR)?$", re.IGNORECASE)
NUM_DATE_RE = re.compile(r"(\d{1,4})[/\-.](\d{1,2})[/\-.](\d{1,4})")
MONTH_DATE_RE = re.compile(
    r"(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(\d{1,2})\s*,?\s+(\d{4})",
    re.IGNORECASE)
MONTHS = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
          "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12}
REF_RE = re.compile(r"(?:INV|C/N)\s*[:#]?\s*(\d{4,})", re.IGNORECASE)


def match_column(text):
    t = str(text or "").strip().lower().strip(":.")
    for col, words in COLUMN_KEYWORDS.items():
        if t in words:
            return col
    parts = [p.strip(":.") for p in t.split()]
    for col, words in COLUMN_KEYWORDS.items():
        if any(p in words for p in parts):
            return col
    return None


def id_column_rank(header_text):
    """How specific an id-like header is, for picking the best of several
    id columns on one spreadsheet row. Lower is more trustworthy."""
    t = str(header_text or "").strip().lower()
    for i, kw in enumerate(ID_COLUMN_PRIORITY):
        if kw in t:
            return i
    return len(ID_COLUMN_PRIORITY)


def parse_amount(token):
    token = str(token).strip()
    m = AMOUNT_RE.match(token)
    if not m:
        return None
    s = token[: len(token) - 2] if m.group(1) else token
    s = s.replace(",", "").replace("(", "-").replace(")", "")
    try:
        return float(s)
    except ValueError:
        return None


def find_date(text):
    """Search a cell for a date in 'Jan 12, 2026' or numeric form -> ISO."""
    text = str(text or "")
    m = MONTH_DATE_RE.search(text)
    if m:
        mo = MONTHS[m.group(1).lower()[:3]]
        d, y = int(m.group(2)), int(m.group(3))
        if 1 <= d <= 31:
            return f"{y:04d}-{mo:02d}-{d:02d}"
    m = NUM_DATE_RE.search(text)
    if m:
        a, b, c = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if a > 31:
            y, mo, d = a, b, c
        else:
            d, mo, y = a, b, c
            if y < 100:
                y += 2000
        if 1 <= d <= 31 and 1 <= mo <= 12 and 1900 < y < 2100:
            return f"{y:04d}-{mo:02d}-{d:02d}"
    return None


def extract_row_id(id_cell, description):
    m = REF_RE.search(str(description or ""))
    if m:
        return m.group(1)
    for tok in str(id_cell or "").split():
        if re.search(r"\d", tok) and not find_date(tok):
            return tok
    return ""


def build_row(cells, raw_low):
    date = find_date(cells.get("date", "")) or ""
    debit = parse_amount(cells.get("debit", "")) if cells.get("debit", "") != "" else None
    credit = parse_amount(cells.get("credit", "")) if cells.get("credit", "") != "" else None
    balance = parse_amount(cells.get("balance", "")) if cells.get("balance", "") != "" else None

    if any(k in raw_low for k in OPENING_WORDS):
        row_type = "opening_balance"
    elif any(k in raw_low for k in CLOSING_WORDS):
        row_type = "closing_balance"
    else:
        row_type = "transaction"

    if not (date or any(v is not None for v in (debit, credit, balance))):
        return None
    desc = cells.get("description", "")
    return {
        "date": date,
        "id": extract_row_id(cells.get("id", ""), desc) if row_type == "transaction" else "",
        "description": desc,
        # sign is presentation-only (some suppliers print credits negative)
        "debit": abs(debit) if debit is not None else 0.0,
        "credit": abs(credit) if credit is not None else 0.0,
        "balance": balance if balance is not None else 0.0,
        "row_type": row_type,
    }


def is_continuation(cells):
    if find_date(cells.get("date", "")):
        return False
    if any(parse_amount(cells.get(c, "")) is not None for c in ("debit", "credit", "balance") if cells.get(c, "") != ""):
        return False
    return bool(cells.get("description") or cells.get("id"))


# ================================================================== PDF

def group_lines(page):
    words = page.extract_words(x_tolerance=1.5, y_tolerance=2.0, keep_blank_chars=False)
    words.sort(key=lambda w: (w["top"], w["x0"]))
    lines, cur, cur_top = [], [], None
    for w in words:
        if cur_top is None or abs(w["top"] - cur_top) <= 2.5:
            cur.append(w)
            if cur_top is None:
                cur_top = w["top"]
        else:
            lines.append(sorted(cur, key=lambda x: x["x0"]))
            cur, cur_top = [w], w["top"]
    if cur:
        lines.append(sorted(cur, key=lambda x: x["x0"]))
    return lines


def find_header(lines):
    best, best_count = None, 0
    for line in lines:
        cols = {}
        for w in line:
            col = match_column(w["text"])
            if col and col not in cols:
                cols[col] = (w["x0"] + w["x1"]) / 2.0
        if len(cols) >= 3 and any(c in cols for c in ("debit", "credit", "balance")):
            if len(cols) > best_count:
                best, best_count = cols, len(cols)
    return best


def build_intervals(anchors, page_width):
    ordered = sorted(anchors.items(), key=lambda kv: kv[1])
    intervals = []
    for idx, (col, x) in enumerate(ordered):
        left = 0 if idx == 0 else (ordered[idx - 1][1] + x) / 2.0
        right = page_width if idx == len(ordered) - 1 else (x + ordered[idx + 1][1]) / 2.0
        intervals.append((col, left, right))
    return intervals


def assign_columns(line, intervals):
    cells = {col: [] for col, _, _ in intervals}
    for w in line:
        center = (w["x0"] + w["x1"]) / 2.0
        for col, left, right in intervals:
            if left <= center < right:
                cells[col].append(w["text"].strip())
                break
    return {col: " ".join(v).strip() for col, v in cells.items()}


def parse_words_strategy(pdf):
    rows, warnings = [], []
    anchors = None
    for page_no, page in enumerate(pdf.pages, start=1):
        lines = group_lines(page)
        page_anchors = find_header(lines)
        if page_anchors:
            anchors = page_anchors
        if not anchors:
            continue
        intervals = build_intervals(anchors, page.width)

        prev_was_data = False
        for line in lines:
            raw = " ".join(w["text"] for w in line).strip()
            if not raw:
                continue
            raw_low = raw.lower()
            header_hits = sum(1 for w in line if match_column(w["text"]))
            if header_hits >= 3:
                prev_was_data = False
                continue
            if any(s in raw_low for s in SKIP_WORDS):
                prev_was_data = False
                continue

            cells = assign_columns(line, intervals)
            row = build_row(cells, raw_low)
            if row:
                rows.append(row)
                prev_was_data = True
            elif prev_was_data and is_continuation(cells) and rows:
                extra = (cells.get("description", "") + " " + cells.get("id", "")).strip()
                rows[-1]["description"] = (rows[-1]["description"] + " " + extra).strip()
                if rows[-1]["row_type"] == "transaction" and not rows[-1]["id"]:
                    rows[-1]["id"] = extract_row_id("", rows[-1]["description"])
            elif len(raw) > 3:
                warnings.append(f"page {page_no}: unclassified line: {raw[:120]}")
                prev_was_data = False
    return rows, warnings, anchors is not None


def parse_tables_strategy(pdf):
    rows, warnings = [], []
    col_map = None
    found_any = False
    for page in pdf.pages:
        for table in page.extract_tables():
            prev_was_data = False
            for cells in table:
                cells = [(c or "").replace("\n", " ").strip() for c in cells]
                mapped = {}
                for idx, c in enumerate(cells):
                    col = match_column(c)
                    if col and col not in mapped:
                        mapped[col] = idx
                if len(mapped) >= 3 and any(c in mapped for c in ("debit", "credit", "balance")):
                    col_map = mapped
                    found_any = True
                    prev_was_data = False
                    continue
                if not col_map:
                    continue

                def get(field):
                    i = col_map.get(field, -1)
                    return cells[i] if 0 <= i < len(cells) else ""

                celld = {f: get(f) for f in COLUMN_KEYWORDS}
                raw_low = " ".join(cells).lower()
                if any(s in raw_low for s in SKIP_WORDS):
                    prev_was_data = False
                    continue
                row = build_row(celld, raw_low)
                if row:
                    rows.append(row)
                    prev_was_data = True
                elif prev_was_data and is_continuation(celld) and rows:
                    extra = (celld.get("description", "") + " " + celld.get("id", "")).strip()
                    if extra:
                        rows[-1]["description"] = (rows[-1]["description"] + " " + extra).strip()
    return rows, warnings, found_any


def parse_pdf(file_bytes):
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        pages = len(pdf.pages)
        total_chars = sum(len(p.chars) for p in pdf.pages)
        if total_chars < 20:
            raise ValueError(
                "This PDF has no text layer (it is a scan/image). "
                "It needs OCR before it can be parsed without an LLM."
            )
        t_rows, t_warn, t_found = parse_tables_strategy(pdf)
        w_rows, w_warn, w_found = parse_words_strategy(pdf)

    # Ruled-table extraction reads each printed cell directly and is immune
    # to the column-bleed that can happen with the word-position strategy
    # (e.g. an address like "CHTAURA-" running into the next column with no
    # gap). Prefer it whenever the PDF actually draws table lines; only fall
    # back to word-position matching when no ruled table is found.
    if t_found and t_rows:
        strategy, rows, warnings = "table", t_rows, t_warn
    elif w_found:
        strategy, rows, warnings = "words", w_rows, w_warn
    else:
        raise ValueError(
            "Could not find a recognizable column header (Date/Debit/Credit/"
            "Balance or Arabic equivalents) in this PDF. This supplier's "
            "format needs a keyword added to COLUMN_KEYWORDS."
        )
    warnings = [f"[pdf/{strategy}] {w}" for w in warnings]
    if not rows:
        warnings.append(f"[pdf/{strategy}] Header found but no data rows extracted.")
    return rows, warnings, {"pages": pages}


# ================================================================== XLSX

def cell_to_text(v):
    """Excel cells already carry typed values (dates, numbers) - normalize
    everything to the string form the shared header/amount matchers expect,
    except dates which are converted straight to ISO to avoid ambiguity."""
    if v is None:
        return ""
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def map_sheet_header(header_row):
    """Map each column index to a field name. Unlike the PDF path, a
    spreadsheet can have several id-like columns on one row (Trans Id,
    Invoice number, Check Num); keep all of them ranked so build_sheet_row
    can pick the most specific one actually populated on each row, instead
    of only keeping whichever appeared first."""
    mapped = {}       # field -> single column index (date/description/debit/credit/balance)
    id_candidates = []  # list of (rank, column index) for every id-like header
    for idx, cell in enumerate(header_row):
        col = match_column(cell)
        if col == "id":
            id_candidates.append((id_column_rank(cell), idx))
        elif col and col not in mapped:
            mapped[col] = idx
    id_candidates.sort()
    return mapped, id_candidates


def build_sheet_row(row_values, mapped, id_candidates):
    def get(field):
        i = mapped.get(field, -1)
        return cell_to_text(row_values[i]) if 0 <= i < len(row_values) else ""

    cells = {f: get(f) for f in ("date", "description", "debit", "credit", "balance")}
    # pick the first (highest-priority) id column that actually has a value
    # on this specific row, since e.g. "Check Num" is usually blank except
    # on cheque rows while "Invoice number" is blank on those same rows
    id_text = ""
    for _, idx in id_candidates:
        if idx < len(row_values):
            v = cell_to_text(row_values[idx])
            if v:
                id_text = v
                break
    cells["id"] = id_text

    raw_low = " ".join(cell_to_text(v) for v in row_values if v is not None).lower()
    return build_row(cells, raw_low), cells


def parse_xlsx(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    rows, warnings = [], []
    sheets_used = 0

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        all_rows = list(ws.iter_rows(values_only=True))
        if not all_rows:
            continue

        header_idx, mapped, id_candidates = None, None, None
        for i, r in enumerate(all_rows[:5]):
            m, ids = map_sheet_header(r)
            if len(m) + (1 if ids else 0) >= 3 and any(c in m for c in ("debit", "credit", "balance")):
                header_idx, mapped, id_candidates = i, m, ids
                break
        if header_idx is None:
            continue  # this sheet doesn't look like a ledger - skip quietly

        sheets_used += 1
        prev_row_ref = None
        for r in all_rows[header_idx + 1:]:
            if r is None or all(v is None for v in r):
                continue
            row, cells = build_sheet_row(r, mapped, id_candidates)
            if row:
                rows.append(row)
                prev_row_ref = rows[-1]
            elif prev_row_ref is not None and is_continuation(cells):
                extra = (cells.get("description", "") + " " + cells.get("id", "")).strip()
                if extra:
                    prev_row_ref["description"] = (prev_row_ref["description"] + " " + extra).strip()

    if sheets_used == 0:
        raise ValueError(
            "Could not find a recognizable column header (Date/Debit/Credit/"
            "Balance or similar) in any sheet of this workbook. This "
            "supplier's format needs a keyword added to COLUMN_KEYWORDS."
        )
    if not rows:
        warnings.append("[xlsx] Header(s) found but no data rows extracted.")
    return rows, warnings, {"sheets": len(wb.sheetnames), "sheets_used": sheets_used}


# ================================================================== CSV

def parse_csv(file_bytes):
    text = None
    for encoding in ("utf-8-sig", "utf-8", "cp1256", "latin-1"):
        try:
            text = file_bytes.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError("Could not decode this CSV file as text.")

    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
    except csv.Error:
        dialect = csv.excel
    reader = list(csv.reader(io.StringIO(text), dialect))
    reader = [r for r in reader if any(c.strip() for c in r)]
    if not reader:
        raise ValueError("This CSV file has no rows.")

    header_idx, mapped, id_candidates = None, None, None
    for i, r in enumerate(reader[:5]):
        m, ids = map_sheet_header(r)
        if len(m) + (1 if ids else 0) >= 3 and any(c in m for c in ("debit", "credit", "balance")):
            header_idx, mapped, id_candidates = i, m, ids
            break
    if header_idx is None:
        raise ValueError(
            "Could not find a recognizable column header (Date/Debit/Credit/"
            "Balance or similar) in this CSV. This supplier's format needs "
            "a keyword added to COLUMN_KEYWORDS."
        )

    rows, warnings = [], []
    prev_row_ref = None
    for r in reader[header_idx + 1:]:
        row, cells = build_sheet_row(r, mapped, id_candidates)
        if row:
            rows.append(row)
            prev_row_ref = rows[-1]
        elif prev_row_ref is not None and is_continuation(cells):
            extra = (cells.get("description", "") + " " + cells.get("id", "")).strip()
            if extra:
                prev_row_ref["description"] = (prev_row_ref["description"] + " " + extra).strip()

    if not rows:
        warnings.append("[csv] Header found but no data rows extracted.")
    return rows, warnings, {}


# ================================================================== dispatch

def sniff_format(file_bytes, filename):
    """Identify the real file format from its bytes first (magic numbers
    can't be spoofed by a wrong extension), falling back to the filename
    extension only when the bytes are inconclusive (plain text formats)."""
    if file_bytes.startswith(b"%PDF"):
        return "pdf"
    if file_bytes.startswith(b"PK\x03\x04"):
        return "xlsx"     # xlsx/xlsm are zip containers
    if file_bytes.startswith(b"\xd0\xcf\x11\xe0"):
        raise ValueError(
            "This looks like a legacy .xls file (pre-2007 Excel format). "
            "Please re-save it as .xlsx or .csv and upload again."
        )
    ext = (filename or "").rsplit(".", 1)[-1].lower() if filename else ""
    if ext in ("csv", "tsv", "txt"):
        return "csv"
    if ext in ("xlsx", "xlsm"):
        return "xlsx"
    if ext == "pdf":
        return "pdf"
    # last resort: does it decode as text at all?
    try:
        file_bytes[:2048].decode("utf-8")
        return "csv"
    except UnicodeDecodeError:
        raise ValueError(
            "Could not identify this file's format. Supported formats: "
            "PDF, Excel (.xlsx), and CSV."
        )


def parse_supplier_file(file_bytes, filename=None):
    fmt = sniff_format(file_bytes, filename)
    if fmt == "pdf":
        rows, warnings, meta = parse_pdf(file_bytes)
    elif fmt == "xlsx":
        rows, warnings, meta = parse_xlsx(file_bytes)
    else:
        rows, warnings, meta = parse_csv(file_bytes)
    result = {"rows": rows, "warnings": warnings, "build_tag": BUILD_TAG, "format": fmt}
    result.update(meta)
    return result


class handler(BaseHTTPRequestHandler):
    def _send(self, status, payload):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            data = json.loads(self.rfile.read(length) or b"{}")
            b64 = data.get("file_base64") or data.get("pdf_base64", "")
            filename = data.get("filename", "")
            if "," in b64[:80]:
                b64 = b64.split(",", 1)[1]
            if not b64:
                return self._send(400, {"error": "file_base64 is required."})
            file_bytes = base64.b64decode(b64)
            self._send(200, parse_supplier_file(file_bytes, filename))
        except json.JSONDecodeError:
            self._send(400, {"error": "Invalid JSON body."})
        except ValueError as e:
            self._send(422, {"error": str(e)})
        except Exception as e:  # noqa: BLE001
            self._send(500, {"error": f"Extraction failed: {e}"})
