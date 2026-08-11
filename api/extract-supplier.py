"""
POST /api/extract-supplier
Deterministic supplier-statement parser - no LLM, no external calls.
Finds a header line with recognizable column keywords (English or Arabic),
uses header positions as column anchors, and buckets every word into a
column. Also runs pdfplumber's ruled-table extraction; the strategy that
recognizes a header and yields more rows wins.

Row id priority: INV/C\\N number inside the description (this is the number
that also appears on our statement as "Inv#"), else the reference column.
Dates like "Jan 12, 2026" and negative-signed amounts are normalized.
"""

from http.server import BaseHTTPRequestHandler
import base64
import io
import json
import re

import pdfplumber

COLUMN_KEYWORDS = {
    "date": ["date", "تاريخ", "التاريخ"],
    "id": ["ref", "reference", "invoice", "inv", "voucher", "vch", "doc",
           "document", "no", "number", "num", "trx", "nb", "رقم", "سند",
           "مستند", "فاتورة", "المرجع", "مرجع"],
    "description": ["description", "details", "narration", "particulars",
                    "memo", "remarks", "بيان", "البيان", "التفاصيل",
                    "الوصف", "ملاحظات"],
    "debit": ["debit", "dr", "مدين"],
    "credit": ["credit", "cr", "دائن"],
    "balance": ["balance", "رصيد", "الرصيد"],
}

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
    t = text.strip().lower().strip(":.")
    for col, words in COLUMN_KEYWORDS.items():
        if t in words:
            return col
    parts = [p.strip(":.") for p in t.split()]
    for col, words in COLUMN_KEYWORDS.items():
        if any(p in words for p in parts):
            return col
    return None


def parse_amount(token):
    token = token.strip()
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
    m = REF_RE.search(description)
    if m:
        return m.group(1)
    for tok in id_cell.split():
        if re.search(r"\d", tok) and not find_date(tok):
            return tok
    return ""


def build_row(cells, raw_low):
    date = find_date(cells.get("date", "")) or ""
    debit = parse_amount(cells.get("debit", ""))
    credit = parse_amount(cells.get("credit", ""))
    balance = parse_amount(cells.get("balance", ""))

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
        # sign is presentation (this supplier prints credits negative)
        "debit": abs(debit) if debit is not None else 0.0,
        "credit": abs(credit) if credit is not None else 0.0,
        "balance": balance if balance is not None else 0.0,
        "row_type": row_type,
    }


def is_continuation(cells):
    if find_date(cells.get("date", "")):
        return False
    if any(parse_amount(cells.get(c, "")) is not None for c in ("debit", "credit", "balance")):
        return False
    return bool(cells.get("description") or cells.get("id"))


# ------------------------------------------------------------ word strategy

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


# ------------------------------------------------------------ table strategy

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


def parse_pdf(pdf_bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        pages = len(pdf.pages)
        total_chars = sum(len(p.chars) for p in pdf.pages)
        if total_chars < 20:
            raise ValueError(
                "This PDF has no text layer (it is a scan/image). "
                "It needs OCR before it can be parsed without an LLM."
            )
        t_rows, t_warn, t_found = parse_tables_strategy(pdf)
        w_rows, w_warn, w_found = parse_words_strategy(pdf)

    candidates = []
    if t_found:
        candidates.append(("table", t_rows, t_warn))
    if w_found:
        candidates.append(("words", w_rows, w_warn))
    if not candidates:
        raise ValueError(
            "Could not find a recognizable column header (Date/Debit/Credit/"
            "Balance or Arabic equivalents). This supplier's format needs a "
            "keyword added to COLUMN_KEYWORDS in extract-supplier.py."
        )
    strategy, rows, warnings = max(candidates, key=lambda c: len(c[1]))
    warnings = [f"[{strategy} strategy] {w}" for w in warnings]
    if not rows:
        warnings.append(f"[{strategy} strategy] Header found but no data rows extracted.")
    return {"rows": rows, "warnings": warnings, "pages": pages}


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
            b64 = data.get("pdf_base64", "")
            if "," in b64[:80]:
                b64 = b64.split(",", 1)[1]
            pdf_bytes = base64.b64decode(b64)
            if not pdf_bytes.startswith(b"%PDF"):
                return self._send(400, {"error": "File does not look like a PDF."})
            self._send(200, parse_pdf(pdf_bytes))
        except json.JSONDecodeError:
            self._send(400, {"error": "Invalid JSON body."})
        except ValueError as e:
            self._send(422, {"error": str(e)})
        except Exception as e:  # noqa: BLE001
            self._send(500, {"error": f"Extraction failed: {e}"})
