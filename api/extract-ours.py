"""
POST /api/extract-ours
Body: { "pdf_base64": "<base64 of the doPDF statement>" }
Deterministic parser for our fixed-template statement of account.
Columns: Co, Br, Date, Type, OTL, VCHNO, Description, Debit, Credit, Balance.
No LLM, no external calls. Uses word x/y coordinates from pdfplumber, because
adjacent fields are packed with no spaces (e.g. "WD03" = Co+Br,
"RJV0320260504" = Type+OTL+VCHNO).

Response: { "rows": [...], "warnings": [...], "pages": n }
Row schema: { date, id, description, debit, credit, balance, row_type }
row_type: "transaction" | "opening_balance" | "closing_balance"
"""

from http.server import BaseHTTPRequestHandler
import base64
import io
import json
import re

import pdfplumber

# ---------------------------------------------------------------- regexes

DATE_RE = re.compile(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$")
COBR_RE = re.compile(r"^([A-Za-z]{2})(\d{2})$")              # packed Co+Br
TYPEOTL_RE = re.compile(r"^([A-Za-z]{3})(\d{2})(\S*)$")      # packed Type+OTL+VCHNO
AMOUNT_RE = re.compile(r"^\(?-?[\d,]+(?:\.\d+)?\)?(CR|DR)?$", re.IGNORECASE)
PURE_DIGITS_RE = re.compile(r"^[\d\s./,\-]+$")               # continuation lines

OPENING_WORDS = ("opening", "b/f", "b.f", "brought")
CLOSING_WORDS = ("closing", "c/f", "c.f", "carried")
SKIP_WORDS = (
    "statement of account", "page ", "printed", "dopdf", "print to pdf",
    "from :", "to :", "from:", "to:",
)

LINE_Y_TOLERANCE = 2.5   # words whose tops differ by <= this are one line
WORD_X_TOLERANCE = 1.0   # gap (pt) below which adjacent chars merge into a word


# ---------------------------------------------------------------- helpers

def parse_amount(token):
    """'1,234.56CR' -> (1234.56, 'CR'); '450.00' -> (450.0, None)."""
    m = AMOUNT_RE.match(token)
    if not m:
        return None, None
    suffix = m.group(1).upper() if m.group(1) else None
    num = token[: len(token) - 2] if suffix else token
    num = num.replace(",", "").replace("(", "-").replace(")", "")
    try:
        return float(num), suffix
    except ValueError:
        return None, None


def normalize_date(token):
    """dd/mm/yyyy (or dd-mm-yy) -> YYYY-MM-DD. Returns None if not a date."""
    m = DATE_RE.match(token)
    if not m:
        return None
    d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y < 100:
        y += 2000
    if not (1 <= d <= 31 and 1 <= mo <= 12):
        return None
    return f"{y:04d}-{mo:02d}-{d:02d}"


def group_lines(page):
    """All words on the page, grouped into visual lines by their y position."""
    words = page.extract_words(
        x_tolerance=WORD_X_TOLERANCE, y_tolerance=2.0, keep_blank_chars=False
    )
    words.sort(key=lambda w: (w["top"], w["x0"]))
    lines, current, current_top = [], [], None
    for w in words:
        if current_top is None or abs(w["top"] - current_top) <= LINE_Y_TOLERANCE:
            current.append(w)
            if current_top is None:
                current_top = w["top"]
        else:
            lines.append(sorted(current, key=lambda x: x["x0"]))
            current, current_top = [w], w["top"]
    if current:
        lines.append(sorted(current, key=lambda x: x["x0"]))
    return lines


def find_anchors(lines):
    """x-centers of the Debit / Credit / Balance header words on this page."""
    for line in lines:
        found = {}
        for w in line:
            t = w["text"].strip().lower()
            if t in ("debit", "credit", "balance"):
                found[t] = (w["x0"] + w["x1"]) / 2.0
        if len(found) == 3:
            return found
    return None


def nearest_anchor(center, anchors, threshold):
    best_col, best_dist = None, None
    for col, ax in anchors.items():
        d = abs(center - ax)
        if best_dist is None or d < best_dist:
            best_col, best_dist = col, d
    return best_col if best_dist is not None and best_dist <= threshold else None


def split_amounts(line, anchors, threshold):
    """Assign amount-looking tokens to debit/credit/balance by x position.
    Returns (amounts dict, balance suffix, remaining words)."""
    amounts, suffix, rest = {}, None, []
    for w in line:
        token = w["text"].strip()
        center = (w["x0"] + w["x1"]) / 2.0
        # standalone CR/DR token printed after the balance figure
        if token.upper() in ("CR", "DR") and anchors and center >= anchors["balance"] - threshold:
            suffix = token.upper()
            continue
        val, suf = parse_amount(token)
        col = nearest_anchor(center, anchors, threshold) if (val is not None and anchors) else None
        if col:
            amounts[col] = val
            if col == "balance" and suf:
                suffix = suf
        else:
            rest.append(w)
    return amounts, suffix, rest


def parse_leading_fields(tokens):
    """Consume Co, Br, Date, Type, OTL, VCHNO from the start of the token list.
    Fields may arrive packed ('WD03', 'RJV0320260504') or split by the word
    extractor. Returns (fields dict, remaining description tokens) or None."""
    i, n = 0, len(tokens)
    if n == 0:
        return None

    # --- Co + Br
    m = COBR_RE.match(tokens[0])
    if m:
        co, br = m.group(1), m.group(2)
        i = 1
    elif n >= 2 and re.fullmatch(r"[A-Za-z]{2}", tokens[0]) and re.fullmatch(r"\d{2}", tokens[1]):
        co, br, i = tokens[0], tokens[1], 2
    else:
        return None

    # --- Date
    if i >= n:
        return None
    date = normalize_date(tokens[i])
    if not date:
        return None
    i += 1

    # --- Type + OTL + VCHNO
    if i >= n:
        return None
    m = TYPEOTL_RE.match(tokens[i])
    if not m:
        return None
    ttype, otl, vchno = m.group(1), m.group(2), m.group(3)
    i += 1
    if not vchno and i < n and re.fullmatch(r"[\w\-/]+", tokens[i]):
        vchno = tokens[i]
        i += 1

    return {"co": co, "br": br, "date": date, "type": ttype, "otl": otl, "vchno": vchno}, tokens[i:]


# ---------------------------------------------------------------- main parse

def parse_pdf(pdf_bytes):
    rows, warnings = [], []
    anchors = None  # header positions persist across pages
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page_count = len(pdf.pages)
        for page_no, page in enumerate(pdf.pages, start=1):
            lines = group_lines(page)
            page_anchors = find_anchors(lines)
            if page_anchors:
                anchors = page_anchors
            if anchors:
                xs = sorted(anchors.values())
                gap = min(xs[1] - xs[0], xs[2] - xs[1])
                threshold = max(25.0, gap * 0.75)
            else:
                threshold = 0.0
                if page_no == 1:
                    warnings.append(
                        "Debit/Credit/Balance header row not found on page 1 - "
                        "amount columns cannot be located reliably."
                    )

            for line in lines:
                raw = " ".join(w["text"] for w in line).strip()
                if not raw:
                    continue
                low = raw.lower()

                # header / footer / boilerplate
                if "debit" in low and "credit" in low and "balance" in low:
                    continue
                if any(s in low for s in SKIP_WORDS):
                    continue

                amounts, suffix, rest = split_amounts(line, anchors or {}, threshold)
                rest_tokens = [w["text"].strip() for w in rest if w["text"].strip()]
                rest_text = " ".join(rest_tokens)

                # opening / closing balance rows
                if any(k in low for k in OPENING_WORDS) or any(k in low for k in CLOSING_WORDS):
                    row_type = "opening_balance" if any(k in low for k in OPENING_WORDS) else "closing_balance"
                    date = next((normalize_date(t) for t in rest_tokens if normalize_date(t)), None)
                    balance = amounts.get("balance")
                    if balance is None:
                        # balance-only rows sometimes print outside the located columns
                        for t in reversed(rest_tokens):
                            val, suf = parse_amount(t)
                            if val is not None:
                                balance, suffix = val, suf or suffix
                                break
                    rows.append({
                        "date": date or "",
                        "id": "",
                        "description": (rest_text + (f" {suffix}" if suffix else "")).strip(),
                        "debit": amounts.get("debit", 0.0) or 0.0,
                        "credit": amounts.get("credit", 0.0) or 0.0,
                        "balance": balance if balance is not None else 0.0,
                        "row_type": row_type,
                    })
                    continue

                # transaction row
                parsed = parse_leading_fields(rest_tokens)
                if parsed:
                    fields, desc_tokens = parsed
                    rows.append({
                        "date": fields["date"],
                        "id": fields["vchno"],
                        "description": " ".join(desc_tokens).strip(),
                        "debit": amounts.get("debit", 0.0) or 0.0,
                        "credit": amounts.get("credit", 0.0) or 0.0,
                        "balance": amounts.get("balance", 0.0) or 0.0,
                        "row_type": "transaction",
                        "_meta": {"co": fields["co"], "br": fields["br"],
                                  "type": fields["type"], "otl": fields["otl"],
                                  "balance_side": suffix},
                    })
                    continue

                # continuation line: pure digits, no date, belongs to previous row
                has_date = any(normalize_date(t) for t in rest_tokens)
                if not has_date and rows and PURE_DIGITS_RE.match(raw):
                    rows[-1]["description"] = (rows[-1]["description"] + " " + raw).strip()
                    continue

                # nothing matched: report instead of silently dropping
                warnings.append(f"page {page_no}: unclassified line: {raw}")

    return {"rows": rows, "warnings": warnings, "pages": page_count}


# ---------------------------------------------------------------- handler

class handler(BaseHTTPRequestHandler):
    def _send(self, status, payload):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        try:
            length = int(self.headers.get("Content-Length", 0))
            data = json.loads(self.rfile.read(length) or b"{}")
            b64 = data.get("pdf_base64", "")
            if "," in b64[:80]:  # tolerate a data: URL prefix
                b64 = b64.split(",", 1)[1]
            pdf_bytes = base64.b64decode(b64)
            if not pdf_bytes.startswith(b"%PDF"):
                return self._send(400, {"error": "File does not look like a PDF."})
            result = parse_pdf(pdf_bytes)
            self._send(200, result)
        except json.JSONDecodeError:
            self._send(400, {"error": "Invalid JSON body."})
        except Exception as e:  # noqa: BLE001
            self._send(500, {"error": f"Parse failed: {e}"})
