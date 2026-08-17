"""
POST /api/extract-ours
Deterministic parser for our fixed-template doPDF statement (accsta04).
Columns: Co, Br, Date, Type, OTL, VCHNO, Description, Debit, Credit, Balance.
Fields may be packed ("WD03") or separate words ("RJV 03 20260504").
The supplier's invoice number ("Inv# 15905326" in the description, often on a
wrapped continuation line) is extracted as the row id, since VCHNO repeats
per day and never appears on supplier statements. Falls back to VCHNO.
No LLM, no external calls.
"""

from http.server import BaseHTTPRequestHandler
import base64
import io
import json
import re

import pdfplumber

BUILD_TAG = "2026-08-14-cn-ref"

DATE_RE = re.compile(r"^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})$")
COBR_RE = re.compile(r"^([A-Za-z]{2})(\d{2})$")
TYPEOTL_RE = re.compile(r"^([A-Za-z]{3})(\d{1,3})(\S*)$")
AMOUNT_RE = re.compile(r"^\(?-?(?:[\d,]+(?:\.\d+)?|\.\d+)\)?(CR|DR)?$", re.IGNORECASE)
PURE_DIGITS_RE = re.compile(r"^[\d\s./,\-]+$")
INV_ID_RE = re.compile(r"(?:inv\s*#|c\s*/?\s*n\s*#)\s*(\d+)", re.IGNORECASE)

OPENING_WORDS = ("opening", "balance until", "b/f", "brought")
CLOSING_WORDS = ("closing", "balance as at", "c/f", "carried")
SKIP_WORDS = ("statement of account", "page:", "page ", "printed", "dopdf",
              "a/c no", "accsta", "orig.curr", "requested curr")

LINE_Y_TOLERANCE = 2.5
WORD_X_TOLERANCE = 1.0


def parse_amount(token):
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
    words = page.extract_words(x_tolerance=WORD_X_TOLERANCE, y_tolerance=2.0,
                               keep_blank_chars=False)
    words.sort(key=lambda w: (w["top"], w["x0"]))
    lines, cur, cur_top = [], [], None
    for w in words:
        if cur_top is None or abs(w["top"] - cur_top) <= LINE_Y_TOLERANCE:
            cur.append(w)
            if cur_top is None:
                cur_top = w["top"]
        else:
            lines.append(sorted(cur, key=lambda x: x["x0"]))
            cur, cur_top = [w], w["top"]
    if cur:
        lines.append(sorted(cur, key=lambda x: x["x0"]))
    return lines


def find_anchors(lines):
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
    amounts, suffix, rest = {}, None, []
    for w in line:
        token = w["text"].strip()
        center = (w["x0"] + w["x1"]) / 2.0
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
    """Co, Br, Date, Type, OTL, VCHNO from the start of the token list.
    Handles both packed ('WD03', 'RJV0320260504') and separate words
    ('WD', '03', ..., 'RJV', '03', '20260504')."""
    i, n = 0, len(tokens)
    if n == 0:
        return None

    m = COBR_RE.match(tokens[0])
    if m:
        co, br = m.group(1), m.group(2)
        i = 1
    elif n >= 2 and re.fullmatch(r"[A-Za-z]{2}", tokens[0]) and re.fullmatch(r"\d{2}", tokens[1]):
        co, br, i = tokens[0], tokens[1], 2
    else:
        return None

    if i >= n:
        return None
    date = normalize_date(tokens[i])
    if not date:
        return None
    i += 1

    if i >= n:
        return None
    m = TYPEOTL_RE.match(tokens[i])
    if m and m.group(3):                       # fully packed
        ttype, otl, vchno = m.group(1), m.group(2), m.group(3)
        i += 1
    elif m:                                    # Type+OTL packed, VCHNO separate
        ttype, otl = m.group(1), m.group(2)
        i += 1
        vchno = ""
        if i < n and re.fullmatch(r"[\w\-/]+", tokens[i]):
            vchno = tokens[i]
            i += 1
    elif re.fullmatch(r"[A-Za-z]{3}", tokens[i]):   # all separate
        ttype = tokens[i]
        i += 1
        otl = ""
        if i + 1 < n and re.fullmatch(r"\d{2}", tokens[i]) and re.fullmatch(r"[\w\-/]{3,}", tokens[i + 1]):
            otl = tokens[i]
            i += 1
        vchno = ""
        if i < n and re.fullmatch(r"[\w\-/]+", tokens[i]):
            vchno = tokens[i]
            i += 1
    else:
        return None

    return {"co": co, "br": br, "date": date, "type": ttype, "otl": otl, "vchno": vchno}, tokens[i:]


def parse_pdf(pdf_bytes):
    rows, warnings = [], []
    anchors = None
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
                    warnings.append("Debit/Credit/Balance header row not found on page 1.")

            prev_was_data = False
            for line in lines:
                raw = " ".join(w["text"] for w in line).strip()
                if not raw:
                    continue
                low = raw.lower()

                if "debit" in low and "credit" in low and "balance" in low:
                    prev_was_data = False
                    continue
                if any(s in low for s in SKIP_WORDS):
                    continue

                amounts, suffix, rest = split_amounts(line, anchors or {}, threshold)
                rest_tokens = [w["text"].strip() for w in rest if w["text"].strip()]
                rest_text = " ".join(rest_tokens)

                # Try parsing as a normal transaction FIRST. Only treat a line
                # as an opening/closing balance marker if it does NOT parse as
                # a transaction - otherwise a real transaction whose free-text
                # description happens to mention "opening balance" (e.g. a
                # payment memo referencing an opening-balance settlement)
                # would be wrongly excluded from matching.
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
                                  "vchno": fields["vchno"], "balance_side": suffix},
                    })
                    prev_was_data = True
                    continue

                # Fell through: not a parseable transaction line. If it
                # mentions opening/closing balance, treat it as that special
                # summary row (these lack normal Co/Br/Type/VCHNO fields
                # entirely, e.g. "Balance Until 01/01/2026 ........").
                if any(k in low for k in OPENING_WORDS) or any(k in low for k in CLOSING_WORDS):
                    row_type = "opening_balance" if any(k in low for k in OPENING_WORDS) else "closing_balance"
                    date = next((normalize_date(t) for t in rest_tokens if normalize_date(t)), None)
                    balance = amounts.get("balance")
                    if balance is None:
                        for t in reversed(rest_tokens):
                            val, suf = parse_amount(t)
                            if val is not None:
                                balance, suffix = val, suf or suffix
                                break
                    rows.append({
                        "date": date or "",
                        "id": "",
                        "description": rest_text,
                        "debit": amounts.get("debit", 0.0) or 0.0,
                        "credit": amounts.get("credit", 0.0) or 0.0,
                        "balance": balance if balance is not None else 0.0,
                        "row_type": row_type,
                    })
                    prev_was_data = True
                    continue

                has_date = any(normalize_date(t) for t in rest_tokens)
                if not has_date and prev_was_data and rows and PURE_DIGITS_RE.match(raw):
                    rows[-1]["description"] = (rows[-1]["description"] + " " + raw).strip()
                    continue

                warnings.append(f"page {page_no}: unclassified line: {raw[:120]}")
                prev_was_data = False

    # the supplier-facing reference is the Inv# inside the description
    for r in rows:
        if r["row_type"] == "transaction":
            m = INV_ID_RE.search(r["description"])
            if m:
                r["id"] = m.group(1)

    return {"rows": rows, "warnings": warnings, "pages": page_count, "build_tag": BUILD_TAG}


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
        except Exception as e:  # noqa: BLE001
            self._send(500, {"error": f"Parse failed: {e}"})
