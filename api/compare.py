"""
POST /api/compare
Body: { "ours": [rows], "supplier": [rows] }   (row schema from the extractors)

Matching, plain code, no LLM:
  1. exact match on transaction id, normalized (strip leading zeros,
     whitespace, lowercase), one-to-one
  2. fallback for the leftovers: same date + a debit or credit amount that
     matches (within 0.01) in either of the supplier's columns
Opening/closing balance rows are excluded. Only problem rows are returned:
  missing_in_supplier | missing_in_tahan | value_mismatch
"""

from http.server import BaseHTTPRequestHandler
import json
import re
from datetime import datetime

TOLERANCE = 0.01

# Supplier statements are written from the supplier's perspective, so their
# debit column is usually our credit and vice versa. With MIRROR_OK a row
# whose amounts match cross-wise (our debit == their credit) counts as
# matched. Set to False to require same-side matches only.
MIRROR_OK = True


def norm_id(v):
    s = re.sub(r"\s+", "", str(v or "")).lower()
    return s.lstrip("0")


def norm_date(v):
    s = str(v or "").strip()
    if not s:
        return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return s  # unknown format: compare as-is


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def close(a, b):
    return abs(a - b) <= TOLERANCE + 1e-9


def values_match(o, s):
    direct = close(o["debit"], s["debit"]) and close(o["credit"], s["credit"])
    mirrored = MIRROR_OK and close(o["debit"], s["credit"]) and close(o["credit"], s["debit"])
    return direct or mirrored


def clean(rows):
    out = []
    for r in rows or []:
        out.append({
            "date": norm_date(r.get("date")),
            "id": str(r.get("id") or ""),
            "nid": norm_id(r.get("id")),
            "description": str(r.get("description") or ""),
            "debit": num(r.get("debit")),
            "credit": num(r.get("credit")),
            "row_type": r.get("row_type", "transaction"),
        })
    return [r for r in out if r["row_type"] == "transaction"]


def issue(kind, o=None, s=None):
    src = o or s
    return {
        "issue": kind,
        "date": src["date"],
        "id": src["id"],
        "description": src["description"],
        "our_debit": o["debit"] if o else None,
        "our_credit": o["credit"] if o else None,
        "supplier_debit": s["debit"] if s else None,
        "supplier_credit": s["credit"] if s else None,
    }


def compare(ours_raw, supplier_raw):
    ours = clean(ours_raw)
    supplier = clean(supplier_raw)

    matched_pairs = []
    s_unmatched = list(range(len(supplier)))
    o_unmatched = []

    # pass 1: normalized id, one-to-one
    by_id = {}
    for j in s_unmatched:
        nid = supplier[j]["nid"]
        if nid:
            by_id.setdefault(nid, []).append(j)
    taken = set()
    for i, o in enumerate(ours):
        j = None
        if o["nid"] and by_id.get(o["nid"]):
            j = by_id[o["nid"]].pop(0)
        if j is not None:
            matched_pairs.append((i, j))
            taken.add(j)
        else:
            o_unmatched.append(i)
    s_unmatched = [j for j in s_unmatched if j not in taken]

    # pass 2: same date + amount within tolerance in either supplier column
    still = []
    for i in o_unmatched:
        o = ours[i]
        amt = o["debit"] if o["debit"] > 0 else o["credit"]
        j_found = None
        if amt > 0 and o["date"]:
            for j in s_unmatched:
                s = supplier[j]
                if s["date"] == o["date"] and (close(amt, s["debit"]) or close(amt, s["credit"])):
                    j_found = j
                    break
        if j_found is not None:
            matched_pairs.append((i, j_found))
            s_unmatched.remove(j_found)
        else:
            still.append(i)
    o_unmatched = still

    # Informational only: the date range actually present in our file. Every
    # unmatched/mismatched row is still reported regardless of its date -
    # nothing is ever hidden based on period. "out_of_our_range" just flags
    # supplier rows dated outside that range, since those are commonly rows
    # the supplier will also show on an adjacent statement (e.g. a January
    # invoice appearing on a "Jan to date" export when we only sent May).
    def date_range(rows):
        ds = [r["date"] for r in rows if r["date"]]
        return (min(ds), max(ds)) if ds else None

    our_range = date_range(ours)

    def out_of_our_range(r):
        if our_range is None or not r["date"]:
            return False
        return not (our_range[0] <= r["date"] <= our_range[1])

    issues = []
    matched_rows = []
    mismatches = 0
    for i, j in matched_pairs:
        o, s = ours[i], supplier[j]
        if values_match(o, s):
            matched_rows.append({
                "date": o["date"],
                "id": o["id"] or s["id"],
                "description": o["description"],
                "our_debit": o["debit"],
                "our_credit": o["credit"],
                "supplier_debit": s["debit"],
                "supplier_credit": s["credit"],
            })
        else:
            mismatches += 1
            issues.append(issue("value_mismatch", o, s))
    for i in o_unmatched:
        issues.append(issue("missing_in_supplier", o=ours[i]))
    for j in s_unmatched:
        row = issue("missing_in_tahan", s=supplier[j])
        row["out_of_our_range"] = out_of_our_range(supplier[j])
        issues.append(row)

    issues.sort(key=lambda r: (r["date"], r["id"]))
    matched_rows.sort(key=lambda r: (r["date"], r["id"]))

    return {
        "summary": {
            "our_transactions": len(ours),
            "supplier_transactions": len(supplier),
            "matched": len(matched_rows),
            "value_mismatch": mismatches,
            "missing_in_supplier": len(o_unmatched),
            "missing_in_tahan": len(s_unmatched),
            "our_date_range": list(our_range) if our_range else None,
        },
        "issues": issues,
        "matched": matched_rows,
    }


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
            result = compare(data.get("ours"), data.get("supplier"))
            self._send(200, result)
        except json.JSONDecodeError:
            self._send(400, {"error": "Invalid JSON body."})
        except Exception as e:  # noqa: BLE001
            self._send(500, {"error": f"Compare failed: {e}"})
