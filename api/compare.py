"""
POST /api/compare
Body: { "ours": [rows], "supplier": [rows] }   (row schema from the extractors)

Matching, plain code, no LLM. A transaction is identified by three fields:
id (normalized), date, amount (magnitude, side-independent). Matching runs
in priority order:
  1. all three agree                          -> common (clean match)
  2. id + date agree, amount differs           -> value_mismatch
  3. id + amount agree, date differs           -> date_mismatch
  4. date + amount agree, id differs           -> id_mismatch
  5. left over on our side                     -> missing_in_supplier
  6. left over on supplier side                -> missing_in_tahan
Each step only consumes rows not already matched by an earlier, stricter
step, and each row is used at most once. Opening/closing balance rows are
excluded from matching entirely.
"""

from http.server import BaseHTTPRequestHandler
import json
import re
from datetime import datetime

TOLERANCE = 0.01


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
    return s


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def clean(rows):
    """Normalize raw extractor rows into comparison-ready records.
    'amt' is the transaction's magnitude regardless of which column
    (debit/credit) it printed in on either side - this is what makes an
    id+amount or date+amount match work even though supplier statements
    commonly mirror our debit/credit sides."""
    out = []
    for r in rows or []:
        debit = num(r.get("debit"))
        credit = num(r.get("credit"))
        amt = round(debit if debit > TOLERANCE else credit, 2)
        out.append({
            "date": norm_date(r.get("date")),
            "id": str(r.get("id") or ""),
            "nid": norm_id(r.get("id")),
            "description": str(r.get("description") or ""),
            "debit": debit,
            "credit": credit,
            "amt": amt,
            "row_type": r.get("row_type", "transaction"),
        })
    return [r for r in out if r["row_type"] == "transaction"]


def match_by_keys(ours, supplier, ours_pool, supplier_pool, keys):
    """Greedily pair remaining rows that agree on every field in `keys`
    (all must be non-empty on both sides to count). Returns matched pairs
    plus the two leftover pools."""
    sup_map = {}
    for j in supplier_pool:
        vals = tuple(supplier[j][k] for k in keys)
        if any(v in ("", None) for v in vals):
            continue
        sup_map.setdefault(vals, []).append(j)

    pairs, leftover_ours = [], []
    for i in ours_pool:
        vals = tuple(ours[i][k] for k in keys)
        bucket = sup_map.get(vals) if not any(v in ("", None) for v in vals) else None
        if bucket:
            pairs.append((i, bucket.pop(0)))
        else:
            leftover_ours.append(i)

    matched_j = {j for _, j in pairs}
    leftover_supplier = [j for j in supplier_pool if j not in matched_j]
    return pairs, leftover_ours, leftover_supplier


def row_out(kind, o=None, s=None):
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


def matched_out(o, s):
    return {
        "date": o["date"],
        "id": o["id"] or s["id"],
        "description": o["description"],
        "our_debit": o["debit"],
        "our_credit": o["credit"],
        "supplier_debit": s["debit"],
        "supplier_credit": s["credit"],
    }


def compare(ours_raw, supplier_raw):
    ours = clean(ours_raw)
    supplier = clean(supplier_raw)

    o_pool = list(range(len(ours)))
    s_pool = list(range(len(supplier)))

    exact, o_pool, s_pool = match_by_keys(ours, supplier, o_pool, s_pool, ("nid", "date", "amt"))
    value_mm, o_pool, s_pool = match_by_keys(ours, supplier, o_pool, s_pool, ("nid", "date"))
    date_mm, o_pool, s_pool = match_by_keys(ours, supplier, o_pool, s_pool, ("nid", "amt"))
    id_mm, o_pool, s_pool = match_by_keys(ours, supplier, o_pool, s_pool, ("date", "amt"))

    matched_rows = [matched_out(ours[i], supplier[j]) for i, j in exact]
    issues = []
    for i, j in value_mm:
        issues.append(row_out("value_mismatch", ours[i], supplier[j]))
    for i, j in date_mm:
        issues.append(row_out("date_mismatch", ours[i], supplier[j]))
    for i, j in id_mm:
        issues.append(row_out("id_mismatch", ours[i], supplier[j]))

    # Informational only: the date range our own file actually covers.
    # Nothing is ever hidden based on period - "out_of_our_range" just
    # flags supplier rows dated outside it, since a supplier's "to date"
    # export commonly spans months we didn't send a statement for.
    def date_range(rows):
        ds = [r["date"] for r in rows if r["date"]]
        return (min(ds), max(ds)) if ds else None

    our_range = date_range(ours)

    def out_of_our_range(r):
        if our_range is None or not r["date"]:
            return False
        return not (our_range[0] <= r["date"] <= our_range[1])

    missing_in_supplier = [row_out("missing_in_supplier", o=ours[i]) for i in o_pool]
    missing_in_tahan = []
    for j in s_pool:
        r = row_out("missing_in_tahan", s=supplier[j])
        r["out_of_our_range"] = out_of_our_range(supplier[j])
        missing_in_tahan.append(r)

    issues += missing_in_supplier + missing_in_tahan
    issues.sort(key=lambda r: (r["date"], r["id"]))
    matched_rows.sort(key=lambda r: (r["date"], r["id"]))

    return {
        "summary": {
            "our_transactions": len(ours),
            "supplier_transactions": len(supplier),
            "matched": len(matched_rows),
            "value_mismatch": len(value_mm),
            "date_mismatch": len(date_mm),
            "id_mismatch": len(id_mm),
            "missing_in_supplier": len(missing_in_supplier),
            "missing_in_tahan": len(missing_in_tahan),
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
