"""
POST /api/compare
Body: { "ours": [rows], "supplier": [rows] }   (row schema from the extractors)

Matching, plain code, no LLM. A transaction is identified by three fields:
id (normalized), date, amount. Amounts are compared with a tolerance (see
AMOUNT_TOLERANCE below) rather than exact equality, since suppliers commonly
round VAT/totals a cent or two differently than we do for the same invoice.
Matching runs in priority order:
  1. id + date agree, amount within tolerance    -> common (clean match)
  2. id + date agree, amount differs by more      -> value_mismatch
  3. id + amount (within tolerance) agree, date differs -> date_mismatch
  4. date + amount (within tolerance) agree, id differs -> id_mismatch
  5. left over on our side                        -> missing_in_supplier
  6. left over on supplier side                    -> missing_in_tahan
Each step only consumes rows not already matched by an earlier, stricter
step, and each row is used at most once. Opening/closing balance rows are
excluded from matching entirely.
"""

from http.server import BaseHTTPRequestHandler
import json
import re
from datetime import datetime

AMOUNT_ONLY_DATE_WINDOW_DAYS = 30  # see match_amount_within_date_window


def _parse_iso_date(s):
    try:
        y, m, d = (int(p) for p in s.split("-"))
        return datetime(y, m, d).date()
    except (ValueError, AttributeError, TypeError):
        return None


# Amounts within this many dollars/cents of each other count as "the same"
# for matching purposes. Suppliers frequently round a total a cent or two
# differently than we do for the exact same invoice (e.g. 1450.99 vs
# 1451.00) - those are not real discrepancies worth flagging.
AMOUNT_TOLERANCE = 1.00
BUILD_TAG = "2026-08-26-balance-column-net-v2"  # bump this string on every change; it is
                             # echoed back in the API response so you can
                             # confirm in the browser Network tab which
                             # build is live


def norm_id(v):
    """Normalize an id for matching. Suppliers commonly prefix their own
    reference with a document-type code ('1/249', 'INV-249', 'CN#295068'),
    so the whole string rarely matches ours verbatim even when it's really
    the same invoice number. Extract the longest run of digits - that's
    reliably the actual invoice/voucher number in every format we've seen,
    regardless of what prefix or separator surrounds it.

    One exception: a 4-digit run that looks like a calendar year (e.g. the
    "2026" inside "SAL-2026-2796") gets deprioritized, but ONLY when
    something else follows it in the string - a year embedded mid-string
    with a real number after it is almost always a batch/period marker,
    not the id itself. A year-like run at the very END with nothing after
    it is treated normally, since plenty of real invoice numbering schemes
    just happen to produce ids in the 1900-2099 range (e.g. "1/1907")."""
    s = str(v or "")
    digit_runs = re.findall(r"\d+", s)
    if not digit_runs:
        return ""

    def looks_like_year(run):
        return len(run) == 4 and run[:2] in ("19", "20")

    filtered = [r for idx, r in enumerate(digit_runs)
                if not (looks_like_year(r) and idx < len(digit_runs) - 1)]
    candidates = filtered or digit_runs
    best = max(candidates, key=len)
    return best.lstrip("0")


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
        amt = round(debit if debit > 0.01 else credit, 2)
        out.append({
            "date": norm_date(r.get("date")),
            "id": str(r.get("id") or ""),
            "nid": norm_id(r.get("id")),
            "description": str(r.get("description") or ""),
            "debit": debit,
            "credit": credit,
            "amt": amt,
            "row_type": r.get("row_type", "transaction"),
            "balance": r.get("balance"),
        })
    # A transaction with $0 on both debit and credit has no monetary
    # substance to reconcile - suppliers and our own ledger both sometimes
    # post these (a voided/reversed entry, a duplicate reference line with
    # the real amount posted separately, etc.). Drop them entirely rather
    # than let them show up as a phantom "missing" row on whichever side
    # has no matching zero-value counterpart.
    return [r for r in out if r["row_type"] == "transaction" and (r["debit"] > 0.01 or r["credit"] > 0.01)]


def amounts_close(a, b):
    return abs(a - b) <= AMOUNT_TOLERANCE + 1e-9


def match_exact_key(ours, supplier, ours_pool, supplier_pool, keys):
    """Greedily pair remaining rows that agree EXACTLY on every field in
    `keys` (all must be non-empty on both sides to count). Used only for
    key fields that need exact equality (id, date) - amount matching is
    always tolerance-based and handled separately by the tolerant matchers
    below, since a hash-map lookup can't do fuzzy numeric comparison."""
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


def match_exact_plus_tolerant_amount(ours, supplier, ours_pool, supplier_pool, exact_key):
    """Pair rows that agree exactly on `exact_key` (a single field: 'nid'
    or 'date') AND have amounts within AMOUNT_TOLERANCE of each other.
    Groups the smaller side by the exact key first so this stays fast even
    with hundreds of rows, then does a short linear scan within each group
    for the tolerant amount check."""
    sup_by_key = {}
    for j in supplier_pool:
        k = supplier[j][exact_key]
        if not k:
            continue
        sup_by_key.setdefault(k, []).append(j)

    pairs, leftover_ours = [], []
    for i in ours_pool:
        k = ours[i][exact_key]
        candidates = sup_by_key.get(k, []) if k else []
        found_idx = None
        for idx, j in enumerate(candidates):
            if amounts_close(ours[i]["amt"], supplier[j]["amt"]):
                found_idx = idx
                break
        if found_idx is not None:
            j = candidates.pop(found_idx)
            pairs.append((i, j))
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
        "our_date": o["date"] if o else None,
        "supplier_date": s["date"] if s else None,
        "our_id": o["id"] if o else None,
        "supplier_id": s["id"] if s else None,
        "our_debit": o["debit"] if o else None,
        "our_credit": o["credit"] if o else None,
        "supplier_debit": s["debit"] if s else None,
        "supplier_credit": s["credit"] if s else None,
    }


def match_amount_within_date_window(ours, supplier, ours_pool, supplier_pool, max_days):
    """Last-resort match for leftovers sharing neither id nor an exact
    date - common for cheque/receipt pairs where each side uses its own
    reference numbering and the posting date can drift a few days (cheque
    written vs. cash cleared/recorded). Pairs by amount within tolerance
    AND date within max_days.

    When several candidates on one side compete for the same amount (two
    of our entries near one supplier receipt, say), assignment happens in
    GLOBAL order of closest date match first - not just whichever ours-row
    happens to be visited first - so the genuinely nearer pairing always
    wins the shared candidate rather than an accident of list order."""
    candidates = []
    for i in ours_pool:
        o_date = _parse_iso_date(ours[i]["date"])
        if o_date is None:
            continue
        for j in supplier_pool:
            if not amounts_close(ours[i]["amt"], supplier[j]["amt"]):
                continue
            s_date = _parse_iso_date(supplier[j]["date"])
            if s_date is None:
                continue
            diff = abs((o_date - s_date).days)
            if diff <= max_days:
                candidates.append((diff, i, j))
    candidates.sort(key=lambda c: c[0])

    used_o, used_s, pairs = set(), set(), []
    for diff, i, j in candidates:
        if i in used_o or j in used_s:
            continue
        pairs.append((i, j))
        used_o.add(i)
        used_s.add(j)
    leftover_ours = [i for i in ours_pool if i not in used_o]
    leftover_supplier = [j for j in supplier_pool if j not in used_s]
    return pairs, leftover_ours, leftover_supplier


def matched_out(o, s):
    return {
        "date": o["date"],
        "id": o["id"] or s["id"],
        "our_id": o["id"],
        "supplier_id": s["id"],
        "description": o["description"],
        "our_debit": o["debit"],
        "our_credit": o["credit"],
        "supplier_debit": s["debit"],
        "supplier_credit": s["credit"],
    }


def compare(ours_raw, supplier_raw, supplier_pre_range_balance=None):
    ours = clean(ours_raw)
    supplier = clean(supplier_raw)

    o_pool = list(range(len(ours)))
    s_pool = list(range(len(supplier)))

    # Step 1: same id + same date. Split by whether the amount is within
    # tolerance (clean match) or not (a real value mismatch worth flagging).
    id_date_pairs, o_pool, s_pool = match_exact_key(ours, supplier, o_pool, s_pool, ("nid", "date"))
    exact, value_mm = [], []
    for i, j in id_date_pairs:
        if amounts_close(ours[i]["amt"], supplier[j]["amt"]):
            exact.append((i, j))
        else:
            value_mm.append((i, j))

    # Step 2: same id, amount within tolerance, date differs.
    date_mm, o_pool, s_pool = match_exact_plus_tolerant_amount(ours, supplier, o_pool, s_pool, "nid")

    # Step 3: same date, amount within tolerance, id differs.
    id_mm_candidates, o_pool, s_pool = match_exact_plus_tolerant_amount(ours, supplier, o_pool, s_pool, "date")

    # A pair only counts as a genuine id mismatch when BOTH sides actually
    # printed an id and those ids are truly different. If either side has
    # no id at all (common for cheque/payment rows some suppliers don't
    # reference by number), there's nothing conflicting to flag - date and
    # amount already agree, so it's a clean match, not a discrepancy.
    id_mm, exact_via_step3 = [], []
    for i, j in id_mm_candidates:
        if ours[i]["id"].strip() and supplier[j]["id"].strip():
            id_mm.append((i, j))
        else:
            exact_via_step3.append((i, j))

    # If literally nothing matched by id in any form (steps 1-2 both came
    # up empty) but there ARE real date+amount matches, that's a strong
    # signal this supplier's id format simply isn't compatible with ours -
    # not that 17 individual transactions all happen to disagree. Treat
    # date+amount as the real match in that case rather than flagging every
    # single row as a difference; matched_out still carries both raw id
    # strings so the discrepancy stays visible, just not as a blocker.
    id_disregarded = not exact and not value_mm and not date_mm and bool(id_mm)
    if id_disregarded:
        exact_via_step3 = exact_via_step3 + id_mm
        id_mm = []

    # Step 4: genuine leftovers - neither id nor date lined up anywhere
    # above. Suppliers commonly record a payment under their own receipt
    # number while we record it under our own cheque batch number, with
    # the posting date drifting a few days between "cheque written" and
    # "cash cleared" - amount is the only thing guaranteed to agree. Treat
    # these as clean matches (not a flagged difference) since there's
    # nothing conflicting to show, just two different bookkeeping systems
    # describing the same payment.
    amt_only_pairs, o_pool, s_pool = match_amount_within_date_window(
        ours, supplier, o_pool, s_pool, AMOUNT_ONLY_DATE_WINDOW_DAYS)

    matched_rows = [matched_out(ours[i], supplier[j]) for i, j in exact] + \
                   [matched_out(ours[i], supplier[j]) for i, j in exact_via_step3] + \
                   [matched_out(ours[i], supplier[j]) for i, j in amt_only_pairs]
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

    # Net totals for a quick "do the two books roughly agree" check. Debit
    # and credit are mirrored between the two sides (our credit = their
    # debit, by the same MIRROR_OK convention used throughout matching),
    # so "net" here means the same thing on both sides: the net amount
    # outstanding on this account per that side's own books.
    #
    # The supplier total only counts rows within OUR file's date range -
    # otherwise a supplier statement spanning more months than ours would
    # show a large "difference" that's really just missing prior-period
    # data on our side, not a real discrepancy (see id_disregarded above
    # for the same kind of period-mismatch problem in matching).
    #
    # EXCEPTION: a supplier row that was actually matched to one of our
    # transactions is always counted, regardless of its date. Matching
    # already proves it's genuinely relevant to this reconciliation - a
    # real-world example that surfaced this: a lump cheque payment posted
    # by us on the last day of our statement, with the supplier recording
    # the matching credit a day or two later (past our file's own date
    # range). Excluding a verified match purely because of a one-day
    # boundary created a phantom "unexplained" gap equal to the entire
    # payment, even though the transaction was correctly reconciled.
    our_total_debit = round(sum(r["debit"] for r in ours), 2)
    our_total_credit = round(sum(r["credit"] for r in ours), 2)
    our_net_debit_credit = round(our_total_credit - our_total_debit, 2)

    matched_supplier_indices = (
        set(j for _, j in exact) | set(j for _, j in exact_via_step3) | set(j for _, j in amt_only_pairs)
    )
    supplier_in_range = [
        supplier[j] for j in range(len(supplier))
        if j in matched_supplier_indices or not out_of_our_range(supplier[j])
    ]
    supplier_total_debit = round(sum(r["debit"] for r in supplier_in_range), 2)
    supplier_total_credit = round(sum(r["credit"] for r in supplier_in_range), 2)
    supplier_net_debit_credit = round(supplier_total_debit - supplier_total_credit, 2)

    # Net, cross-checked directly against each file's own printed Balance
    # column rather than trusting only our own debit/credit summation.
    # This catches extraction errors that the debit/credit totals alone
    # wouldn't reveal (e.g. two mistakes that happen to cancel out).
    #
    # Sign conventions differ between the two sides and are derived from
    # real extracted data, not assumed: "ours" (accsta04) writes credit-
    # side balances as negative (parenthesized), so opening minus closing
    # gives the same positive-when-credit-heavy convention as our_net.
    # The supplier's balance column runs the other way - it increases
    # with debit and decreases with credit - so end minus start already
    # matches supplier_net's positive-when-debit-heavy convention.
    #
    # The balance-derived figure is only trusted (and only replaces the
    # debit/credit total) when it's available AND agrees with the
    # debit/credit total within $1 - if a supplier's format encodes
    # balance differently than assumed here, the two would diverge
    # sharply, and blindly trusting the balance column then would be
    # worse than falling back to debit/credit. Either way this is
    # surfaced to the caller via *_net_source, never silently guessed.
    NET_CROSSCHECK_TOLERANCE = 1.0

    our_net = our_net_debit_credit
    our_net_source = "debit_credit"
    our_opening_rows = [r for r in ours_raw if r.get("row_type") == "opening_balance" and r.get("balance") is not None]
    our_closing_rows = [r for r in ours_raw if r.get("row_type") == "closing_balance" and r.get("balance") is not None]
    if len(our_opening_rows) == 1 and len(our_closing_rows) == 1:
        our_net_balance = round(our_opening_rows[0]["balance"] - our_closing_rows[0]["balance"], 2)
        if abs(our_net_balance - our_net_debit_credit) <= NET_CROSSCHECK_TOLERANCE:
            our_net = our_net_balance
            our_net_source = "balance"

    supplier_net = supplier_net_debit_credit
    supplier_net_source = "debit_credit"
    supplier_in_range_dated = [r for r in supplier_in_range if r.get("date") and r.get("balance") is not None]
    if supplier_pre_range_balance is not None and supplier_in_range_dated:
        # Multiple transactions can share the same date - the true
        # end-of-range balance is whichever one was posted LAST that day,
        # not just "any row with the maximum date". Rows are already in
        # the source document's own sequential order, so among rows tied
        # on the max date, the last one in that order is the real
        # end-of-day balance.
        max_date = max(r["date"] for r in supplier_in_range_dated)
        same_day = [r for r in supplier_in_range_dated if r["date"] == max_date]
        end_row = same_day[-1]
        supplier_net_balance = round(end_row["balance"] - supplier_pre_range_balance, 2)
        if abs(supplier_net_balance - supplier_net_debit_credit) <= NET_CROSSCHECK_TOLERANCE:
            supplier_net = supplier_net_balance
            supplier_net_source = "balance"

    net_difference = round(our_net - supplier_net, 2)

    return {
        "summary": {
            "build_tag": BUILD_TAG,
            "our_transactions": len(ours),
            "supplier_transactions": len(supplier),
            "matched": len(matched_rows),
            "value_mismatch": len(value_mm),
            "date_mismatch": len(date_mm),
            "id_mismatch": len(id_mm),
            "missing_in_supplier": len(missing_in_supplier),
            "missing_in_tahan": len(missing_in_tahan),
            "our_date_range": list(our_range) if our_range else None,
            "id_disregarded": id_disregarded,
            "our_total_debit": our_total_debit,
            "our_total_credit": our_total_credit,
            "our_net": our_net,
            "our_net_source": our_net_source,
            "our_net_debit_credit": our_net_debit_credit,
            "supplier_total_debit": supplier_total_debit,
            "supplier_total_credit": supplier_total_credit,
            "supplier_net": supplier_net,
            "supplier_net_source": supplier_net_source,
            "supplier_net_debit_credit": supplier_net_debit_credit,
            "net_difference": net_difference,
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
            result = compare(data.get("ours"), data.get("supplier"), data.get("supplier_pre_range_balance"))
            self._send(200, result)
        except json.JSONDecodeError:
            self._send(400, {"error": "Invalid JSON body."})
        except Exception as e:  # noqa: BLE001
            self._send(500, {"error": f"Compare failed: {e}"})
