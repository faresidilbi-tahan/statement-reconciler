"""
POST /api/ai-verify
Optional, manually-triggered AI double-check layer. This is entirely
separate from extraction and matching - it never runs on its own, and if
it's unavailable or fails, extraction/compare continue to work exactly as
before. Reintroduced explicitly at the person's request (2026-08-21) after
the earlier AI-extraction fallback was reverted; unlike that fallback, this
does NOT parse or replace anything - it only checks the deterministic
parser's output against the raw PDF, after the fact, only when asked.

Two modes, chosen by "mode" in the request body:

  mode="totals" - reads the PDF directly with Claude and asks it to work
    out the statement's true total debit/credit (preferring the
    document's own printed grand-total line when one exists, since that's
    more reliable than an LLM re-summing a long table). Compares that
    against the total the deterministic extractor already computed for
    the same file. Returns whether they agree.

  mode="rows" - only meant to be called after "totals" found a mismatch.
    Sends the PDF plus the already-extracted row table and asks Claude to
    flag specific rows that look wrong or missing, rather than re-parsing
    the whole document blind.

Requires ANTHROPIC_API_KEY in this Vercel project's environment variables
(Project Settings -> Environment Variables -> redeploy). No extra pip
dependency - talks to the Anthropic API directly over HTTPS.
"""

from http.server import BaseHTTPRequestHandler
import base64
import json
import os
import re
import urllib.request
import urllib.error

BUILD_TAG = "2026-08-21-ai-verify-v1"

ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages"
MODEL = "claude-sonnet-5"
AMOUNT_TOLERANCE = 1.0

TOTALS_PROMPT = """You are auditing a financial statement PDF for a bookkeeping reconciliation tool.

Read the ENTIRE document carefully, including every page.

Your task: determine the correct TOTAL debit and TOTAL credit across every real transaction line in this statement. Do not count an opening-balance line's own debit/credit as a transaction, but do include everything else.

If the document itself prints a grand total or running closing-balance line that already sums debit and credit across the whole statement (for example a line like "Balance as at ...", "Total", or "Closing Balance" near the end), prefer that printed total over summing manually - it is more reliable than your own arithmetic on a long table. Only sum the rows yourself if no such line exists.

Respond with ONLY a single JSON object, nothing else, no markdown fences, no explanation before or after:
{"total_debit": <number>, "total_credit": <number>, "source": "printed_total" or "manual_sum", "notes": "<one short sentence - e.g. which line you used, or any page you had trouble reading>"}
"""

ROWS_PROMPT = """You are auditing a financial statement PDF against a table of transactions that was already extracted from it by another program, which may contain errors.

Extracted table (JSON array, one object per row - fields: date, id, description, debit, credit):

{rows_json}

Read the actual PDF and compare it row by row against this extracted table. Only flag genuine discrepancies:
- A row where the extracted debit or credit amount does not match what the PDF actually shows for that transaction (ignore differences smaller than {tolerance})
- A real transaction row that exists in the PDF but is completely missing from the extracted table
- A row in the extracted table that does not correspond to any real transaction in the PDF (e.g. a spurious or duplicated row)

Do NOT flag minor description wording differences, formatting differences, or opening/closing balance summary rows. Be conservative - only report something you are genuinely confident about, since this feeds a report a person will read and act on.

Respond with ONLY a single JSON object, nothing else, no markdown fences:
{{"suspect_rows": [{{"date": "...", "id": "...", "extracted_debit": <number>, "extracted_credit": <number>, "correct_debit": <number>, "correct_credit": <number>, "note": "<short reason>"}}], "missing_rows": [{{"date": "...", "id": "...", "debit": <number>, "credit": <number>, "description": "...", "note": "present in PDF but not in extracted table"}}], "spurious_rows": [{{"date": "...", "id": "...", "note": "in extracted table but not found in PDF"}}]}}

If there are no issues at all in a category, return an empty array for it.
"""


def call_claude(pdf_bytes, prompt_text):
    payload = {
        "model": MODEL,
        "max_tokens": 4096,
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": base64.b64encode(pdf_bytes).decode("ascii"),
                        },
                    },
                    {"type": "text", "text": prompt_text},
                ],
            }
        ],
    }
    req = urllib.request.Request(
        ANTHROPIC_API_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "x-api-key": os.environ.get("ANTHROPIC_API_KEY", ""),
            "anthropic-version": "2023-06-01",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=55) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    text_parts = [b["text"] for b in data.get("content", []) if b.get("type") == "text"]
    raw = "".join(text_parts).strip()
    # Strip markdown fences in case the model adds them despite instructions.
    raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw.strip())
    return json.loads(raw)


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

            if not os.environ.get("ANTHROPIC_API_KEY"):
                return self._send(400, {
                    "error": "ANTHROPIC_API_KEY is not set in this Vercel project's "
                             "environment variables. Add it under Project Settings -> "
                             "Environment Variables, then redeploy."
                })

            b64 = data.get("pdf_base64", "")
            if "," in b64[:80]:
                b64 = b64.split(",", 1)[1]
            pdf_bytes = base64.b64decode(b64)
            if not pdf_bytes.startswith(b"%PDF"):
                return self._send(400, {"error": "AI Verify currently only supports PDF files."})

            mode = data.get("mode", "totals")

            if mode == "totals":
                expected_debit = float(data.get("expected_total_debit", 0) or 0)
                expected_credit = float(data.get("expected_total_credit", 0) or 0)
                result = call_claude(pdf_bytes, TOTALS_PROMPT)
                ai_debit = float(result.get("total_debit", 0) or 0)
                ai_credit = float(result.get("total_credit", 0) or 0)
                matches = (
                    abs(ai_debit - expected_debit) <= AMOUNT_TOLERANCE
                    and abs(ai_credit - expected_credit) <= AMOUNT_TOLERANCE
                )
                return self._send(200, {
                    "ai_total_debit": ai_debit,
                    "ai_total_credit": ai_credit,
                    "expected_total_debit": expected_debit,
                    "expected_total_credit": expected_credit,
                    "matches_extractor": matches,
                    "source": result.get("source"),
                    "notes": result.get("notes", ""),
                    "build_tag": BUILD_TAG,
                })

            elif mode == "rows":
                rows = data.get("rows", []) or []
                slim_rows = [
                    {
                        "date": r.get("date"),
                        "id": r.get("id"),
                        "description": (r.get("description") or "")[:120],
                        "debit": r.get("debit"),
                        "credit": r.get("credit"),
                    }
                    for r in rows
                ]
                prompt = ROWS_PROMPT.format(
                    rows_json=json.dumps(slim_rows, ensure_ascii=False),
                    tolerance="${:.2f}".format(AMOUNT_TOLERANCE),
                )
                result = call_claude(pdf_bytes, prompt)
                result["build_tag"] = BUILD_TAG
                return self._send(200, result)

            else:
                return self._send(400, {"error": "Unknown mode: {}".format(mode)})

        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace")
            return self._send(502, {"error": "Anthropic API error ({}): {}".format(e.code, body[:300])})
        except json.JSONDecodeError as e:
            return self._send(502, {"error": "AI response was not valid JSON: {}".format(e)})
        except Exception as e:  # noqa: BLE001
            return self._send(500, {"error": "AI Verify failed: {}".format(e)})
