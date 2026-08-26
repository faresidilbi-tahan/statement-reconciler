"""
POST /api/ai-totals
Runs automatically after every Compare that has totals. Re-sums the
Debit/Credit totals for both sides using AI instead of (well, alongside)
the deterministic arithmetic in compare.py.

Deliberately cheap: no PDFs are sent (the rows were already extracted by
the deterministic parsers), uses the cheapest available model, and caps
the response length. Typical cost per call is a small fraction of a cent
- there is no literal per-request spend cap in the Anthropic API, so this
keeps cost low structurally (small text-only input, capped output)
rather than by any enforced ceiling.

The frontend keeps the exact deterministic totals as a silent background
check: if this endpoint's numbers ever drift more than $1 from the real
sum, the UI falls back to showing the exact figures instead of AI's -
this endpoint's output is a display convenience, not the source of truth.

Requires an "accounting" environment variable holding a valid Anthropic
API key. If missing or the call fails, the frontend falls back to the
exact deterministic totals with a quiet note.
"""

from http.server import BaseHTTPRequestHandler
import json
import os
import re
import urllib.request
import urllib.error

BUILD_TAG = "2026-08-26-ai-totals-v1"

ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages"
MODEL = "claude-haiku-4-5-20251001"
API_KEY_ENV_VAR = "accounting"
MAX_TOKENS = 200

PROMPT = """Sum two lists of transactions. Each row has a debit and credit amount (numbers, possibly 0).

OUR_ROWS:
{our_rows_json}

SUPPLIER_ROWS:
{supplier_rows_json}

For OUR_ROWS: add up every "debit" value into our_total_debit, and every "credit" value into our_total_credit.
For SUPPLIER_ROWS: add up every "debit" value into supplier_total_debit, and every "credit" value into supplier_total_credit.
Round every total to 2 decimal places.

Respond with ONLY a single JSON object, nothing else, no markdown fences:
{{"our_total_debit": <number>, "our_total_credit": <number>, "supplier_total_debit": <number>, "supplier_total_credit": <number>}}
"""


def call_claude(prompt_text):
    payload = {
        "model": MODEL,
        "max_tokens": MAX_TOKENS,
        "messages": [{"role": "user", "content": prompt_text}],
    }
    req = urllib.request.Request(
        ANTHROPIC_API_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "x-api-key": os.environ.get(API_KEY_ENV_VAR, ""),
            "anthropic-version": "2023-06-01",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        data = json.loads(resp.read().decode("utf-8"))

    text_parts = [b.get("text", "") for b in data.get("content", []) if b.get("type") == "text"]
    raw = "".join(text_parts).strip()
    if not raw:
        stop_reason = data.get("stop_reason", "unknown")
        raise ValueError("Claude returned no text content (stop_reason={}).".format(stop_reason))

    raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw.strip())
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", raw, re.S)
        if m:
            return json.loads(m.group(0))
        raise


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

            if not os.environ.get(API_KEY_ENV_VAR):
                return self._send(400, {
                    "error": "\"{}\" is not set in this Vercel project's environment "
                             "variables (or has no value).".format(API_KEY_ENV_VAR)
                })

            our_rows = data.get("our_rows", []) or []
            supplier_rows = data.get("supplier_rows", []) or []

            prompt = PROMPT.format(
                our_rows_json=json.dumps(our_rows, ensure_ascii=False),
                supplier_rows_json=json.dumps(supplier_rows, ensure_ascii=False),
            )
            result = call_claude(prompt)
            result["build_tag"] = BUILD_TAG
            return self._send(200, result)

        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", errors="replace")
            return self._send(502, {"error": "Anthropic API error ({}): {}".format(e.code, body[:300])})
        except json.JSONDecodeError as e:
            return self._send(502, {"error": "AI response was not valid JSON: {}".format(e)})
        except Exception as e:  # noqa: BLE001
            return self._send(500, {"error": "AI totals failed: {}".format(e)})
