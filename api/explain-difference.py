"""
POST /api/explain-difference
Runs automatically after every Compare that has at least one issue.
Powers the "Why the difference?" panel: takes the EXACT, already-computed
dollar breakdown (never recalculated here - the numbers shown on screen
always come from the deterministic JS/compare.py math) plus the
underlying issue rows, and asks Claude to:

  1. Independently sanity-check that the breakdown actually sums to the
     real net difference (a verification pass, not a recalculation the
     app trusts over its own arithmetic).
  2. Write a short plain-English narrative explaining what's likely going
     on, using the row descriptions - something arithmetic alone can't
     do (e.g. noticing several missing rows cluster around one cash-
     payment date, or a batch of invoices from the same week).
  3. Optionally surface 2-4 short thematic groupings among the issue rows
     if a real pattern exists, rather than restating the category counts.

Requires an "accounting" environment variable holding a valid
Anthropic API key. If missing or the call fails, the frontend falls back
to a plain auto-generated sentence - this endpoint failing never breaks
the rest of the dashboard.
"""

from http.server import BaseHTTPRequestHandler
import json
import os
import re
import urllib.request
import urllib.error

BUILD_TAG = "2026-08-26-explain-difference-v2-accounting"

ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages"
MODEL = "claude-sonnet-5"
API_KEY_ENV_VAR = "accounting"

PROMPT = """You're helping someone reconcile two versions of the same supplier account: their own company's ledger ("ours") against the supplier's own statement. Below is the EXACT dollar breakdown of the difference between the two files, already computed deterministically - do not recalculate or alter these numbers, only sanity-check them.

SUMMARY:
{summary_json}

BREAKDOWN (already computed - category, row count, dollar contribution to the difference):
{breakdown_json}

UNEXPLAINED RESIDUAL (breakdown total minus this equals the real difference): {residual}

UNDERLYING ROWS (the actual transactions behind each breakdown category - issue type, date, id, description, and each side's debit/credit):
{issues_json}

Your tasks:
1. VERIFY: Add up the breakdown amounts plus the residual. Does it equal net_difference (summary.net_difference) within $0.01? Set "verified" accordingly - this is a check, not a recalculation the app will trust over its own numbers.
2. NARRATIVE: Write 2-4 sentences in plain business English explaining what's likely driving this difference, referencing real patterns in the row descriptions where they're informative (e.g. timing of postings, cash-payment batching, invoice numbering gaps, credit notes, discounts). Avoid vague filler like "there are some differences" - be specific about what you actually see in the data. If nothing beyond the raw category breakdown is evident, say so plainly rather than padding.
3. THEMES (optional): If there's a real pattern among the rows beyond what the category breakdown already shows - e.g. several missing rows all dated the same week, or several value mismatches sharing a consistent dollar gap - list up to 4 short themes, each with a title, how many rows it covers, and one sentence of insight. Return an empty array if there's no meaningful pattern beyond the categories already shown - don't invent themes to fill the list.

Respond with ONLY a single JSON object, nothing else, no markdown fences:
{{"verified": true or false, "verification_note": "<short note if not verified, else empty string>", "narrative": "<2-4 sentences>", "themes": [{{"title": "...", "count": <number>, "note": "<one sentence>"}}]}}
"""


def call_claude(prompt_text, max_tokens=900):
    payload = {
        "model": MODEL,
        "max_tokens": max_tokens,
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
    with urllib.request.urlopen(req, timeout=45) as resp:
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

            summary = data.get("summary", {})
            breakdown = data.get("breakdown", [])
            residual = data.get("residual", 0)
            issues = data.get("issues", [])[:80]

            prompt = PROMPT.format(
                summary_json=json.dumps(summary, ensure_ascii=False),
                breakdown_json=json.dumps(breakdown, ensure_ascii=False),
                residual=residual,
                issues_json=json.dumps(issues, ensure_ascii=False),
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
            return self._send(500, {"error": "Explain-difference failed: {}".format(e)})
