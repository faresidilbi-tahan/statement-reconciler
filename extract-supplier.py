"""
POST /api/extract-supplier
Body: { "pdf_base64": "<base64 of the supplier statement>" }
The supplier's format is unknown and varies per supplier, so the PDF is sent
to Claude as a document content block (preserves the visual table layout,
unlike flattened text) with an extraction prompt.

Response: { "rows": [...] }  -- same row schema as extract-ours.
Requires the ANTHROPIC_API_KEY environment variable in Vercel.
"""

from http.server import BaseHTTPRequestHandler
import json
import os

import anthropic

MODEL = "claude-haiku-4-5-20251001"
MAX_TOKENS = 16000

EXTRACTION_PROMPT = """You are extracting line items from a supplier's statement of account PDF so they can be reconciled against our own accounting records. The statement may be in English or Arabic, and its layout is unknown.

Return ONLY a JSON array. No markdown fences, no commentary, no keys outside the array.

Each statement line item becomes one object:
{"date": "YYYY-MM-DD", "id": "...", "description": "...", "debit": 0, "credit": 0, "balance": 0, "row_type": "transaction"}

Rules:
- Include every line item from every page, in document order.
- "date": convert whatever date format is printed to YYYY-MM-DD. Assume day/month/year when ambiguous. Use "" if the row has no date.
- "id": the invoice number, voucher number, or document reference printed for that row (letters and digits exactly as printed). If several references exist, prefer the invoice/voucher number. Use "" if none.
- "description": the narrative/description text as printed (Arabic text is fine as-is).
- "debit" and "credit": plain numbers, no thousands separators. Use 0 for an empty column. Copy each amount into the column it is printed in; do not reinterpret or swap sides.
- "balance": the running balance for the row if printed, else 0. If the balance carries a CR/DR or -/+ marker, output just the number.
- "row_type": "opening_balance" for opening / brought-forward rows, "closing_balance" for closing / carried-forward / final total rows, "transaction" for everything else.
- Do NOT emit rows for column headers, page headers/footers, addresses, subtotals repeated at page breaks, or blank separators.
- If two printed lines are one logical item (wrapped description), merge them into one object.

Output the JSON array and nothing else."""


def extract_rows(pdf_b64):
    client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
    message = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        temperature=0,
        messages=[{
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": pdf_b64,
                    },
                },
                {"type": "text", "text": EXTRACTION_PROMPT},
            ],
        }],
    )

    if message.stop_reason == "max_tokens":
        raise ValueError(
            "Statement is too long for one extraction pass - split the PDF and retry."
        )

    text = "".join(block.text for block in message.content if block.type == "text")
    # tolerate stray fences or preamble despite the prompt
    start, end = text.find("["), text.rfind("]")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("Claude did not return a JSON array.")
    rows = json.loads(text[start:end + 1])

    cleaned = []
    for r in rows:
        if not isinstance(r, dict):
            continue
        cleaned.append({
            "date": str(r.get("date") or ""),
            "id": str(r.get("id") or ""),
            "description": str(r.get("description") or ""),
            "debit": _num(r.get("debit")),
            "credit": _num(r.get("credit")),
            "balance": _num(r.get("balance")),
            "row_type": r.get("row_type") if r.get("row_type") in
                ("transaction", "opening_balance", "closing_balance") else "transaction",
        })
    return cleaned


def _num(v):
    try:
        return float(str(v).replace(",", "")) if v not in (None, "") else 0.0
    except ValueError:
        return 0.0


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
            if not os.environ.get("ANTHROPIC_API_KEY"):
                return self._send(500, {"error": "ANTHROPIC_API_KEY is not set in Vercel environment variables."})
            length = int(self.headers.get("Content-Length", 0))
            data = json.loads(self.rfile.read(length) or b"{}")
            b64 = data.get("pdf_base64", "")
            if "," in b64[:80]:
                b64 = b64.split(",", 1)[1]
            if not b64:
                return self._send(400, {"error": "pdf_base64 is required."})
            rows = extract_rows(b64)
            self._send(200, {"rows": rows})
        except json.JSONDecodeError:
            self._send(400, {"error": "Invalid JSON body."})
        except anthropic.APIStatusError as e:
            self._send(502, {"error": f"Claude API error {e.status_code}: {e.message}"})
        except Exception as e:  # noqa: BLE001
            self._send(500, {"error": f"Extraction failed: {e}"})
