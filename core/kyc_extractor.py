"""
kyc_extractor.py
Universal document parser — intelligently detects any document type
and extracts all fields into a clean section → attribute structure.

Handles: Aadhaar, PAN, Passport, Driving Licence, Voter ID,
         Marksheets, Experience Letters, Salary Slips,
         Appointment Letters, Bank Statements, Offer Letters,
         Company ID Cards, Birth Certificates, and any other document.
"""

from __future__ import annotations
import json, re, time, base64
from typing import Any
from openai import OpenAI

# ── Prompts ───────────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """You are an expert document intelligence system.
You can identify ANY type of document — identity cards, certificates, letters, slips, statements, marksheets — and extract every piece of information from it.
Return ONLY a valid JSON object. No markdown, no code fences, no extra text."""

USER_PROMPT = """Carefully examine this document image.

STEP 1 — Identify the document type precisely.
Examples: Aadhaar Card, PAN Card, Passport, Driving Licence, Voter ID,
          Marksheet (10th/12th/Degree/Semester), Experience Letter,
          Salary Slip, Appointment Letter, Offer Letter, Relieving Letter,
          Bank Statement, Bank Passbook, Payslip, Birth Certificate,
          Bonafide Certificate, Migration Certificate, Character Certificate,
          Company ID Card, Employee ID Card, Visa, Insurance Policy,
          Property Document, Rental Agreement, Income Tax Return, Form 16,
          or any other document type you can identify.

STEP 2 — Extract ALL information and organise into logical sections.
Each section should group related fields together with clear human-readable attribute names.

Return this exact JSON structure:
{
  "document_type": "<precise document name, e.g. 'Aadhaar Card', '12th Marksheet', 'Salary Slip - April 2024'>",
  "document_category": "<one of: identity | academic | employment | financial | legal | medical | other>",
  "issuer": "<name of the issuing organisation/authority>",
  "sections": [
    {
      "section_title": "<logical group name, e.g. 'Personal Information', 'Academic Details', 'Earnings', 'Address'>",
      "attributes": [
        {
          "name": "<clear human-readable attribute name, e.g. 'Full Name', 'Date of Birth', 'Employee ID'>",
          "value": "<exact value as printed on the document>",
          "confidence": "<high | medium | low>"
        }
      ]
    }
  ],
  "flags": {
    "has_photo": true,
    "has_signature": false,
    "has_stamp": false,
    "has_qr_code": false,
    "has_barcode": false,
    "is_masked": false,
    "masked_fields": []
  },
  "document_date": "<primary date on the document, e.g. issue date, salary month, exam date — in DD/MM/YYYY or MMM YYYY format, or null>",
  "validity": {
    "valid_from": "<DD/MM/YYYY or null>",
    "valid_until": "<DD/MM/YYYY or null>"
  },
  "raw_text": "<all visible text in reading order, line breaks as \\n>"
}

SECTION GUIDANCE by document type (adapt intelligently):

Identity Documents (Aadhaar, PAN, Passport, DL, Voter ID):
  Sections: Personal Information | Address | Document Details | Security Features

Marksheets / Academic Certificates:
  Sections: Student Information | Institution Details | Examination Details | Subject-wise Marks | Result Summary

Experience / Relieving / Recommendation Letters:
  Sections: Employee Information | Employment Details | Letter Content | Signatory Details

Salary Slips / Payslips:
  Sections: Employee Information | Pay Period | Earnings | Deductions | Net Pay Summary

Appointment / Offer Letters:
  Sections: Candidate Information | Position Details | Compensation Details | Terms & Conditions | Signatory Details

Bank Statements / Passbooks:
  Sections: Account Holder Information | Account Details | Branch Details | Statement Period | Transaction Summary

Insurance Policies:
  Sections: Policyholder Information | Policy Details | Coverage Details | Premium Details | Nominee Details

For ANY other document: use your best judgement to create 2-5 meaningful sections.

CRITICAL RULES:
1. Extract EVERY piece of visible text — miss nothing.
2. Use clear, natural English attribute names (not field codes).
3. For tables (like marks, transactions, earnings/deductions), list each row as a separate attribute.
4. Preserve exact values: numbers, dates, amounts with currency symbols.
5. If a value is masked/redacted (e.g. XXXX), still include it with confidence 'low'.
6. Never guess or infer — only extract what is visibly printed.
7. Return ONLY the JSON. Nothing else."""


# ── Extractor ─────────────────────────────────────────────────────────────────

class DocumentExtractor:
    def __init__(self, api_key: str):
        self.client = OpenAI(api_key=api_key)
        self.model  = "gpt-4o"

    def extract(
        self,
        b64_image: str,
        media_type: str = "image/jpeg",
        retries: int = 3,
        retry_delay: float = 4.0,
    ) -> dict[str, Any]:
        for attempt in range(1, retries + 1):
            try:
                response = self.client.chat.completions.create(
                    model=self.model,
                    max_tokens=3000,
                    temperature=0,
                    messages=[
                        {"role": "system", "content": SYSTEM_PROMPT},
                        {
                            "role": "user",
                            "content": [
                                {
                                    "type": "image_url",
                                    "image_url": {
                                        "url": f"data:{media_type};base64,{b64_image}",
                                        "detail": "high",
                                    },
                                },
                                {"type": "text", "text": USER_PROMPT},
                            ],
                        },
                    ],
                )
                raw = response.choices[0].message.content or ""
                return self._parse(raw)
            except Exception as exc:
                if attempt < retries:
                    time.sleep(retry_delay * attempt)
                else:
                    return self._error(str(exc))
        return self._error("Max retries exceeded")

    @staticmethod
    def _parse(raw: str) -> dict[str, Any]:
        cleaned = re.sub(r"^```(?:json)?\s*", "", raw.strip())
        cleaned = re.sub(r"\s*```$", "", cleaned)
        try:
            data = json.loads(cleaned)
            data["_parse_error"] = False
            return data
        except json.JSONDecodeError:
            return {
                "_parse_error": True,
                "_raw": raw,
                "document_type": "Unknown Document",
                "document_category": "other",
                "issuer": "",
                "sections": [],
                "flags": {},
                "document_date": None,
                "validity": {},
                "raw_text": raw,
            }

    @staticmethod
    def _error(msg: str) -> dict[str, Any]:
        return {
            "_parse_error": True,
            "_error_message": msg,
            "document_type": "Unknown Document",
            "document_category": "other",
            "issuer": "",
            "sections": [],
            "flags": {},
            "document_date": None,
            "validity": {},
            "raw_text": "",
        }


# ── Image loader ──────────────────────────────────────────────────────────────

SUPPORTED_TYPES = {
    ".jpg":  "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png":  "image/png",
    ".webp": "image/webp",
    ".gif":  "image/gif",
}

def load_image_b64(file_bytes: bytes, ext: str) -> tuple[str, str]:
    media_type = SUPPORTED_TYPES.get(ext.lower(), "image/jpeg")
    b64 = base64.b64encode(file_bytes).decode()
    return b64, media_type
