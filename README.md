# KYC Document Annotator
**Mobius Knowledge Services**

A Streamlit application that uses Claude vision AI to extract and annotate all fields from KYC / identity documents with Named Entity Recognition (NER) labels.

---

## Supported Document Types

| Document | Fields Extracted |
|---|---|
| Aadhaar Card | Name, DOB, Gender, Address, UID (masked support) |
| PAN Card | Name, Father Name, DOB, PAN Number |
| Passport | Name, DOB, Nationality, Passport No., Expiry |
| Driving Licence | Name, DOB, DL No., Vehicle Class, Expiry |
| Voter ID | Name, Father/Husband Name, Age, EPIC No., Address |
| Company ID Card | Name, Employee ID, Designation, Department, Organisation |
| Generic ID | Best-effort extraction of all visible fields |

---

## Entity Types (NER Labels)

| Category | Entity Types |
|---|---|
| Identity | PERSON_NAME, FATHER_NAME, MOTHER_NAME, SPOUSE_NAME, GENDER, DATE_OF_BIRTH, BLOOD_GROUP |
| Location | ADDRESS_LINE, CITY, STATE, PINCODE, COUNTRY, PLACE_OF_BIRTH |
| Document | ID_NUMBER, ID_NUMBER_SECONDARY, ISSUE_DATE, EXPIRY_DATE, DOCUMENT_TITLE |
| Organisation | ORGANISATION_NAME, ISSUER_NAME, DEPARTMENT, DESIGNATION, EMPLOYEE_ID |
| Contact | PHONE_NUMBER, EMAIL |
| Licence | VEHICLE_CLASS, ENDORSEMENTS |
| Other | NATIONALITY, SIGNATURE, PHOTO, QR_CODE, BARCODE |

---

## Setup

### 1. Install dependencies
```bash
pip install -r requirements.txt
```

### 2. Configure API key
```bash
cp .streamlit/secrets_template.toml .streamlit/secrets.toml
# Edit secrets.toml and add your Anthropic API key
```

### 3. Run
```bash
streamlit run app.py
```

---

## Architecture

```
app.py                    # Streamlit UI
core/
  kyc_extractor.py        # Claude vision API integration + NER prompt
  excel_builder.py        # Colour-coded Excel export
  __init__.py
```

### Key differences from Financial Extractor
- Uses **Anthropic Claude** (claude-sonnet-4-20250514) instead of OpenAI GPT-4o
- Accepts **image files** (JPG, PNG, WEBP) instead of PDFs
- Batch upload of **multiple documents** in one session
- **NER annotation layer** on top of field extraction
- **Confidence scoring** per field (high / medium / low)
- **Document flags** (has_photo, has_signature, is_masked, masked_fields)
- **Grouped annotation view** in UI by entity category

---

## Output Excel Structure

- **Sheet 1 — Summary**: One row per document with key fields
- **Sheet N — Document**: Full annotation table with entity colours, confidence badges, and raw text block

---

## Privacy Notes

- Documents are processed in memory; no data is stored on disk beyond the session
- API calls go to Anthropic's claude-sonnet-4-20250514 via their API
- Do not use on production PII without appropriate data processing agreements
