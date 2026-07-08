# GST ITC Matcher

Match Purchase Register invoices with GSTR-2A/2B and generate an ITC Taken report with match status and eligible credit totals.

## Features

- Auto-detects columns from Tally, Busy, SAP, and GST portal Excel exports
- Match statuses: Fully Matched, Tax Mismatch, GSTIN Mismatch, Inv No Mismatch, Inv Date Mismatch, Not Matched, Duplicate Invoice
- Calculates eligible ITC (IGST, CGST, SGST)
- Exports formatted Excel with Summary sheet
- Web UI via Streamlit

## Quick Start

```bash
cd gst-itc-matcher
pip install -r requirements.txt
streamlit run app.py
```

Or double-click **`start.bat`** on Windows.

### Usage

**Option A — Single Purchase Register**
1. Select **Single file**
2. Upload Purchase Register + GSTR-2A/2B

**Option B — Sales + Service (Consolidated)**
1. Select **Sales + Service (Consolidate)**
2. Upload Sales Purchase Register
3. Upload Service Purchase Register
4. Upload GSTR-2A/2B

The app automatically merges Sales + Service into one consolidated Purchase Register, then matches with GSTR-2A/2B. Download both:
- **Consolidated Purchase Register**
- **ITC Taken Excel**

## Match Logic

| Field | Comparison |
|-------|------------|
| Supplier GSTIN | Primary match key |
| Invoice Number | Primary match key |
| Invoice Date | Secondary validation |
| Tax Amounts | IGST, CGST, SGST (+ taxable value) |

## ITC Decision Rules

| Status | ITC Decision |
|--------|--------------|
| Fully Matched | Take full ITC |
| Tax Mismatch | Take lower of PR vs GSTR tax (verify first) |
| Inv Date Mismatch | Hold — verify with supplier |
| GSTIN / Inv No Mismatch | Do not take ITC |
| Not Matched | Do not take — follow up with supplier |
| Duplicate Invoice | Do not take — remove duplicate |

## Output Columns

Supplier GSTIN, Supplier Name, Invoice No., Invoice Date, Taxable Value, IGST, CGST, SGST, Purchase Register, GSTR-2A/2B, Match Status, ITC Taken, Remarks

## CLI Usage

```python
from matcher.engine import load_and_match, export_to_excel

result, summary = load_and_match("purchase.xlsx", "gstr2b.xlsx")
with open("report.xlsx", "wb") as f:
    f.write(export_to_excel(result, summary))
```

## Disclaimer

ITC decisions are indicative. Verify eligibility under GST rules before claiming credit.
