# GST ITC Matcher

Match Purchase Register invoices with GSTR-2A/2B and generate an ITC Taken report with match status and eligible credit totals.

## Features

- Auto-detects columns from Tally, Busy, SAP, and GST portal Excel exports
- Match statuses: Fully Matched, Tax Mismatch, GSTIN Mismatch, Inv No Mismatch, Inv Date Mismatch, Not Matched, Duplicate Invoice
- Calculates eligible ITC (IGST, CGST, SGST)
- **Vendor email reminders** — one-click SMTP reminders to non-compliant suppliers; Day 3 & Day 7 follow-ups
- Exports formatted Excel with Summary, ITC Claim Plan, Action Plan, and Vendor Follow-up sheets
- Web UI via Streamlit

## Quick Start

```bash
cd gst-itc-matcher
pip install -r requirements.txt
streamlit run app.py
```

Or double-click **`start.bat`** on Windows.

## Vendor email reminders (recover blocked ITC)

After matching, the **Vendor ITC Recovery** section lists suppliers blocking ITC (Not Matched, GSTIN/invoice/tax mismatches).

1. Optionally upload a **vendor contacts** CSV/Excel with `GSTIN` and `Email` columns
2. Enter vendor emails in the table (or load from contacts file)
3. Click **Send reminders to all vendors (1-click)** — initial email goes out; follow-ups auto-schedule for Day 3 and Day 7
4. On later visits, click **Send due follow-ups** for automated escalation emails

### SMTP setup (Streamlit Cloud)

In your app **Settings → Secrets**, add:

```toml
[email]
smtp_host = "smtp.gmail.com"
smtp_port = 587
smtp_user = "your-email@gmail.com"
smtp_password = "your-app-password"
from_email = "your-email@gmail.com"
from_name = "Your Company Accounts"
company_name = "Your Company Pvt Ltd"
```

For Gmail, use an [App Password](https://support.google.com/accounts/answer/185833). Without SMTP, use **Open in email** mailto links instead.

See `.streamlit/secrets.toml.example` for a full template.

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
