"""Vendor email reminders for blocked ITC recovery and automated follow-ups."""

from __future__ import annotations

import smtplib
import ssl
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO
from typing import Any
from urllib.parse import quote

import pandas as pd

from .itc_dashboard import ITC_NON_ELIGIBLE, MATCH_GSTIN, MATCH_INV_NO, MATCH_NOT, MATCH_TAX

VENDOR_FOLLOWUP_STATUSES = {MATCH_NOT, MATCH_GSTIN, MATCH_INV_NO, MATCH_TAX}

FOLLOWUP_INTERVALS_DAYS = (3, 7)


@dataclass
class SmtpConfig:
    host: str
    port: int
    user: str
    password: str
    from_email: str
    from_name: str


@dataclass
class VendorReminder:
    supplier_gstin: str
    supplier_name: str
    email: str
    blocked_itc: float
    invoice_count: int
    invoice_list: str
    primary_issue: str
    issue_summary: str


@dataclass
class FollowUpLogEntry:
    supplier_gstin: str
    supplier_name: str
    email: str
    reminder_number: int
    sent_at: datetime
    blocked_itc: float
    next_follow_up_date: date | None
    status: str
    subject: str


@dataclass
class SendResult:
    supplier_name: str
    email: str
    success: bool
    message: str
    log_entry: FollowUpLogEntry | None = None


def load_smtp_config(secrets: dict[str, Any] | None = None) -> SmtpConfig | None:
    """Load SMTP settings from Streamlit secrets or environment-style dict."""
    email_cfg = (secrets or {}).get("email") or {}
    required = ("smtp_host", "smtp_user", "smtp_password", "from_email")
    if not all(email_cfg.get(key) for key in required):
        return None
    return SmtpConfig(
        host=str(email_cfg["smtp_host"]),
        port=int(email_cfg.get("smtp_port", 587)),
        user=str(email_cfg["smtp_user"]),
        password=str(email_cfg["smtp_password"]),
        from_email=str(email_cfg["from_email"]),
        from_name=str(email_cfg.get("from_name", email_cfg.get("company_name", "Accounts Team"))),
    )


def _round(value: float) -> float:
    return round(float(value), 2)


def _issue_label(status: str) -> str:
    labels = {
        MATCH_NOT: "Invoice not filed in GSTR-1 / GSTR-2B",
        MATCH_GSTIN: "GSTIN mismatch in GSTR-2B",
        MATCH_INV_NO: "Invoice number mismatch in GSTR-2B",
        MATCH_TAX: "Tax amount mismatch in GSTR-2B",
    }
    return labels.get(status, status)


def _primary_issue(statuses: list[str]) -> str:
    priority = [MATCH_NOT, MATCH_GSTIN, MATCH_INV_NO, MATCH_TAX]
    for status in priority:
        if status in statuses:
            return status
    return statuses[0] if statuses else MATCH_NOT


def extract_vendor_reminders(
    enriched: pd.DataFrame,
    contacts: pd.DataFrame | None = None,
) -> list[VendorReminder]:
    """Group blocked-ITC rows by supplier for vendor follow-up."""
    blocked = enriched[
        (enriched["ITC Category"] == ITC_NON_ELIGIBLE)
        & (enriched["Match Status"].isin(VENDOR_FOLLOWUP_STATUSES))
    ].copy()
    if blocked.empty:
        return []

    contact_map: dict[str, str] = {}
    if contacts is not None and not contacts.empty:
        cols = {c.lower().strip(): c for c in contacts.columns}
        gstin_col = cols.get("gstin") or cols.get("supplier gstin")
        email_col = cols.get("email") or cols.get("vendor email")
        if gstin_col and email_col:
            for _, row in contacts.iterrows():
                gstin = str(row.get(gstin_col, "")).strip().upper()
                email = str(row.get(email_col, "")).strip()
                if gstin and email and email.lower() != "nan":
                    contact_map[gstin] = email

    reminders: list[VendorReminder] = []
    grouped = blocked.groupby(["Supplier GSTIN", "Supplier Name"], dropna=False)

    for (gstin, name), group in grouped:
        gstin_str = str(gstin or "").strip()
        name_str = str(name or "Supplier").strip()
        statuses = group["Match Status"].tolist()
        primary = _primary_issue(statuses)
        blocked_itc = _round(group["Non-Eligible ITC"].sum())
        invoices = group["Invoice No."].astype(str).tolist()
        invoice_list = ", ".join(invoices[:8])
        if len(invoices) > 8:
            invoice_list += f" (+{len(invoices) - 8} more)"

        issue_counts = group["Match Status"].value_counts()
        issue_summary = "; ".join(
            f"{count} × {_issue_label(status)}" for status, count in issue_counts.items()
        )

        email = contact_map.get(gstin_str.upper(), "")
        reminders.append(
            VendorReminder(
                supplier_gstin=gstin_str,
                supplier_name=name_str,
                email=email,
                blocked_itc=blocked_itc,
                invoice_count=len(group),
                invoice_list=invoice_list,
                primary_issue=primary,
                issue_summary=issue_summary,
            )
        )

    reminders.sort(key=lambda item: item.blocked_itc, reverse=True)
    return reminders


def vendor_reminders_dataframe(reminders: list[VendorReminder]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "Select": True,
                "Supplier Name": item.supplier_name,
                "Supplier GSTIN": item.supplier_gstin,
                "Vendor Email": item.email,
                "Blocked ITC (₹)": item.blocked_itc,
                "Invoices": item.invoice_count,
                "Invoice Nos.": item.invoice_list,
                "Issue": _issue_label(item.primary_issue),
                "Details": item.issue_summary,
            }
            for item in reminders
        ]
    )


def build_reminder_email(
    vendor: VendorReminder,
    reminder_number: int,
    company_name: str,
    sender_name: str,
    return_period: str = "",
) -> tuple[str, str]:
    """Return (subject, plain-text body) for initial or follow-up reminder."""
    period_note = f" for return period {return_period}" if return_period else ""
    amount = vendor.blocked_itc
    invoices = vendor.invoice_list
    gstin = vendor.supplier_gstin or "your GSTIN"

    if reminder_number <= 1:
        subject = _short_subject(vendor, reminder_number, return_period)
        body = f"""Dear {vendor.supplier_name},

We are reconciling our Input Tax Credit (ITC) with GSTR-2B{period_note}.

The following invoice(s) from your firm are blocking ITC of ₹{amount:,.2f} in our books:
  • Invoice(s): {invoices}
  • GSTIN: {gstin}
  • Issue: {_issue_label(vendor.primary_issue)}
  • Details: {vendor.issue_summary}

Kindly file or amend your GSTR-1 at the earliest so these invoices reflect correctly in our GSTR-2B. Until then, we cannot claim this ITC under GST rules.

Please confirm once updated, or share the correct invoice / credit note details if any correction is needed on our side.

Regards,
{sender_name}
{company_name}
"""
    elif reminder_number == 2:
        subject = _short_subject(vendor, reminder_number, return_period)
        body = f"""Dear {vendor.supplier_name},

This is a follow-up to our earlier email regarding GST invoice(s): {invoices}.

₹{amount:,.2f} of Input Tax Credit remains blocked because the invoice(s) are still not matching in GSTR-2B ({_issue_label(vendor.primary_issue)}).

Please treat this as urgent and update GSTR-1 / amend the return within 2 business days. Delay affects our GST compliance and your payment relationship with us.

Reply with the ARN or filing confirmation once done.

Regards,
{sender_name}
{company_name}
"""
    else:
        subject = _short_subject(vendor, reminder_number, return_period)
        body = f"""Dear {vendor.supplier_name},

Despite earlier reminders, invoice(s) {invoices} (GSTIN {gstin}) are still not compliant in GSTR-2B.

Blocked ITC: ₹{amount:,.2f}
Issue: {vendor.issue_summary}

This is our final reminder before we escalate to our purchase / finance team. Please file or amend GSTR-1 immediately and share confirmation.

If not resolved within 48 hours, we may withhold pending payments until GST records are corrected.

Regards,
{sender_name}
{company_name}
"""
    return subject, body.strip()


def _ascii_for_email_url(text: str) -> str:
    """Use ASCII in URL params — special chars break mailto/Gmail compose links."""
    return (
        text.replace("₹", "Rs.")
        .replace("—", "-")
        .replace("–", "-")
        .replace("•", "-")
        .replace("×", "x")
    )


def _short_subject(vendor: VendorReminder, reminder_number: int, return_period: str = "") -> str:
    """Keep subject short — long subjects break mailto and get truncated in Gmail."""
    period_note = f" ({return_period})" if return_period else ""
    amount = vendor.blocked_itc
    count = vendor.invoice_count
    if reminder_number <= 1:
        return (
            f"Request: File/amend GSTR-1{period_note} - "
            f"ITC blocked ({count} invoices, Rs. {amount:,.2f})"
        )
    if reminder_number == 2:
        return f"Follow-up Day 3: GSTR-1 pending - Rs. {amount:,.2f} ITC blocked"
    return f"Final reminder Day 7: Rs. {amount:,.2f} ITC blocked - action required"


def build_email_compose_links(
    email: str,
    subject: str,
    body: str,
    max_gmail_url: int = 7500,
    max_mailto_url: int = 1800,
) -> dict[str, Any]:
    """Build Gmail web, Outlook web, and mailto links. Gmail web works best in browser."""
    safe_subject = _ascii_for_email_url(subject)
    safe_body = _ascii_for_email_url(body)

    def gmail_url(body_text: str) -> str:
        return (
            "https://mail.google.com/mail/?view=cm&fs=1"
            f"&to={quote(email)}"
            f"&su={quote(safe_subject)}"
            f"&body={quote(body_text)}"
        )

    def outlook_url(body_text: str) -> str:
        return (
            "https://outlook.live.com/mail/0/deeplink/compose"
            f"?to={quote(email)}"
            f"&subject={quote(safe_subject)}"
            f"&body={quote(body_text)}"
        )

    body_for_url = safe_body
    url_truncated = False
    gmail = gmail_url(body_for_url)
    if len(gmail) > max_gmail_url:
        body_for_url = (
            safe_body[:1200]
            + "\n\n[Full message shortened in link. Use Copy text in the app and paste into Gmail.]"
        )
        url_truncated = True
        gmail = gmail_url(body_for_url)
    if len(gmail) > max_gmail_url:
        body_for_url = (
            safe_body[:500]
            + "\n\n[Copy the full email from the app using Copy text, then paste into Gmail body.]"
        )
        gmail = gmail_url(body_for_url)

    mailto_body = body_for_url
    mailto = f"mailto:{quote(email)}?subject={quote(safe_subject)}&body={quote(mailto_body)}"
    if len(mailto) > max_mailto_url:
        mailto_body = (
            safe_body[:400]
            + "\n\n[Copy full message from app and paste into your email.]"
        )
        mailto = f"mailto:{quote(email)}?subject={quote(safe_subject)}&body={quote(mailto_body)}"
        if len(mailto) > max_mailto_url:
            mailto = f"mailto:{quote(email)}?subject={quote(safe_subject)}"

    return {
        "gmail_link": gmail,
        "outlook_link": outlook_url(body_for_url),
        "mailto_link": mailto,
        "url_truncated": url_truncated or body_for_url != safe_body,
    }


def build_mailto_link(email: str, subject: str, body: str, max_length: int = 1800) -> str:
    return build_email_compose_links(email, subject, body, max_mailto_url=max_length)["mailto_link"]


def create_manual_followup_log_entry(
    vendor: VendorReminder,
    reminder_number: int,
    subject: str,
) -> FollowUpLogEntry:
    sent_at = datetime.now()
    return FollowUpLogEntry(
        supplier_gstin=vendor.supplier_gstin,
        supplier_name=vendor.supplier_name,
        email=vendor.email,
        reminder_number=reminder_number,
        sent_at=sent_at,
        blocked_itc=vendor.blocked_itc,
        next_follow_up_date=_next_follow_up_date(reminder_number, sent_at.date()),
        status="Sent via email client",
        subject=subject,
    )


def build_vendor_mailto_entries(
    vendors: list[VendorReminder],
    reminder_number: int,
    company_name: str,
    sender_name: str,
    return_period: str = "",
) -> list[dict[str, Any]]:
    entries: list[dict[str, Any]] = []
    for vendor in vendors:
        if not vendor.email or "@" not in vendor.email:
            continue
        subject, body = build_reminder_email(
            vendor, reminder_number, company_name, sender_name, return_period
        )
        links = build_email_compose_links(vendor.email, subject, body)
        entries.append(
            {
                "vendor": vendor,
                "supplier_name": vendor.supplier_name,
                "email": vendor.email,
                "blocked_itc": vendor.blocked_itc,
                "subject": subject,
                "body": body,
                "gmail_link": links["gmail_link"],
                "outlook_link": links["outlook_link"],
                "mailto_link": links["mailto_link"],
                "url_truncated": links["url_truncated"],
                "reminder_number": reminder_number,
            }
        )
    return entries


def reminder_type_label(reminder_number: int) -> str:
    labels = {1: "Initial reminder", 2: "Day 3 follow-up", 3: "Day 7 final reminder"}
    return labels.get(reminder_number, f"Reminder #{reminder_number}")


def _next_follow_up_date(reminder_number: int, sent_on: date) -> date | None:
    if reminder_number == 1:
        return sent_on + timedelta(days=FOLLOWUP_INTERVALS_DAYS[0])
    if reminder_number == 2:
        return sent_on + timedelta(days=FOLLOWUP_INTERVALS_DAYS[1] - FOLLOWUP_INTERVALS_DAYS[0])
    return None


def send_vendor_email(
    smtp: SmtpConfig,
    to_email: str,
    subject: str,
    body: str,
) -> tuple[bool, str]:
    if not to_email or "@" not in to_email:
        return False, "Missing or invalid vendor email"

    message = MIMEMultipart()
    message["From"] = f"{smtp.from_name} <{smtp.from_email}>"
    message["To"] = to_email
    message["Subject"] = subject
    message.attach(MIMEText(body, "plain", "utf-8"))

    try:
        context = ssl.create_default_context()
        with smtplib.SMTP(smtp.host, smtp.port, timeout=30) as server:
            server.starttls(context=context)
            server.login(smtp.user, smtp.password)
            server.sendmail(smtp.from_email, [to_email], message.as_string())
        return True, "Sent"
    except Exception as exc:
        return False, str(exc)


def send_vendor_reminder(
    vendor: VendorReminder,
    reminder_number: int,
    smtp: SmtpConfig | None,
    company_name: str,
    sender_name: str,
    return_period: str = "",
) -> SendResult:
    subject, body = build_reminder_email(
        vendor, reminder_number, company_name, sender_name, return_period
    )
    if not smtp:
        return SendResult(
            vendor.supplier_name,
            vendor.email,
            False,
            "SMTP not configured — use mailto link or add email settings in Streamlit secrets",
        )

    ok, message = send_vendor_email(smtp, vendor.email, subject, body)
    log_entry = None
    if ok:
        sent_at = datetime.now()
        log_entry = FollowUpLogEntry(
            supplier_gstin=vendor.supplier_gstin,
            supplier_name=vendor.supplier_name,
            email=vendor.email,
            reminder_number=reminder_number,
            sent_at=sent_at,
            blocked_itc=vendor.blocked_itc,
            next_follow_up_date=_next_follow_up_date(reminder_number, sent_at.date()),
            status="Sent",
            subject=subject,
        )
    return SendResult(vendor.supplier_name, vendor.email, ok, message, log_entry)


def merge_editor_into_reminders(
    reminders: list[VendorReminder],
    editor_df: pd.DataFrame,
) -> list[VendorReminder]:
    """Apply email / select edits from the Streamlit data editor."""
    if editor_df.empty:
        return []

    by_gstin = {item.supplier_gstin: item for item in reminders}
    updated: list[VendorReminder] = []
    for _, row in editor_df.iterrows():
        if not row.get("Select", True):
            continue
        gstin = str(row.get("Supplier GSTIN", "")).strip()
        base = by_gstin.get(gstin)
        if not base:
            continue
        email = str(row.get("Vendor Email", "")).strip()
        updated.append(
            VendorReminder(
                supplier_gstin=base.supplier_gstin,
                supplier_name=base.supplier_name,
                email=email,
                blocked_itc=base.blocked_itc,
                invoice_count=base.invoice_count,
                invoice_list=base.invoice_list,
                primary_issue=base.primary_issue,
                issue_summary=base.issue_summary,
            )
        )
    return updated


def append_followup_log(
    existing: list[dict[str, Any]],
    entries: list[FollowUpLogEntry],
) -> list[dict[str, Any]]:
    log = list(existing)
    for entry in entries:
        log.append(
            {
                "Supplier GSTIN": entry.supplier_gstin,
                "Supplier Name": entry.supplier_name,
                "Email": entry.email,
                "Reminder #": entry.reminder_number,
                "Sent At": entry.sent_at.isoformat(timespec="seconds"),
                "Blocked ITC (₹)": entry.blocked_itc,
                "Next Follow-up": entry.next_follow_up_date.isoformat()
                if entry.next_follow_up_date
                else "",
                "Status": entry.status,
                "Subject": entry.subject,
            }
        )
    return log


def followup_log_dataframe(log: list[dict[str, Any]]) -> pd.DataFrame:
    if not log:
        return pd.DataFrame(
            columns=[
                "Supplier GSTIN",
                "Supplier Name",
                "Email",
                "Reminder #",
                "Sent At",
                "Blocked ITC (₹)",
                "Next Follow-up",
                "Status",
                "Subject",
            ]
        )
    return pd.DataFrame(log)


def get_due_followups(
    log: list[dict[str, Any]],
    reminders: list[VendorReminder],
    today: date | None = None,
) -> list[tuple[VendorReminder, int]]:
    """Return vendors due for the next automated follow-up email."""
    today = today or date.today()
    due: list[tuple[VendorReminder, int]] = []
    reminder_by_gstin = {item.supplier_gstin: item for item in reminders}

    latest: dict[str, dict[str, Any]] = {}
    for row in log:
        gstin = str(row.get("Supplier GSTIN", ""))
        if not gstin:
            continue
        sent_at = row.get("Sent At", "")
        if gstin not in latest or sent_at > latest[gstin].get("Sent At", ""):
            latest[gstin] = row

    for gstin, row in latest.items():
        next_date_str = str(row.get("Next Follow-up", "")).strip()
        if not next_date_str:
            continue
        try:
            next_date = date.fromisoformat(next_date_str[:10])
        except ValueError:
            continue
        if next_date > today:
            continue
        vendor = reminder_by_gstin.get(gstin)
        if not vendor:
            continue
        last_num = int(row.get("Reminder #", 1))
        if last_num >= 3:
            continue
        due.append((vendor, last_num + 1))

    return due


def export_followup_log_excel(log: list[dict[str, Any]]) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        followup_log_dataframe(log).to_excel(writer, sheet_name="Follow-up Log", index=False)
    return buffer.getvalue()


def scheduled_followups_preview(log: list[dict[str, Any]]) -> pd.DataFrame:
    rows = []
    for row in log:
        next_date = str(row.get("Next Follow-up", "")).strip()
        if not next_date:
            continue
        reminder_num = int(row.get("Reminder #", 1))
        if reminder_num >= 3:
            continue
        rows.append(
            {
                "Supplier": row.get("Supplier Name"),
                "GSTIN": row.get("Supplier GSTIN"),
                "Last Reminder": f"#{reminder_num}",
                "Next Follow-up Due": next_date[:10],
                "Blocked ITC (₹)": row.get("Blocked ITC (₹)"),
            }
        )
    if not rows:
        return pd.DataFrame(columns=["Supplier", "GSTIN", "Last Reminder", "Next Follow-up Due", "Blocked ITC (₹)"])
    return pd.DataFrame(rows).drop_duplicates(subset=["GSTIN"], keep="last")
