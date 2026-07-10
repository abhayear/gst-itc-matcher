"""Streamlit app for GST ITC invoice matching — upload files, get results."""

from __future__ import annotations

import sys
from pathlib import Path

# Ensure project root is on path for Streamlit Cloud
ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pandas as pd
import streamlit as st

st.set_page_config(page_title="GST ITC Matcher", page_icon=":bar_chart:", layout="wide")

try:
    from matcher.consolidate import (
        consolidated_gstr_to_display,
        consolidated_pr_to_display,
        export_consolidated_gstr,
        export_consolidated_purchase_register,
        gstr_summary_caption,
        label_from_filename,
    )
    from matcher.engine import (
        export_to_csv,
        export_to_excel,
        export_to_pdf,
        generate_ai_recommendations,
        load_and_match_with_consolidation,
    )
    from matcher.itc_dashboard import (
        ITC_ELIGIBLE,
        ITC_NON_ELIGIBLE,
        ITC_PENDING,
        action_plan_dataframe,
        build_itc_claim_plan,
        generate_action_plan,
        generate_optimization_insights,
    )
    from matcher.vendor_followup import (
        append_followup_log,
        build_vendor_mailto_entries,
        create_manual_followup_log_entry,
        create_manual_followup_log_entry,
        export_followup_log_excel,
        extract_vendor_reminders,
        followup_log_dataframe,
        get_due_followups,
        load_smtp_config,
        merge_editor_into_reminders,
        reminder_type_label,
        scheduled_followups_preview,
        send_vendor_reminder,
        vendor_reminders_dataframe,
    )
except Exception as import_error:
    st.error("App failed to start. Please refresh in a minute or contact support.")
    st.exception(import_error)
    st.stop()


def _secrets_dict() -> dict:
    try:
        return dict(st.secrets)
    except Exception:
        return {}


st.title("GST ITC Matcher")
st.caption("Upload Purchase Register(s) and GSTR-2A/2B — consolidation and matching run automatically.")

col_pr, col_gstr = st.columns(2)
with col_pr:
    pr_mode = st.radio(
        "Purchase Register",
        options=["Single file", "Sales + Service (Consolidate)"],
        horizontal=True,
    )
with col_gstr:
    gstr_mode = st.radio(
        "GSTR-2A / 2B",
        options=["Single file", "Multiple periods (Consolidate)"],
        horizontal=True,
    )

sales_pr = service_pr = pr_file = None
gstr_file = None
gstr_period_files: list = []

st.subheader("Purchase Register")
if pr_mode == "Sales + Service (Consolidate)":
    c1, c2 = st.columns(2)
    with c1:
        sales_pr = st.file_uploader("Sales Purchase Register", type=["xlsx", "xls"], key="sales_pr")
    with c2:
        service_pr = st.file_uploader("Service Purchase Register", type=["xlsx", "xls"], key="service_pr")
else:
    st.caption(
        "Upload **Vyapar Purchase Report** Excel (must include **Purchase Items** sheet). "
        "Reports → Purchase Report → Excel export. "
        "The app converts Purchase Items into sample Purchase Register format automatically."
    )
    pr_file = st.file_uploader("Purchase Register (Excel)", type=["xlsx", "xls"], key="pr")

st.subheader("GSTR-2A / 2B")
if gstr_mode == "Multiple periods (Consolidate)":
    st.caption(
        "Upload one GSTR-2B Excel per period from the **GST portal** "
        "(e.g. 092025_GSTIN_GSTR2BQ_....xlsx). "
        "Do not upload Payment Register or Tally/Busy exports here."
    )
    gstr_period_files = st.file_uploader(
        "GSTR-2A/2B files (one per period)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="gstr_periods",
    ) or []
else:
    st.caption(
        "Upload the **GSTR-2B Excel downloaded from GST Portal** "
        "(e.g. `09GSTIN_062026_GSTR2BQ_08072026.xlsx` with B2B/CDNR sheets). "
        "Do **not** upload Payment Register, Tally payment exports, or files named only with a date."
    )
    gstr_file = st.file_uploader("GSTR-2A / 2B (Excel)", type=["xlsx", "xls"], key="gstr")

pr_ready = (sales_pr and service_pr) if pr_mode == "Sales + Service (Consolidate)" else bool(pr_file)
gstr_ready = len(gstr_period_files) >= 1 if gstr_mode == "Multiple periods (Consolidate)" else bool(gstr_file)
files_ready = pr_ready and gstr_ready


def _build_file_key() -> str:
    parts = [pr_mode, gstr_mode]
    if pr_mode == "Sales + Service (Consolidate)":
        parts.extend([sales_pr.name, str(sales_pr.size), service_pr.name, str(service_pr.size)])
    else:
        parts.extend([pr_file.name, str(pr_file.size)])
    if gstr_mode == "Multiple periods (Consolidate)":
        for uploaded in gstr_period_files:
            parts.extend([uploaded.name, str(uploaded.size)])
    else:
        parts.extend([gstr_file.name, str(gstr_file.size)])
    return "|".join(parts)


def _build_sources() -> tuple[list, list]:
    if pr_mode == "Sales + Service (Consolidate)":
        purchase_sources = [(sales_pr, "Sales"), (service_pr, "Service")]
    else:
        purchase_sources = [(pr_file, "Purchase Register")]

    if gstr_mode == "Multiple periods (Consolidate)":
        gstr_sources = [(f, label_from_filename(f.name)) for f in gstr_period_files]
    else:
        gstr_sources = [(gstr_file, "GSTR")]

    return purchase_sources, gstr_sources


if files_ready:
    file_key = _build_file_key()

    if st.session_state.get("last_file_key") != file_key:
        with st.spinner("Consolidating registers and matching invoices..."):
            try:
                purchase_sources, gstr_sources = _build_sources()
                for source, _ in purchase_sources + gstr_sources:
                    source.seek(0)

                consolidated_pr, consolidated_gstr, result, summary, dashboard = load_and_match_with_consolidation(
                    purchase_sources,
                    gstr_sources,
                )
                st.session_state["consolidated_pr"] = consolidated_pr
                st.session_state["consolidated_gstr"] = consolidated_gstr
                st.session_state["match_result"] = result
                st.session_state["match_summary"] = summary
                st.session_state["itc_dashboard"] = dashboard
                st.session_state["last_file_key"] = file_key
            except Exception as exc:
                for key in (
                    "match_result",
                    "match_summary",
                    "itc_dashboard",
                    "consolidated_pr",
                    "consolidated_gstr",
                    "last_file_key",
                ):
                    st.session_state.pop(key, None)
                st.error(f"Could not process files: {exc}")

    if "match_result" in st.session_state:
        result: pd.DataFrame = st.session_state["match_result"]
        summary = st.session_state["match_summary"]
        dashboard = st.session_state["itc_dashboard"]

        st.success("Done! Review the ITC dashboard below or download reports.")

        st.subheader("ITC Dashboard")
        d1, d2, d3, d4 = st.columns(4)
        d1.metric("Available ITC (GSTR-2B)", f"₹{dashboard.available_total:,.2f}")
        d2.metric(
            "Eligible ITC",
            f"₹{dashboard.eligible_total:,.2f}",
            help="Safe to claim in GSTR-3B after verification",
        )
        d3.metric("Pending Review", f"₹{dashboard.pending_total:,.2f}")
        d4.metric("Non-Eligible ITC", f"₹{dashboard.non_eligible_total:,.2f}")

        c1, c2 = st.columns(2)
        with c1:
            st.markdown("**ITC balance overview**")
            balance_df = pd.DataFrame(
                {
                    "Amount": [
                        dashboard.eligible_total,
                        dashboard.pending_total,
                        dashboard.non_eligible_total,
                    ]
                },
                index=["Eligible", "Pending Review", "Non-Eligible"],
            )
            st.bar_chart(balance_df, height=280)
        with c2:
            st.markdown("**Available vs eligible by tax head**")
            tax_df = pd.DataFrame(
                {
                    "Available (GSTR-2B)": [
                        dashboard.available_igst,
                        dashboard.available_cgst,
                        dashboard.available_sgst,
                    ],
                    "Eligible ITC": [
                        dashboard.eligible_igst,
                        dashboard.eligible_cgst,
                        dashboard.eligible_sgst,
                    ],
                },
                index=["IGST", "CGST", "SGST"],
            )
            st.bar_chart(tax_df, height=280)

        inv1, inv2, inv3, inv4 = st.columns(4)
        inv1.metric("Eligible invoices", dashboard.eligible_invoices)
        inv2.metric("Pending invoices", dashboard.pending_invoices)
        inv3.metric("Non-eligible invoices", dashboard.non_eligible_invoices)
        inv4.metric("Claim rate", f"{dashboard.claim_rate_pct}%")

        claim_plan = build_itc_claim_plan(dashboard, result)
        optimization_insights = generate_optimization_insights(dashboard, claim_plan, result)
        action_plan = generate_action_plan(dashboard, result)

        st.subheader("ITC Claim & Planning")
        readiness = claim_plan.filing_readiness
        if readiness == "Ready to File":
            st.success(f"**{readiness}** — {claim_plan.filing_advice}")
        elif readiness in ("File with Caution", "Partial Claim Only"):
            st.warning(f"**{readiness}** — {claim_plan.filing_advice}")
        else:
            st.error(f"**{readiness}** — {claim_plan.filing_advice}")

        p1, p2, p3, p4 = st.columns(4)
        p1.metric("Claim Now (GSTR-3B)", f"₹{claim_plan.claim_now_total:,.2f}")
        p2.metric("Hold (Pending)", f"₹{claim_plan.hold_itc:,.2f}")
        p3.metric("Blocked", f"₹{claim_plan.blocked_itc:,.2f}")
        p4.metric(
            "Optimized Potential",
            f"₹{claim_plan.optimized_claim_total:,.2f}",
            delta=f"+₹{claim_plan.optimization_gain:,.2f}" if claim_plan.optimization_gain > 0 else None,
        )

        t1, t2 = st.columns(2)
        with t1:
            st.markdown("**GSTR-3B Table 4 — Claim amounts**")
            table4_df = pd.DataFrame(
                {
                    "Claim Now": [
                        f"₹{claim_plan.claim_now_igst:,.2f}",
                        f"₹{claim_plan.claim_now_cgst:,.2f}",
                        f"₹{claim_plan.claim_now_sgst:,.2f}",
                    ],
                    "Hold (Pending)": [
                        f"₹{dashboard.pending_igst:,.2f}",
                        f"₹{dashboard.pending_cgst:,.2f}",
                        f"₹{dashboard.pending_sgst:,.2f}",
                    ],
                },
                index=["IGST", "CGST", "SGST"],
            )
            st.dataframe(table4_df, use_container_width=True)
        with t2:
            st.markdown("**Current vs optimized claim**")
            claim_compare_df = pd.DataFrame(
                {
                    "Amount": [
                        claim_plan.claim_now_total,
                        claim_plan.optimized_claim_total,
                    ]
                },
                index=["Current Claim", "Optimized Potential"],
            )
            st.bar_chart(claim_compare_df, height=280)

        st.markdown("**Optimization insights**")
        for insight in optimization_insights:
            clean = insight.replace("**", "")
            st.info(clean)

        st.markdown("**Action plan**")
        st.dataframe(
            action_plan_dataframe(action_plan),
            use_container_width=True,
            hide_index=True,
        )

        vendor_reminders = extract_vendor_reminders(
            result,
            st.session_state.get("vendor_contacts_df"),
        )
        if "vendor_followup_log" not in st.session_state:
            st.session_state["vendor_followup_log"] = []

        st.subheader("Vendor ITC Recovery — Email Reminders")
        st.caption(
            "Email non-compliant vendors to recover blocked ITC. "
            "Use **Open in email** (Outlook/Gmail — no setup) or optional SMTP for automatic send."
        )

        with st.expander("Email settings (company name, contacts, optional SMTP)", expanded=False):
            es1, es2, es3 = st.columns(3)
            company_name = es1.text_input(
                "Your company name",
                value=st.session_state.get("vf_company_name", ""),
                key="vf_company_name_input",
            )
            sender_name = es2.text_input(
                "Sender name",
                value=st.session_state.get("vf_sender_name", "Accounts Team"),
                key="vf_sender_name_input",
            )
            return_period = es3.text_input(
                "Return period (optional)",
                value=st.session_state.get("vf_return_period", ""),
                placeholder="e.g. May 2025",
                key="vf_return_period_input",
            )
            st.session_state["vf_company_name"] = company_name
            st.session_state["vf_sender_name"] = sender_name
            st.session_state["vf_return_period"] = return_period

            smtp_cfg = load_smtp_config(_secrets_dict())
            if smtp_cfg:
                st.success(f"Optional SMTP active — can auto-send from **{smtp_cfg.from_email}**")
            else:
                st.info(
                    "No SMTP configured — that's fine. Use **Open in email** below to send via "
                    "Outlook, Gmail, or your default mail app. No secrets or setup required."
                )

            contacts_file = st.file_uploader(
                "Vendor contacts (optional CSV/Excel: GSTIN + Email)",
                type=["csv", "xlsx", "xls"],
                key="vendor_contacts_upload",
            )
            if contacts_file is not None:
                try:
                    if contacts_file.name.lower().endswith(".csv"):
                        st.session_state["vendor_contacts_df"] = pd.read_csv(contacts_file)
                    else:
                        st.session_state["vendor_contacts_df"] = pd.read_excel(contacts_file)
                    st.caption(f"Loaded {len(st.session_state['vendor_contacts_df'])} contact row(s)")
                except Exception as exc:
                    st.error(f"Could not read contacts file: {exc}")

        if not vendor_reminders:
            st.info("No vendor follow-up needed — all blocked ITC is internal (not supplier-related).")
        else:
            total_blocked = sum(v.blocked_itc for v in vendor_reminders)
            v1, v2, v3 = st.columns(3)
            v1.metric("Non-compliant vendors", len(vendor_reminders))
            v2.metric("Blocked ITC (recoverable)", f"₹{total_blocked:,.2f}")
            due_followups = get_due_followups(st.session_state["vendor_followup_log"], vendor_reminders)
            v3.metric("Follow-ups due today", len(due_followups))

            editor_key = f"vendor_editor_{st.session_state.get('last_file_key', '')}"
            vendor_df = vendor_reminders_dataframe(vendor_reminders)
            edited_vendors = st.data_editor(
                vendor_df,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                column_config={
                    "Select": st.column_config.CheckboxColumn("Send?", default=True),
                    "Vendor Email": st.column_config.TextColumn("Vendor Email", required=False),
                    "Blocked ITC (₹)": st.column_config.NumberColumn(format="₹%.2f"),
                },
                disabled=[
                    "Supplier Name",
                    "Supplier GSTIN",
                    "Blocked ITC (₹)",
                    "Invoices",
                    "Invoice Nos.",
                    "Issue",
                    "Details",
                ],
                key=editor_key,
            )

            selected_vendors = merge_editor_into_reminders(vendor_reminders, edited_vendors)
            vendors_with_email = [v for v in selected_vendors if v.email and "@" in v.email]
            missing_emails = [v.supplier_name for v in selected_vendors if not v.email or "@" not in v.email]
            if missing_emails:
                st.warning(
                    f"Add email for {len(missing_emails)} vendor(s) before sending: "
                    + ", ".join(missing_emails[:5])
                    + ("…" if len(missing_emails) > 5 else "")
                )

            company_name = st.session_state.get("vf_company_name", "Our Company")
            sender_name = st.session_state.get("vf_sender_name", "Accounts Team")
            return_period = st.session_state.get("vf_return_period", "")
            smtp_cfg = load_smtp_config(_secrets_dict())
            due_followups = get_due_followups(st.session_state["vendor_followup_log"], vendor_reminders)

            st.markdown("**Open in email (Outlook / Gmail — recommended, no setup)**")
            st.caption(
                "Use **Open in Gmail** if you use Gmail in the browser. "
                "`mailto:` links often open Outlook/desktop Mail instead of Gmail, and long URLs can fail. "
                "After sending, click **Mark as sent** to schedule follow-ups."
            )

            default_reminder = 2 if due_followups else 1
            reminder_choice = st.radio(
                "Email type",
                options=[1, 2, 3],
                format_func=reminder_type_label,
                index=max(0, default_reminder - 1),
                horizontal=True,
                key="vf_reminder_type",
            )

            mailto_entries = build_vendor_mailto_entries(
                vendors_with_email,
                reminder_choice,
                company_name,
                sender_name,
                return_period,
            )

            if not mailto_entries:
                st.info("Add vendor email addresses above to enable Open in email.")
            else:
                for idx, entry in enumerate(mailto_entries):
                    vc1, vc2, vc3, vc4 = st.columns([3, 1.5, 1.5, 1.5])
                    with vc1:
                        st.markdown(
                            f"**{entry['supplier_name']}**  \n"
                            f"{entry['email']} · ₹{entry['blocked_itc']:,.2f} blocked"
                        )
                        if entry.get("url_truncated"):
                            st.caption(
                                "Long message — Gmail link may be shortened. "
                                "Use **Copy text** and paste into the email body."
                            )
                    with vc2:
                        st.link_button(
                            "Open in Gmail",
                            entry["gmail_link"],
                            use_container_width=True,
                            help="Opens Gmail compose in your browser (recommended for Gmail users)",
                            key=f"gmail_{entry['vendor'].supplier_gstin}_{reminder_choice}_{idx}",
                        )
                    with vc3:
                        st.link_button(
                            "Outlook web",
                            entry["outlook_link"],
                            use_container_width=True,
                            help="Opens Outlook.com compose in browser",
                            key=f"outlook_{entry['vendor'].supplier_gstin}_{reminder_choice}_{idx}",
                        )
                    with vc4:
                        with st.popover("Copy text"):
                            st.text_input("Subject", entry["subject"], disabled=True, key=f"subj_{idx}")
                            st.text_area(
                                "Body (select all & copy, then paste in Gmail)",
                                entry["body"],
                                height=220,
                                key=f"body_{idx}",
                            )
                            st.link_button(
                                "Desktop app (mailto)",
                                entry["mailto_link"],
                                use_container_width=True,
                                help="Opens default desktop mail app — may not work with Gmail web",
                                key=f"mailto_{entry['vendor'].supplier_gstin}_{reminder_choice}_{idx}",
                            )

                mark_col, log_col = st.columns(2)
                with mark_col:
                    mark_sent = st.button(
                        f"Mark {len(mailto_entries)} email(s) as sent — schedule follow-ups",
                        type="primary",
                        use_container_width=True,
                        disabled=not mailto_entries,
                    )
                with log_col:
                    st.download_button(
                        "Download follow-up log",
                        data=export_followup_log_excel(st.session_state["vendor_followup_log"]),
                        file_name="vendor_followup_log.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                if mark_sent and mailto_entries:
                    sent_entries = [
                        create_manual_followup_log_entry(
                            entry["vendor"],
                            entry["reminder_number"],
                            entry["subject"],
                        )
                        for entry in mailto_entries
                    ]
                    st.session_state["vendor_followup_log"] = append_followup_log(
                        st.session_state["vendor_followup_log"],
                        sent_entries,
                    )
                    st.success(
                        f"Logged {len(sent_entries)} reminder(s). "
                        f"Day 3 and Day 7 follow-ups are now scheduled — return and use "
                        f"**{reminder_type_label(min(reminder_choice + 1, 3))}** when due."
                    )

            if smtp_cfg:
                st.divider()
                st.markdown("**Optional: auto-send via SMTP**")
                btn_col1, btn_col2 = st.columns(2)
                with btn_col1:
                    send_all = st.button(
                        "Auto-send to all vendors (SMTP)",
                        use_container_width=True,
                        disabled=not vendors_with_email,
                    )
                with btn_col2:
                    send_due = st.button(
                        f"Auto-send due follow-ups ({len(due_followups)})",
                        use_container_width=True,
                        disabled=not due_followups,
                    )

                if send_all and smtp_cfg:
                    sent_entries = []
                    errors = []
                    for vendor in vendors_with_email:
                        outcome = send_vendor_reminder(
                            vendor,
                            reminder_choice,
                            smtp_cfg,
                            company_name,
                            sender_name,
                            return_period,
                        )
                        if outcome.success and outcome.log_entry:
                            sent_entries.append(outcome.log_entry)
                        else:
                            errors.append(f"{vendor.supplier_name}: {outcome.message}")
                    if sent_entries:
                        st.session_state["vendor_followup_log"] = append_followup_log(
                            st.session_state["vendor_followup_log"],
                            sent_entries,
                        )
                        st.success(f"SMTP sent {len(sent_entries)} email(s).")
                    for err in errors:
                        st.error(err)

                if send_due and smtp_cfg:
                    sent_entries = []
                    for vendor, reminder_num in due_followups:
                        outcome = send_vendor_reminder(
                            vendor,
                            reminder_num,
                            smtp_cfg,
                            company_name,
                            sender_name,
                            return_period,
                        )
                        if outcome.success and outcome.log_entry:
                            sent_entries.append(outcome.log_entry)
                        else:
                            st.error(f"{vendor.supplier_name}: {outcome.message}")
                    if sent_entries:
                        st.session_state["vendor_followup_log"] = append_followup_log(
                            st.session_state["vendor_followup_log"],
                            sent_entries,
                        )
                        st.success(f"SMTP sent {len(sent_entries)} follow-up email(s).")

            scheduled = scheduled_followups_preview(st.session_state["vendor_followup_log"])
            if not scheduled.empty:
                st.markdown("**Scheduled follow-ups**")
                st.dataframe(scheduled, use_container_width=True, hide_index=True)
                if due_followups:
                    st.info(
                        f"{len(due_followups)} follow-up(s) due today — select "
                        f"**{reminder_type_label(due_followups[0][1])}** above and use Open in email."
                    )

            if st.session_state["vendor_followup_log"]:
                with st.expander("Follow-up history"):
                    st.dataframe(
                        followup_log_dataframe(st.session_state["vendor_followup_log"]),
                        use_container_width=True,
                        hide_index=True,
                    )

        st.subheader("AI Recommendations")
        for tip in generate_ai_recommendations(dashboard, result):
            st.info(tip)

        st.markdown("**Download ITC Results**")
        r1, r2, r3 = st.columns(3)
        with r1:
            st.download_button(
                label="Download Excel (.xlsx)",
                data=export_to_excel(result, summary, dashboard),
                file_name="itc_matching_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        with r2:
            st.download_button(
                label="Download CSV (.csv)",
                data=export_to_csv(result, summary, dashboard),
                file_name="itc_matching_report.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with r3:
            st.download_button(
                label="Download PDF (.pdf)",
                data=export_to_pdf(result, summary, dashboard),
                file_name="itc_matching_report.pdf",
                mime="application/pdf",
                use_container_width=True,
            )

        st.markdown("**Download Converted Registers**")
        d2, d3 = st.columns(2)
        with d2:
            if st.session_state.get("consolidated_pr") is not None:
                st.download_button(
                    label="Download Purchase Register (Sample Format)",
                    data=export_consolidated_purchase_register(st.session_state["consolidated_pr"]),
                    file_name="purchase_register_sample.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        with d3:
            if st.session_state.get("consolidated_gstr") is not None:
                st.download_button(
                    label="Download GSTR-2B (Sample Format)",
                    data=export_consolidated_gstr(st.session_state["consolidated_gstr"]),
                    file_name="consolidated_gstr2b.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        consolidated_pr = st.session_state.get("consolidated_pr")
        consolidated_gstr = st.session_state.get("consolidated_gstr")
        if consolidated_pr is not None or consolidated_gstr is not None:
            c1, c2 = st.columns(2)
            with c1:
                if consolidated_pr is not None:
                    st.subheader("Purchase Register (Sample Format)")
                    caption = f"{len(consolidated_pr)} invoices converted from your upload"
                    if (
                        "register_type" in consolidated_pr.columns
                        and consolidated_pr["register_type"].nunique() > 1
                    ):
                        sales_count = len(consolidated_pr[consolidated_pr["register_type"] == "Sales"])
                        service_count = len(consolidated_pr[consolidated_pr["register_type"] == "Service"])
                        caption = f"{sales_count} Sales + {service_count} Service invoices"
                    st.caption(caption)
                    st.dataframe(consolidated_pr_to_display(consolidated_pr), use_container_width=True, height=220)
            with c2:
                if consolidated_gstr is not None:
                    st.subheader("GSTR-2B (Sample Format)")
                    st.caption(
                        f"{len(consolidated_gstr)} invoices converted from GST portal format"
                        + (
                            f" — {gstr_summary_caption(consolidated_gstr)}"
                            if "gstr_source" in consolidated_gstr.columns
                            and consolidated_gstr["gstr_source"].nunique() > 1
                            else ""
                        )
                    )
                    st.dataframe(consolidated_gstr_to_display(consolidated_gstr), use_container_width=True, height=220)

        st.subheader("Matching Results")
        tab_all, tab_eligible, tab_pending, tab_blocked = st.tabs(
            ["All", "Eligible", "Pending Review", "Non-Eligible"]
        )
        with tab_all:
            st.dataframe(result, use_container_width=True, height=400)
        with tab_eligible:
            st.dataframe(
                result[result["ITC Category"] == ITC_ELIGIBLE],
                use_container_width=True,
                height=400,
            )
        with tab_pending:
            st.dataframe(
                result[result["ITC Category"] == ITC_PENDING],
                use_container_width=True,
                height=400,
            )
        with tab_blocked:
            st.dataframe(
                result[result["ITC Category"] == ITC_NON_ELIGIBLE],
                use_container_width=True,
                height=400,
            )

else:
    hints = []
    if not pr_ready:
        hints.append("Purchase Register file(s)")
    if not gstr_ready:
        if gstr_mode == "Multiple periods (Consolidate)":
            hints.append("at least one GSTR-2A/2B file per period")
        else:
            hints.append("GSTR-2A/2B file")
    st.info(f"Upload {' and '.join(hints)} to begin.")

st.divider()
st.caption(
    "Columns are auto-detected from Tally, Busy, SAP, and GST portal exports. "
    "ITC decisions are indicative — verify under GST rules before claiming."
)
