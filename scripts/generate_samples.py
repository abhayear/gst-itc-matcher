"""Generate sample Purchase Register and GSTR-2A files for testing."""

from pathlib import Path

import pandas as pd

SAMPLES_DIR = Path(__file__).resolve().parent.parent / "samples"


def main() -> None:
    SAMPLES_DIR.mkdir(exist_ok=True)

    sales_register = pd.DataFrame(
        [
            {
                "Supplier GSTIN": "27AABCU9603R1ZM",
                "Supplier Name": "ABC Traders Pvt Ltd",
                "Invoice No.": "INV-1001",
                "Invoice Date": "2025-04-05",
                "Taxable Value": 100000,
                "IGST": 0,
                "CGST": 9000,
                "SGST": 9000,
            },
            {
                "Supplier GSTIN": "29AACCT1234M1Z5",
                "Supplier Name": "XYZ Supplies",
                "Invoice No.": "XYZ/2025/42",
                "Invoice Date": "2025-04-10",
                "Taxable Value": 50000,
                "IGST": 9000,
                "CGST": 0,
                "SGST": 0,
            },
            {
                "Supplier GSTIN": "27AABCU9603R1ZM",
                "Supplier Name": "ABC Traders Pvt Ltd",
                "Invoice No.": "INV-1001",
                "Invoice Date": "2025-04-05",
                "Taxable Value": 100000,
                "IGST": 0,
                "CGST": 9000,
                "SGST": 9000,
            },
        ]
    )

    service_register = pd.DataFrame(
        [
            {
                "Supplier GSTIN": "27AABCU9603R1ZM",
                "Supplier Name": "ABC Traders Pvt Ltd",
                "Invoice No.": "INV-1002",
                "Invoice Date": "2025-04-12",
                "Taxable Value": 25000,
                "IGST": 0,
                "CGST": 2250,
                "SGST": 2250,
            },
            {
                "Supplier GSTIN": "07AADCB2230M1Z8",
                "Supplier Name": "Delta Services",
                "Invoice No.": "DS-7788",
                "Invoice Date": "2025-04-15",
                "Taxable Value": 30000,
                "IGST": 0,
                "CGST": 2700,
                "SGST": 2700,
            },
        ]
    )

    purchase_register = pd.concat([sales_register, service_register], ignore_index=True)

    gstr_2b = pd.DataFrame(
        [
            {
                "GSTIN of supplier": "27AABCU9603R1ZM",
                "Trade/Legal name": "ABC TRADERS PVT LTD",
                "Invoice number": "INV-1001",
                "Invoice date": "05-04-2025",
                "Taxable Value": 100000,
                "Integrated tax(₹)": 0,
                "Central tax(₹)": 9000,
                "State/UT tax(₹)": 9000,
            },
            {
                "GSTIN of supplier": "29AACCT1234M1Z5",
                "Trade/Legal name": "XYZ SUPPLIES",
                "Invoice number": "XYZ/2025/42",
                "Invoice date": "10-04-2025",
                "Taxable Value": 50000,
                "Integrated tax(₹)": 8500,
                "Central tax(₹)": 0,
                "State/UT tax(₹)": 0,
            },
            {
                "GSTIN of supplier": "27AABCU9603R1ZM",
                "Trade/Legal name": "ABC TRADERS PVT LTD",
                "Invoice number": "INV-1002",
                "Invoice date": "12-04-2025",
                "Taxable Value": 25000,
                "Integrated tax(₹)": 0,
                "Central tax(₹)": 2250,
                "State/UT tax(₹)": 2250,
            },
            {
                "GSTIN of supplier": "19AABCP9876N1Z3",
                "Trade/Legal name": "Omega Industries",
                "Invoice number": "OI-9901",
                "Invoice date": "18-04-2025",
                "Taxable Value": 15000,
                "Integrated tax(₹)": 2700,
                "Central tax(₹)": 0,
                "State/UT tax(₹)": 0,
            },
        ]
    )

    sales_gstr = gstr_2b.iloc[:2].copy()
    service_gstr = gstr_2b.iloc[2:].copy()
    gstr_apr = gstr_2b.iloc[:2].copy()
    gstr_may = gstr_2b.iloc[2:].copy()

    pr_path = SAMPLES_DIR / "sample_purchase_register.xlsx"
    sales_path = SAMPLES_DIR / "sample_sales_purchase_register.xlsx"
    service_path = SAMPLES_DIR / "sample_service_purchase_register.xlsx"
    gstr_path = SAMPLES_DIR / "sample_gstr2b.xlsx"
    sales_gstr_path = SAMPLES_DIR / "sample_sales_gstr2b.xlsx"
    service_gstr_path = SAMPLES_DIR / "sample_service_gstr2b.xlsx"
    gstr_apr_path = SAMPLES_DIR / "sample_gstr2b_apr_2025.xlsx"
    gstr_may_path = SAMPLES_DIR / "sample_gstr2b_may_2025.xlsx"

    purchase_register.to_excel(pr_path, index=False)
    sales_register.to_excel(sales_path, index=False)
    service_register.to_excel(service_path, index=False)
    gstr_2b.to_excel(gstr_path, index=False)
    sales_gstr.to_excel(sales_gstr_path, index=False)
    service_gstr.to_excel(service_gstr_path, index=False)
    gstr_apr.to_excel(gstr_apr_path, index=False)
    gstr_may.to_excel(gstr_may_path, index=False)

    # GST portal style export with title rows and (₹) in headers
    portal_headers = [
        "GSTIN of supplier",
        "Trade/Legal name",
        "Invoice number",
        "Invoice date",
        "Taxable Value (₹)",
        "Integrated Tax (₹)",
        "Central Tax (₹)",
        "State/UT Tax (₹)",
    ]
    portal_rows = [
        ["27AABCU9603R1ZM", "ABC TRADERS PVT LTD", "INV-1001", "05-04-2025", 100000, 0, 9000, 9000],
        ["29AACCT1234M1Z5", "XYZ SUPPLIES", "XYZ/2025/42", "10-04-2025", 50000, 8500, 0, 0],
    ]
    portal_path = SAMPLES_DIR / "sample_gstr2b_portal_format.xlsx"
    with pd.ExcelWriter(portal_path, engine="openpyxl") as writer:
        title_df = pd.DataFrame([["GSTR-2B", ""], ["Tax Period: 03-2026", ""]])
        title_df.to_excel(writer, sheet_name="B2B", index=False, header=False, startrow=0)
        portal_df = pd.DataFrame(portal_rows, columns=portal_headers)
        portal_df.to_excel(writer, sheet_name="B2B", index=False, startrow=4)

    # Realistic GST portal file with many headings and two-row tax header
    messy_path = SAMPLES_DIR / "sample_gstr2b_many_headings.xlsx"
    with pd.ExcelWriter(messy_path, engine="openpyxl") as writer:
        rows = [
            ["GSTR-2B", "", "", "", "", "", "", ""],
            ["Tax Period: 09-2025", "", "", "", "", "", "", ""],
            ["GSTIN: 09AAMFE1697H1ZZ", "", "", "", "", "", "", ""],
            ["", "", "", "", "", "", "", ""],
            ["Document Details", "", "", "", "", "", "", ""],
            ["Supplier wise details", "", "", "", "", "", "", ""],
            [
                "GSTIN of supplier",
                "Trade/Legal name",
                "Invoice number",
                "Invoice date",
                "Taxable Value (₹)",
                "Integrated Tax (₹)",
                "Central Tax (₹)",
                "State/UT Tax (₹)",
            ],
            ["27AABCU9603R1ZM", "ABC TRADERS PVT LTD", "INV-1001", "05-04-2025", 100000, 0, 9000, 9000],
            ["29AACCT1234M1Z5", "XYZ SUPPLIES", "XYZ/2025/42", "10-04-2025", 50000, 8500, 0, 0],
            ["Document Details", "", "", "", "", "", "", ""],
            [
                "GSTIN of supplier",
                "Trade/Legal name",
                "Invoice number",
                "Invoice date",
                "Taxable Value (₹)",
                "Integrated Tax (₹)",
                "Central Tax (₹)",
                "State/UT Tax (₹)",
            ],
            ["27AABCU9603R1ZM", "ABC TRADERS PVT LTD", "INV-1002", "12-04-2025", 25000, 0, 2250, 2250],
        ]
        pd.DataFrame(rows).to_excel(writer, sheet_name="B2B", index=False, header=False)

    # Exact GST portal layout (B2B sheet with rows 1-6 headers, data from row 7)
    portal_b2b_path = SAMPLES_DIR / "sample_gstr2b_gst_portal_b2b.xlsx"
    with pd.ExcelWriter(portal_b2b_path, engine="openpyxl") as writer:
        rows = [
            ["", "", "", "", "", "", "", "", "", "", "", "", ""],
            ["", "Goods and Services Tax", "", "", "", "", "", "", "", "", "", "", ""],
            ["", "Government of India", "", "", "", "", "", "", "", "", "", "", ""],
            ["", "", "Invoice details", "", "", "", "", "", "", "", "", "", ""],
            [
                "GSTIN of supplier",
                "Trade/Legal name of the Supplier",
                "Invoice number",
                "Invoice type",
                "Invoice Date",
                "Invoice Value (₹)",
                "Place of supply",
                "Supply Attract Reverse Charge",
                "Rate (%)",
                "Taxable Value (₹)",
                "Integrated Tax (₹)",
                "Central Tax (₹)",
                "State/UT Tax (₹)",
            ],
            ["", "", "", "", "", "", "", "", "", "", "", "", ""],
            [
                "09AAHCE2207M1ZJ",
                "ELYF EVSPARE PRIVATE",
                "UPND12627005110",
                "Regular",
                "25-06-2026",
                11610,
                "09-Uttar Pradesh",
                "N",
                18,
                9839.38,
                1771.08,
                0,
                0,
            ],
            ["", "", "", "", "", "", "", "", "-", "", "", "", ""],
            ["", "", "", "", "", "", "", "", "", "", "", "", ""],
            [
                "09AAFCG9772A1Z5",
                "GALLANT OIL AND LUBRICANTS",
                "GOL/25-26/001",
                "Regular",
                "15-06-2026",
                5000,
                "09-Uttar Pradesh",
                "N",
                18,
                4237.29,
                762.71,
                0,
                0,
            ],
        ]
        pd.DataFrame(rows).to_excel(writer, sheet_name="B2B", index=False, header=False)
        pd.DataFrame([["Read me"]]).to_excel(writer, sheet_name="Read me", index=False, header=False)

    for path in (
        pr_path,
        sales_path,
        service_path,
        gstr_path,
        sales_gstr_path,
        service_gstr_path,
        gstr_apr_path,
        gstr_may_path,
        portal_path,
        messy_path,
        portal_b2b_path,
    ):
        print(f"Created {path}")


if __name__ == "__main__":
    main()
