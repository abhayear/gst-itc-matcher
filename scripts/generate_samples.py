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

    pr_path = SAMPLES_DIR / "sample_purchase_register.xlsx"
    sales_path = SAMPLES_DIR / "sample_sales_purchase_register.xlsx"
    service_path = SAMPLES_DIR / "sample_service_purchase_register.xlsx"
    gstr_path = SAMPLES_DIR / "sample_gstr2b.xlsx"
    sales_gstr_path = SAMPLES_DIR / "sample_sales_gstr2b.xlsx"
    service_gstr_path = SAMPLES_DIR / "sample_service_gstr2b.xlsx"

    purchase_register.to_excel(pr_path, index=False)
    sales_register.to_excel(sales_path, index=False)
    service_register.to_excel(service_path, index=False)
    gstr_2b.to_excel(gstr_path, index=False)
    sales_gstr.to_excel(sales_gstr_path, index=False)
    service_gstr.to_excel(service_gstr_path, index=False)

    print(f"Created {pr_path}")
    print(f"Created {sales_path}")
    print(f"Created {service_path}")
    print(f"Created {gstr_path}")
    print(f"Created {sales_gstr_path}")
    print(f"Created {service_gstr_path}")


if __name__ == "__main__":
    main()
