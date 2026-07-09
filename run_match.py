"""Command-line interface for GST ITC matching."""

from __future__ import annotations

import argparse
from pathlib import Path

from matcher.engine import export_to_excel, load_and_match


def main() -> None:
    parser = argparse.ArgumentParser(description="Match Purchase Register with GSTR-2A/2B")
    parser.add_argument("purchase_register", type=Path, help="Purchase Register Excel file")
    parser.add_argument("gstr_file", type=Path, help="GSTR-2A/2B Excel file")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("itc_matching_report.xlsx"),
        help="Output Excel report path",
    )
    parser.add_argument(
        "--tolerance",
        type=float,
        default=1.0,
        help="Tax amount tolerance in rupees (default: 1.0)",
    )
    args = parser.parse_args()

    result, summary, dashboard = load_and_match(
        args.purchase_register,
        args.gstr_file,
        tax_tolerance=args.tolerance,
    )

    args.output.write_bytes(export_to_excel(result, summary, dashboard))

    print(f"Report saved to {args.output}")
    print(f"Fully Matched: {summary.fully_matched}")
    print(f"Tax Mismatch: {summary.tax_mismatch}")
    print(f"Not Matched: {summary.not_matched}")
    print(f"Duplicate: {summary.duplicate}")
    print(f"Total ITC Taken: {summary.total_itc:,.2f}")


if __name__ == "__main__":
    main()
