"""CLI entry point for Excel Automation Tool."""
import argparse
from pathlib import Path
from excel_tool.processor import load_files, clean_orders, write_report


def main():
    parser = argparse.ArgumentParser(description="Merge sales CSV/XLSX files and generate an Excel report")
    parser.add_argument("inputs", nargs="+", help="Input CSV or XLSX files (supports glob patterns)")
    parser.add_argument("-o", "--output", default="output/sales_report.xlsx", help="Output XLSX path")
    args = parser.parse_args()
    files = []
    import glob
    for pattern in args.inputs:
        matches = sorted(glob.glob(pattern))
        files.extend(Path(p) for p in matches if Path(p).is_file())
    if not files:
        parser.error("No matching input files")
    try:
        orders = load_files(files)
        valid, rejected, duplicates = clean_orders(orders)
        result = write_report(valid, rejected, args.output)
    except (ValueError, OSError) as exc:
        parser.exit(1, f"Error: {exc}\n")
    print(f"Report: {result}\nValid orders: {len(valid)}\nRejected rows: {len(rejected)}"
          f"\nDuplicate IDs removed: {duplicates}\nRevenue: {valid['total'].sum():,.2f}")


if __name__ == "__main__":
    main()
