#!/usr/bin/env python3
"""CLI version: read a folder of invoices and create journal.csv/journal.xlsx."""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from invoice_engine import (
    SUPPORTED_EXTENSIONS,
    build_voucher_rows,
    demo_invoice,
    load_rules,
    parse_invoice_file,
    write_csv,
    write_xlsx,
)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Read insurance invoices (PDF/JPG/PNG) and export CSV plus vendor-wise Excel journal vouchers. No AI API key required."
    )
    parser.add_argument("--input", "--folder", dest="input_folder", help="Folder containing invoice files")
    parser.add_argument("--output", default="output", help="Output folder (default: output)")
    parser.add_argument("--rules", default=None, help="Path to a rules.json file (default: project rules.json)")
    parser.add_argument("--demo", action="store_true", help="Generate a demo journal to test the rules; no invoice files needed")
    args = parser.parse_args()

    try:
        rules = load_rules(args.rules)
    except Exception as exc:
        print(f"ERROR: Could not load rules: {exc}", file=sys.stderr)
        return 2

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)
    all_rows = []
    failed_files = 0

    if args.demo:
        invoice = demo_invoice()
        rows, flags = build_voucher_rows(invoice, rules)
        all_rows.extend(rows)
        print("DEMO: generated a sample Universal Sompo Karnataka voucher.")
        print(f"DEMO: ledger lines = {len(rows)}")
        print("DEMO: checks = " + ("; ".join(flags) if flags else "none"))
    else:
        if not args.input_folder:
            parser.error("Provide --input <folder> or use --demo")
        input_dir = Path(args.input_folder)
        if not input_dir.exists() or not input_dir.is_dir():
            print(f"ERROR: Input folder does not exist: {input_dir}", file=sys.stderr)
            return 2
        files = sorted(p for p in input_dir.iterdir() if p.is_file() and p.suffix.lower() in SUPPORTED_EXTENSIONS)
        if not files:
            print(f"No supported invoices found in {input_dir}")
            print("Supported: PDF, JPG, JPEG, PNG, TIFF, BMP")
            return 1
        print(f"Found {len(files)} invoice file(s) in {input_dir}")
        for index, path in enumerate(files, start=1):
            print(f"\n[{index}/{len(files)}] Reading {path.name} ...")
            try:
                invoice, method, pages, log = parse_invoice_file(path.name, path.read_bytes())
                print(f"  Method: {method}; pages: {pages}")
                for line in log:
                    print(f"  - {line}")
                rows, flags = build_voucher_rows(invoice, rules)
                all_rows.extend(rows)
                if flags:
                    print("  CHECK: " + "; ".join(flags))
                else:
                    print("  Checks: none")
                print(f"  Journal lines: {len(rows)}")
            except Exception as exc:
                failed_files += 1
                print(f"  FAILED: {exc}", file=sys.stderr)

    csv_path = output_dir / "journal.csv"
    xlsx_path = output_dir / "journal.xlsx"
    write_csv(all_rows, csv_path)
    write_xlsx(all_rows, xlsx_path)

    debit = sum((r["Ledger Amount"] for r in all_rows if r["Ledger Amount Dr/Cr"] == "Dr"), start=0)
    credit = sum((r["Ledger Amount"] for r in all_rows if r["Ledger Amount Dr/Cr"] == "Cr"), start=0)
    print("\nExport complete")
    print(f"  CSV : {csv_path.resolve()}")
    print(f"  XLSX: {xlsx_path.resolve()} (All Vendors + Vendor Index + one tab per Tally vendor ledger)")
    print(f"  Rows: {len(all_rows)}")
    print(f"  Failed files: {failed_files}")
    print(f"  Dr total: {debit}")
    print(f"  Cr total: {credit}")
    if failed_files:
        print("WARNING: One or more files could not be processed. Review the log above.", file=sys.stderr)
        return 1
    if round(float(debit - credit), 2) != 0:
        print("WARNING: Combined debit and credit totals do not match. Review source data and mappings.", file=sys.stderr)
        return 1
    print("  Batch check: Dr and Cr totals match")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
