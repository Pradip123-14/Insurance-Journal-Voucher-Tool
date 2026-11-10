from datetime import date
from decimal import Decimal
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent))
from invoice_engine import (
    _extract_policy_number, _invoice_quality_score, _ocr_words_as_rows, build_voucher_rows,
    demo_invoice, load_rules, parse_invoice_text,
)


def totals(rows):
    dr = sum((r["Ledger Amount"] for r in rows if r["Ledger Amount Dr/Cr"] == "Dr"), Decimal("0"))
    cr = sum((r["Ledger Amount"] for r in rows if r["Ledger Amount Dr/Cr"] == "Cr"), Decimal("0"))
    return dr, cr


def test_demo_is_balanced_and_uses_customer_gstin_mapping():
    rows, flags = build_voucher_rows(demo_invoice(), load_rules(), voucher_day=date(2026, 11, 4))
    assert rows[0]["Voucher Date"] == "'04-11-2026"
    assert rows[0]["Voucher Number"] == "'29AL260275390"
    assert rows[0]["Reference No."] == "'29AL260275390"
    assert rows[0]["Reference Date"] == "07-09-2026"
    dr, cr = totals(rows)
    assert dr == cr == Decimal("19839.34")
    assert any(r["Ledger Name"] == "Prepaid Insurance Vehicles" and r["Ledger Amount"] == Decimal("16813.00") for r in rows)
    assert any(r["Ledger Name"] == "INPUT CGST -KAR" and r["Ledger Amount"] == Decimal("1513.17") for r in rows)
    assert any(r["Ledger Name"] == "INPUT SGST-KAR" and r["Ledger Amount"] == Decimal("1513.17") for r in rows)
    assert any(r["Ledger Name"] == "Rounding OFF" and r["Ledger Amount"] == Decimal("0.34") and r["Ledger Amount Dr/Cr"] == "Cr" for r in rows)
    assert any(r["Ledger Name"] == "Universal Sompo General Insurance Company Ltd- BNG" for r in rows)
    assert not any("INVOICE TOTAL DIFFERENCE" in f for f in flags)


def test_rounding_dr_when_invoice_total_is_higher():
    invoice = demo_invoice()
    invoice["invoice_total"] = Decimal("19839.84")
    rows, flags = build_voucher_rows(invoice, load_rules(), voucher_day=date(2026, 11, 4))
    dr, cr = totals(rows)
    assert dr == cr == Decimal("19839.84")
    assert any(r["Ledger Name"] == "Rounding OFF" and r["Ledger Amount Dr/Cr"] == "Dr" and r["Ledger Amount"] == Decimal("0.50") for r in rows)


def test_rounding_cr_when_invoice_total_is_lower():
    invoice = demo_invoice()
    invoice["invoice_total"] = Decimal("19838.84")
    rows, flags = build_voucher_rows(invoice, load_rules(), voucher_day=date(2026, 11, 4))
    dr, cr = totals(rows)
    assert dr == cr == Decimal("19839.34")
    assert any(r["Ledger Name"] == "Rounding OFF" and r["Ledger Amount Dr/Cr"] == "Cr" and r["Ledger Amount"] == Decimal("0.50") for r in rows)


def test_large_difference_adds_check_flag_and_balances():
    invoice = demo_invoice()
    invoice["invoice_total"] = Decimal("19841.34")
    rows, flags = build_voucher_rows(invoice, load_rules(), voucher_day=date(2026, 11, 4))
    dr, cr = totals(rows)
    assert dr == cr == Decimal("19841.34")
    assert any("INVOICE TOTAL DIFFERENCE" in f for f in flags)



def test_net_premium_b_is_never_used_when_total_premium_a_plus_b_is_missing():
    parsed = parse_invoice_text("""Universal Sompo General Insurance Company Limited
Net Premium (B): 10748
CGST @9%: 967.32
SGST @9%: 967.32
Gross Premium Paid: 12682.64
GSTIN of Recipient: 29ABLCS9973P1ZW
Commercial Vehicle""", "missing-ab.pdf")
    assert parsed["premium"] == Decimal("0.00")
    assert any("TOTAL PREMIUM (A+B) NOT READ" in warning for warning in parsed["_warnings"])


def test_explicitly_different_cgst_sgst_are_preserved_and_flagged():
    invoice = demo_invoice()
    invoice["cgst"] = Decimal("1513.17")
    invoice["sgst"] = Decimal("1512.17")
    rows, flags = build_voucher_rows(invoice, load_rules(), voucher_day=date(2026, 11, 4))
    assert any("CGST AND SGST AMOUNTS DIFFER" in flag for flag in flags)
    assert any(r["Ledger Name"] == "INPUT CGST -KAR" and r["Ledger Amount"] == Decimal("1513.17") for r in rows)
    assert any(r["Ledger Name"] == "INPUT SGST-KAR" and r["Ledger Amount"] == Decimal("1512.17") for r in rows)


def test_parser_reads_basic_labeled_fields():
    text = """Universal Sompo General Insurance Company Limited
Tax Invoice No: USG-INV-1122
Invoice Date: 03/11/2026
Policy No: USG-POL-9988
Name of Insured: ABC Logistics Pvt Ltd
Policy Period From: 04/11/2026
Policy Period To: 03/11/2027
GSTIN of Supplier: 27AAAAA1111A1Z5
GSTIN of Recipient: 29ABCDE1234F1Z5
Net Premium: 10,000.00
CGST @ 9%: 900.00
SGST @ 9%: 900.00
Total Invoice Amount: 11,800.00
Private Car Package Policy"""
    result = parse_invoice_text(text, "sample.pdf")
    assert result["invoice_number"] == "USG-INV-1122"
    assert result["policy_number"] == "USG-POL-9988"
    assert result["customer_gstin"] == "29ABCDE1234F1Z5"
    assert result["premium"] == Decimal("10000.00")
    assert result["cgst"] == Decimal("900.00")
    assert result["sgst"] == Decimal("900.00")
    assert result["invoice_total"] == Decimal("11800.00")


def test_motor_invoice_prefers_total_premium_a_plus_b_and_reads_both_taxes():
    # Regression fixture based on the uploaded liability/premium table. The
    # Net Liability Premium (B) is deliberately listed before Total Premium
    # (A+B); it must not be booked as the insurance expense.
    text = """Universal Sompo General Insurance Company Limited
Tax Invoice No: 29AL260275390
Invoice Date: 04/11/2026
Policy No: POL-TEST-001
Name of Insured: M/S SAMVREEDHI MOBILITY PRIVATE LIMITED GST No & State
Policy Period From: 07/09/2026
Policy Period To: 06/09/2027
GSTIN of Recipient: 29AB LCS9973P1ZW
Net Liability Premium (B): 10748
Total Premium (A+B): 16813
CGST @9%: 1513.17
SGST @9%: 1513.17
Gross Premium Paid: 19839
Commercial Vehicle"""
    result = parse_invoice_text(text, "motor-invoice.pdf")
    assert result["premium"] == Decimal("16813.00")
    assert result["cgst"] == Decimal("1513.17")
    assert result["sgst"] == Decimal("1513.17")
    assert result["invoice_total"] == Decimal("19839.00")
    assert result["insured"] == "M/S SAMVREEDHI MOBILITY PRIVATE LIMITED"


def test_motor_invoice_with_one_tax_amount_read_infers_equal_cgst_sgst_and_balances():
    text = """Universal Sompo General Insurance Company Limited
Total Premium (A+B): 16813
CGST @9%:
SGST @9%: 1513.17
GSTIN of Recipient: 29AB LCS9973P1ZW
Gross Premium Paid: 19839"""
    parsed = parse_invoice_text(text, "motor-invoice.pdf")
    assert parsed["cgst"] == Decimal("1513.17")
    assert parsed["sgst"] == Decimal("1513.17")
    assert any("CGST AMOUNT INFERRED" in w for w in parsed["_warnings"])

    invoice = {
        **parsed,
        "insurer": "Universal Sompo General Insurance Company Limited",
        "insured": "M/S SAMVREEDHI MOBILITY PRIVATE LIMITED",
        "policy_type": "Commercial Vehicle",
        "policy_number": "POL-TEST-001",
        "invoice_number": "29AL260275390",
        "invoice_date": "04-11-2026",
        "period_from": "07-09-2026",
        "period_to": "06-09-2027",
    }
    rows, flags = build_voucher_rows(invoice, load_rules(), voucher_day=date(2026, 11, 4))
    dr, cr = totals(rows)
    assert dr == cr == Decimal("19839.34")
    assert any(r["Ledger Name"] == "Prepaid Insurance Vehicles" and r["Ledger Amount"] == Decimal("16813.00") and r["Ledger Amount Dr/Cr"] == "Dr" for r in rows)
    assert any(r["Ledger Name"] == "INPUT CGST -KAR" and r["Ledger Amount"] == Decimal("1513.17") for r in rows)
    assert any(r["Ledger Name"] == "INPUT SGST-KAR" and r["Ledger Amount"] == Decimal("1513.17") for r in rows)
    assert any(r["Ledger Name"] == "Rounding OFF" and r["Ledger Amount"] == Decimal("0.34") and r["Ledger Amount Dr/Cr"] == "Cr" for r in rows)
    assert any(r["Ledger Name"] == "Universal Sompo General Insurance Company Ltd- BNG" and r["Ledger Amount"] == Decimal("19839.00") and r["Ledger Amount Dr/Cr"] == "Cr" for r in rows)
    assert not any("INVOICE TOTAL DIFFERENCE" in f for f in flags)
    assert any("CGST AMOUNT INFERRED" in f for f in flags)


def test_motor_invoice_can_infer_sgst_when_cgst_is_the_only_readable_tax():
    text = """Total Premium (A+B): 16813
CGST @9%: 1513.17
SGST @9%:
GSTIN of Recipient: 29AB LCS9973P1ZW
Gross Premium Paid: 19839"""
    parsed = parse_invoice_text(text, "motor-invoice.pdf")
    assert parsed["cgst"] == Decimal("1513.17")
    assert parsed["sgst"] == Decimal("1513.17")
    assert any("SGST AMOUNT INFERRED" in warning for warning in parsed["_warnings"])


def test_unmapped_customer_gstin_uses_fallback_gst_ledger_and_flag():
    invoice = demo_invoice()
    invoice["customer_gstin"] = "07ABCDE1234F1Z5"
    rows, flags = build_voucher_rows(invoice, load_rules(), voucher_day=date(2026, 11, 4))
    dr, cr = totals(rows)
    assert dr == cr == Decimal("19839.34")
    assert any(r["Ledger Name"] == "Input GST - CHECK" and r["Ledger Amount"] == Decimal("3026.34") for r in rows)
    assert any("GSTIN STATE 07 NOT MAPPED" in f for f in flags)
    assert any("INSURER LEDGER NOT MAPPED" in f for f in flags)


def test_total_premium_a_plus_b_beats_earlier_net_premium_row():
    # The preferred premium row is later in the OCR text, and the earlier
    # component is misleadingly labelled simply "Net Premium".
    text = """Net Premium (B): 10748
Total Premium (A+B): 16813
CGST @9%: 1513.17
SGST @9%: 1513.17
Gross Premium Paid: 19839
GSTIN of Recipient: 29ABLCS9973P1ZW"""
    parsed = parse_invoice_text(text, "priority-order.pdf")
    assert parsed["premium"] == Decimal("16813.00")
    assert parsed["cgst"] == Decimal("1513.17")
    assert parsed["sgst"] == Decimal("1513.17")
    assert parsed["customer_gstin"] == "29ABLCS9973P1ZW"


def test_ocr_split_table_rows_and_spaced_tax_labels_are_supported():
    text = """Universal Sompo General Insurance Company Limited
Name of Insured: M/S SAMVREEDHI MOBILITY PRIVATE LIMITED
Net Premium (B)
10748
Total Premium (A+B)
₹ 16,813
C G S T @ 9%
₹ 1,513.17
S G S T @ 9%
₹ 1,513.17
Gross Premium Paid
₹ 19,839
GST No & State
29ABLCS9973P1ZW"""
    parsed = parse_invoice_text(text, "split-rows.pdf")
    assert parsed["premium"] == Decimal("16813.00")
    assert parsed["cgst"] == Decimal("1513.17")
    assert parsed["sgst"] == Decimal("1513.17")
    assert parsed["invoice_total"] == Decimal("19839.00")
    assert parsed["customer_gstin"] == "29ABLCS9973P1ZW"


def test_corrected_invoice_rows_match_required_accounting_output():
    parsed = parse_invoice_text("""Universal Sompo General Insurance Company Limited
Tax Invoice No: 29AL260275390
Invoice Date: 04/11/2026
Policy No: POL-TEST-001
Name of Insured: M/S SAMVREEDHI MOBILITY PRIVATE LIMITED
Policy Period From: 07/09/2026
Policy Period To: 06/09/2027
GSTIN of Recipient: 29ABLCS9973P1ZW
Net Liability Premium (B): 10748
Total Premium (A+B): 16813
CGST @9%: 1513.17
SGST @9%: 1513.17
Gross Premium Paid: 19839
Commercial Vehicle""", "samvreeddhi-test.pdf")
    rows, flags = build_voucher_rows(parsed, load_rules(), voucher_day=date(2026, 11, 5))
    dr, cr = totals(rows)
    assert dr == cr == Decimal("19839.34")
    assert [(r["Ledger Name"], r["Ledger Amount"], r["Ledger Amount Dr/Cr"]) for r in rows] == [
        ("Prepaid Insurance Vehicles", Decimal("16813.00"), "Dr"),
        ("INPUT CGST -KAR", Decimal("1513.17"), "Dr"),
        ("INPUT SGST-KAR", Decimal("1513.17"), "Dr"),
        ("Rounding OFF", Decimal("0.34"), "Cr"),
        ("Universal Sompo General Insurance Company Ltd- BNG", Decimal("19839.00"), "Cr"),
    ]


def test_missing_cgst_is_inferred_even_when_customer_gstin_ocr_is_blank():
    text = """Net Premium (B): 10748
Total Premium (A+B): 16813
CGST @9%:
SGST @9%: 1513.17
Gross Premium Paid: 19839"""
    parsed = parse_invoice_text(text, "gstin-missed.pdf")
    assert parsed["customer_gstin"] == ""
    assert parsed["premium"] == Decimal("16813.00")
    assert parsed["cgst"] == parsed["sgst"] == Decimal("1513.17")
    assert any("CGST AMOUNT INFERRED" in warning for warning in parsed["_warnings"])


def test_customer_gstin_can_be_found_when_only_one_non_supplier_gstin_is_present():
    text = """Universal Sompo General Insurance Company Limited
Name of Insured: Example Mobility Pvt Ltd
GST No & State
29ABLCS9973P1ZW
Total Premium (A+B): 16813"""
    parsed = parse_invoice_text(text, "gstin-near-customer.pdf")
    assert parsed["customer_gstin"] == "29ABLCS9973P1ZW"



def test_ocr_common_premium_typos_still_extract_total_premium_and_gross_paid():
    # Actual Tesseract output from the uploaded invoice screenshot can read
    # "Premium" as "Premiom" or "Premiim". Both must map to the A+B amount.
    text = """Net Liability Premium (B) = 10748
Total Premiom (A+B) #16813
CGST @9% #1513.17
SGST @9% #1513.17
Gross Premiim Paid #19839"""
    parsed = parse_invoice_text(text, "ocr-typo.pdf")
    assert parsed["premium"] == Decimal("16813.00")
    assert parsed["cgst"] == Decimal("1513.17")
    assert parsed["sgst"] == Decimal("1513.17")
    assert parsed["invoice_total"] == Decimal("19839.00")


def test_accounting_quality_score_prefers_complete_ocr_over_broken_pdf_text():
    broken_embedded = """Premium before GST: 2457
Policy No:
Vehicle Identification No.
CGST @9%
SGST @9%
Invoice Total: 0
GSTIN of Recipient: 29ABLCS9973P1ZW"""
    correct_ocr = """Net Liability Premium (B) = 10748
Total Premiom (A+B) #16813
CGST @9% #1513.17
SGST @9% #1513.17
Gross Premiim Paid #19839"""
    assert _invoice_quality_score(correct_ocr) > _invoice_quality_score(broken_embedded)


def test_policy_number_never_uses_vehicle_identification_number():
    text = """Policy No:
Vehicle Identification No.
KA51AN3094"""
    assert _extract_policy_number(text) == ""


def test_policy_number_reads_policy_identifier_when_present():
    text = """Policy No.: USG-POL-9988
Vehicle Identification No.
KA51AN3094"""
    assert _extract_policy_number(text) == "USG-POL-9988"



def test_ocr_word_boxes_rebuild_left_labels_and_right_amounts_into_rows():
    # Tesseract can emit all table labels first and right-column amounts later.
    # The coordinates show that each label and amount belongs on the same row.
    data = {
        "text": ["Total", "Premium", "(A+B)", "16813", "CGST", "@9%", "1513.17", "SGST", "@9%", "1513.17", "Gross", "Premium", "Paid", "19839"],
        "conf": [95] * 14,
        "left": [10, 45, 120, 420, 10, 60, 420, 10, 60, 420, 10, 55, 130, 420],
        "top": [10, 10, 10, 10, 30, 30, 30, 50, 50, 50, 70, 70, 70, 70],
        "height": [12] * 14,
    }
    rebuilt = _ocr_words_as_rows(data)
    parsed = parse_invoice_text(rebuilt, "ocr-boxes.pdf")
    assert parsed["premium"] == Decimal("16813.00")
    assert parsed["cgst"] == Decimal("1513.17")
    assert parsed["sgst"] == Decimal("1513.17")
    assert parsed["invoice_total"] == Decimal("19839.00")


def test_four_uploaded_universal_sompo_scan_layouts_extract_correct_accounting_fields():
    """Regression fixture based on OCR crops from the four uploaded one-page PDFs."""
    cases = [
        ("olicy.No", "2314/86957387/00/000", "29AL260275369", "GST No & State"),
        ("Policy No", "2314/86957518/00/000", "29AL260275390", "Porno"),
        ("olicy.No", "2314/86957439/00/000", "29AL260275377", "GST No & State"),
        ("olicy.No", "2314/86957429/00/000", "29AL260275376", "GST No & State"),
    ]
    for policy_label, policy_number, invoice_number, gst_label in cases:
        text = f"""Universal Sompo General Insurance Company Limited
Package Policy (Commercial Vehicle - Passenger Carrying)
Period of Insurance 07-SEP-2026 20:56 to 06-SEP-2027 23:59
{policy_label} {policy_number} Vehicle Identification No. MA3JMTBISTGF44990
Policy Issued On 07-Sep-2026 20:56
Insured Name M/S SAMVREEDDHI MOBILITY PRIVATE LIMITED {gst_label} 29ABLCS9973P1ZW | Karnataka
Invoice No {invoice_number}
Insured State & Code Karnataka - 29 Place of Supply Karnataka GSTIN of Customer 29ABLCS9973P1ZW
OWN DAMAGE SECTION (A) LIABILITY SECTION (B)
Basic Premium 2457 | Net Liability Premium (B) 10748
Total Premium (A+B) 16813
CGST @9% 1513.17
SGST @9% 1513.17
Gross Premium Paid 19839"""
        parsed = parse_invoice_text(text, invoice_number + ".pdf")
        assert parsed["insurer"] == "Universal Sompo General Insurance Company Limited"
        assert parsed["insured"] == "M/S SAMVREEDDHI MOBILITY PRIVATE LIMITED"
        assert parsed["policy_number"] == policy_number
        assert parsed["invoice_number"] == invoice_number
        assert parsed["customer_gstin"] == "29ABLCS9973P1ZW"
        assert parsed["invoice_date"] == "07-09-2026"
        assert parsed["period_from"] == "07-09-2026"
        assert parsed["period_to"] == "06-09-2027"
        assert parsed["premium"] == Decimal("16813.00")
        assert parsed["cgst"] == Decimal("1513.17")
        assert parsed["sgst"] == Decimal("1513.17")
        assert parsed["invoice_total"] == Decimal("19839.00")
        assert any("POLICY ISSUE DATE USED AS REFERENCE DATE" in item for item in parsed["_warnings"])
        rows, flags = build_voucher_rows(parsed, load_rules(), voucher_day=date(2026, 11, 5))
        dr, cr = totals(rows)
        assert dr == cr == Decimal("19839.34")
        assert any(row["Ledger Name"] == "Rounding OFF" and row["Ledger Amount"] == Decimal("0.34") and row["Ledger Amount Dr/Cr"] == "Cr" for row in rows)


def test_hdfc_mumbai_focused_table_reads_sgst_and_correct_gross_premium():
    # OCR output from the focused right-hand table crop of the HDFC ERGO
    # Maharashtra motor-policy schedule. Full-page OCR had read SGST as
    # "scst" and Gross Premium Paid 19596 as 19506.
    focused_table = """LIABILITY SECTION (B)
Net Liability Premium (B) = 10923
Total Premium (A+B) = 16607
CGST @9% = 1494.63
SGST @9% = 1494.63
Gross Premium Paid = 19596
"""
    metadata = """HDFC ERGO General Insurance Company Limited
Insured Name M/S SAMVREEDDHI MOBILITY PRIVATE LIMITED
Policy No 2314208758482500000
Policy Issued On 10-Jul-2026 22:47
Invoice No 2087584825000
GST No & State 27ABLCS9973P1Z0 Maharashtra
Period of Insurance 10-JUL-2026 to 09-JUL-2027
"""
    parsed = parse_invoice_text(focused_table + "\n" + metadata, "hdfc-mumbai.pdf")
    assert parsed["premium"] == Decimal("16607.00")
    assert parsed["cgst"] == Decimal("1494.63")
    assert parsed["sgst"] == Decimal("1494.63")
    assert parsed["invoice_total"] == Decimal("19596.00")
    assert parsed["customer_gstin"] == "27ABLCS9973P1Z0"

    rows, flags = build_voucher_rows(parsed, load_rules(), voucher_day=date(2026, 11, 9))
    dr, cr = totals(rows)
    assert dr == cr == Decimal("19596.26")
    assert any(r["Ledger Name"] == "Prepaid Insurance Vehicles" and r["Ledger Amount"] == Decimal("16607.00") for r in rows)
    assert any(r["Ledger Name"] == "Input CGST_MUM" and r["Ledger Amount"] == Decimal("1494.63") for r in rows)
    assert any(r["Ledger Name"] == "Input SGST_MUM" and r["Ledger Amount"] == Decimal("1494.63") for r in rows)
    assert any(r["Ledger Name"] == "Rounding OFF" and r["Ledger Amount"] == Decimal("0.26") and r["Ledger Amount Dr/Cr"] == "Cr" for r in rows)
    assert any(r["Ledger Name"] == "HDFC ERGO GENERAL INSURANCE COMPANY LIMITED" and r["Ledger Amount"] == Decimal("19596.00") and r["Ledger Amount Dr/Cr"] == "Cr" for r in rows)


def _synthetic_vendor_voucher(vendor_ledger, voucher_number):
    """Create one balanced four-line voucher for Excel vendor-tab tests."""
    common = {
        "Voucher Date": "'09-11-2026",
        "Voucher Type Name": "Journal",
        "Voucher Number": f"'{voucher_number}",
        "Voucher Narration": f"Being test premium for {vendor_ledger}, Inv No {voucher_number}",
        "Reference No.": f"'{voucher_number}",
        "Reference Date": "09-11-2026",
    }
    specs = [
        ("Prepaid Insurance Vehicles", Decimal("100.00"), "Dr"),
        ("Input GST - CHECK", Decimal("18.00"), "Dr"),
        ("Rounding OFF", Decimal("1.00"), "Cr"),
        (vendor_ledger, Decimal("117.00"), "Cr"),
    ]
    return [
        {**common, "Ledger Name": ledger, "Ledger Amount": amount, "Ledger Amount Dr/Cr": side}
        for ledger, amount, side in specs
    ]


def test_vendor_wise_xlsx_has_master_index_and_dynamic_vendor_tabs(tmp_path):
    from openpyxl import load_workbook
    from invoice_engine import write_xlsx

    vendors = [
        "Universal Sompo General Insurance Company Ltd- BNG",
        "Universal Sompo General Insurance Company Ltd-HYD",
        "Universal Sompo General Insurance Company -Mum",
        "SBI General Insurance Company Limited",
        "HDFC ERGO GENERAL INSURANCE COMPANY LIMITED",
        "A Newly Added General Insurance Provider Private Limited",
    ]
    rows = []
    for index, vendor in enumerate(vendors, start=1):
        rows.extend(_synthetic_vendor_voucher(vendor, f"INV-{index:03d}"))

    output = tmp_path / "vendor-wise.xlsx"
    write_xlsx(rows, output)
    workbook = load_workbook(output, data_only=True)

    assert workbook.sheetnames[:2] == ["All Vendors", "Vendor Index"]
    assert workbook["All Vendors"].max_row == 1 + len(rows)
    index = workbook["Vendor Index"]
    index_rows = list(index.iter_rows(min_row=2, values_only=True))
    indexed_names = {row[0] for row in index_rows}
    assert indexed_names == set(vendors)

    tabs_by_vendor = {row[0]: row[1] for row in index_rows}
    assert tabs_by_vendor[vendors[0]] == "Universal Sompo - BNG"
    assert tabs_by_vendor[vendors[1]] == "Universal Sompo - HYD"
    assert tabs_by_vendor[vendors[2]] == "Universal Sompo - Mum"
    assert tabs_by_vendor[vendors[3]] == "SBI General"
    assert tabs_by_vendor[vendors[4]] == "HDFC ERGO"
    assert tabs_by_vendor[vendors[5]] in workbook.sheetnames

    for vendor in vendors:
        sheet = workbook[tabs_by_vendor[vendor]]
        assert sheet.max_row == 5  # header + all four lines of one complete voucher
        assert [sheet.cell(row, 4).value for row in range(2, 6)].count(vendor) == 1
        assert all(len(name) <= 31 for name in workbook.sheetnames)

    assert all(row[2] == 1 and row[3] == 4 for row in index_rows)
