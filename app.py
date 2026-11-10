from __future__ import annotations

import base64
import copy
import hashlib
import io
import json
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from invoice_engine import (
    COLUMNS,
    build_voucher_rows,
    load_rules,
    money,
    normalize_rules,
    parse_invoice_file,
    rows_to_plain,
    save_rules,
    write_csv,
    write_xlsx,
)

st.set_page_config(page_title="Insurance Journal Voucher Tool", page_icon="📄", layout="wide")

PARSER_BUILD = "2026-11-09.2"
st.title("Insurance Invoice → Journal Voucher Tool")
st.caption(f"Local OCR + editable accounting rules. No AI model, API key, token usage, or invoice data sent to an AI service. · Parser build {PARSER_BUILD}")

# Clear cached invoice extraction when the parser is upgraded. Without this,
# Streamlit can keep old Premium/CGST/SGST widget values in session_state even
# after invoice_engine.py has been replaced and the app rerun.
if st.session_state.get("_parser_build") != PARSER_BUILD:
    st.session_state["invoice_batch"] = {}
    st.session_state.pop("generated_rows", None)
    st.session_state.pop("generated_check_summary", None)
    st.session_state["_parser_build"] = PARSER_BUILD


def _init_rule_state() -> None:
    base = load_rules()
    defaults = {
        "rule_voucher_type": base["voucher_type"],
        "rule_expense": base["ledgers"]["expense"],
        "rule_fallback_gst": base["ledgers"]["fallback_gst"],
        "rule_rounding": base["ledgers"]["rounding"],
    }
    for code, mapping in base["state_gst_mappings"].items():
        defaults[f"rule_state_{code}_cgst"] = mapping.get("cgst", "")
        defaults[f"rule_state_{code}_sgst"] = mapping.get("sgst", "")
    for insurer_key, mapping in base["insurer_mappings"].items():
        for field, value in mapping.items():
            if field == "name_match" or field == "default" or str(field).isdigit():
                defaults[f"rule_insurer_{insurer_key}_{field}"] = value
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)


def _apply_uploaded_rules(upload) -> None:
    if upload is None:
        return
    upload_bytes = upload.getvalue()
    digest = hashlib.sha256(upload_bytes).hexdigest()
    if st.session_state.get("_rules_upload_digest") == digest:
        return
    try:
        uploaded = normalize_rules(json.loads(upload_bytes.decode("utf-8")))
        state = uploaded["state_gst_mappings"]
        for key, value in {
            "rule_voucher_type": uploaded["voucher_type"],
            "rule_expense": uploaded["ledgers"]["expense"],
            "rule_fallback_gst": uploaded["ledgers"]["fallback_gst"],
            "rule_rounding": uploaded["ledgers"]["rounding"],
        }.items():
            st.session_state[key] = value
        for code in ("29", "27", "36"):
            m = state.get(code, {})
            st.session_state[f"rule_state_{code}_cgst"] = m.get("cgst", "")
            st.session_state[f"rule_state_{code}_sgst"] = m.get("sgst", "")
        for insurer_key, mapping in uploaded["insurer_mappings"].items():
            for field, value in mapping.items():
                st.session_state[f"rule_insurer_{insurer_key}_{field}"] = value
        st.session_state["_rules_upload_digest"] = digest
        st.success("Rules loaded into this session.")
    except Exception as exc:
        st.error(f"Could not load rules JSON: {exc}")


def _current_rules() -> dict:
    rules = copy.deepcopy(load_rules())
    rules["voucher_type"] = st.session_state["rule_voucher_type"]
    rules["ledgers"] = {
        "expense": st.session_state["rule_expense"],
        "fallback_gst": st.session_state["rule_fallback_gst"],
        "rounding": st.session_state["rule_rounding"],
    }
    for code in ("29", "27", "36"):
        rules["state_gst_mappings"][code] = {
            "cgst": st.session_state[f"rule_state_{code}_cgst"],
            "sgst": st.session_state[f"rule_state_{code}_sgst"],
        }
    for key, mapping in rules["insurer_mappings"].items():
        for field in list(mapping.keys()):
            session_key = f"rule_insurer_{key}_{field}"
            if session_key in st.session_state:
                mapping[field] = st.session_state[session_key]
    return normalize_rules(rules)


_init_rule_state()
with st.expander("Ledger names and mapping rules", expanded=True):
    uploaded_rules = st.file_uploader("Load a previously saved rules.json (optional)", type=["json"], key="rules_json_upload")
    _apply_uploaded_rules(uploaded_rules)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.text_input("Voucher type name", key="rule_voucher_type")
    with c2:
        st.text_input("Expense ledger", key="rule_expense")
    with c3:
        st.text_input("Fallback GST ledger", key="rule_fallback_gst")
    with c4:
        st.text_input("Rounding ledger", key="rule_rounding")

    st.markdown("**Customer GSTIN state → input GST ledgers**")
    state_cols = st.columns(3)
    state_labels = {"29": "29 — Karnataka", "27": "27 — Maharashtra", "36": "36 — Telangana"}
    for col, code in zip(state_cols, ("29", "27", "36")):
        with col:
            st.caption(state_labels[code])
            st.text_input("CGST ledger", key=f"rule_state_{code}_cgst")
            st.text_input("SGST ledger", key=f"rule_state_{code}_sgst")

    with st.expander("Insurer ledger mappings", expanded=False):
        u = st.columns(3)
        with u[0]:
            st.text_input("Universal Sompo — state 29", key="rule_insurer_universal_sompo_29")
        with u[1]:
            st.text_input("Universal Sompo — state 27", key="rule_insurer_universal_sompo_27")
        with u[2]:
            st.text_input("Universal Sompo — state 36", key="rule_insurer_universal_sompo_36")
        s = st.columns(2)
        with s[0]:
            st.text_input("SBI General ledger", key="rule_insurer_sbi_general_default")
        with s[1]:
            st.text_input("HDFC ERGO ledger", key="rule_insurer_hdfc_ergo_default")

    current_rules_bytes = json.dumps(_current_rules(), indent=2, ensure_ascii=False).encode("utf-8")
    st.download_button(
        "Download current rules.json",
        data=current_rules_bytes,
        file_name="rules.json",
        mime="application/json",
        help="Save your edited ledger mappings so you can upload them in a later session or on another device.",
    )

st.divider()
st.subheader("1. Upload invoices")
st.write("Upload several text PDFs, scanned PDFs, JPGs, or PNGs together. Scanned pages are rendered and OCR'd locally.")
upload_version = st.session_state.setdefault("invoice_upload_version", 0)
uploads = st.file_uploader(
    "Choose invoice files",
    type=["pdf", "jpg", "jpeg", "png", "tif", "tiff", "bmp"],
    accept_multiple_files=True,
    key=f"invoice_uploads_{upload_version}",
)

if "invoice_batch" not in st.session_state:
    st.session_state["invoice_batch"] = {}

if uploads:
    new_uploads = []
    for uploaded in uploads:
        raw = uploaded.getvalue()
        file_key = hashlib.sha256((PARSER_BUILD + ":" + uploaded.name).encode("utf-8") + raw).hexdigest()[:16]
        if file_key not in st.session_state["invoice_batch"]:
            new_uploads.append((file_key, uploaded.name, raw))
    if new_uploads:
        st.session_state.pop("generated_rows", None)
        st.session_state.pop("generated_check_summary", None)
        progress = st.progress(0, text="Reading invoice files...")
        for index, (file_key, file_name, raw) in enumerate(new_uploads, start=1):
            with st.status(f"Processing {file_name}", expanded=True) as status:
                try:
                    parsed, method, pages, log = parse_invoice_file(file_name, raw)
                    parsed["_file_key"] = file_key
                    st.session_state["invoice_batch"][file_key] = {
                        "name": file_name,
                        "data": parsed,
                        "method": method,
                        "pages": pages,
                        "log": log,
                        "error": "",
                    }
                    st.write(f"Extraction: {method} · Pages: {pages}")
                    for line in log:
                        st.write(f"• {line}")
                    status.update(label=f"Processed {file_name}", state="complete")
                except Exception as exc:
                    st.session_state["invoice_batch"][file_key] = {
                        "name": file_name,
                        "data": None,
                        "method": "Failed",
                        "pages": 0,
                        "log": [],
                        "error": str(exc),
                    }
                    st.error(str(exc))
                    status.update(label=f"Could not process {file_name}", state="error")
            progress.progress(index / len(new_uploads), text=f"Processed {index} of {len(new_uploads)} new file(s)")
        progress.empty()

batch = st.session_state["invoice_batch"]
if batch:
    st.subheader("2. Review extracted invoice details")
    log_rows = []
    for file_key, entry in batch.items():
        d = entry["data"]
        log_rows.append({
            "File": entry["name"],
            "Status": "Needs review" if not entry["error"] else "Failed",
            "Read method": entry["method"],
            "Pages": entry["pages"],
            "Message": entry["error"] or "; ".join((d or {}).get("_warnings", [])),
        })
    st.dataframe(pd.DataFrame(log_rows), use_container_width=True, hide_index=True)

    for file_key, entry in batch.items():
        if entry["error"] or entry["data"] is None:
            with st.expander(f"{entry['name']} — could not be read", expanded=True):
                st.error(entry["error"])
            continue
        d = entry["data"]
        with st.expander(f"{entry['name']} · {entry['method']}", expanded=True):
            st.caption("Check the OCR results. You can edit every extracted field before generating the journal.")
            with st.expander("Processing log", expanded=False):
                for log_line in entry.get("log", []):
                    st.write(f"• {log_line}")
            left, right = st.columns(2)
            with left:
                st.text_input("Insurer", value=str(d.get("insurer", "")), key=f"{file_key}_insurer")
                st.text_input("Insurer ledger override (blank = use mapping)", value=str(d.get("insurer_ledger", "")), key=f"{file_key}_insurer_ledger")
                st.text_input("Insured / customer name", value=str(d.get("insured", "")), key=f"{file_key}_insured")
                st.text_input("Policy type", value=str(d.get("policy_type", "")), key=f"{file_key}_policy_type")
                st.text_input("Policy number", value=str(d.get("policy_number", "")), key=f"{file_key}_policy_number")
                st.text_input("Invoice number", value=str(d.get("invoice_number", "")), key=f"{file_key}_invoice_number")
            with right:
                st.text_input("Invoice date (dd-MM-yyyy)", value=str(d.get("invoice_date", "")), key=f"{file_key}_invoice_date")
                st.text_input("Policy period from (dd-MM-yyyy)", value=str(d.get("period_from", "")), key=f"{file_key}_period_from")
                st.text_input("Policy period to (dd-MM-yyyy)", value=str(d.get("period_to", "")), key=f"{file_key}_period_to")
                st.text_input("Customer GSTIN (not insurer GSTIN)", value=str(d.get("customer_gstin", "")), key=f"{file_key}_customer_gstin")
                st.number_input("Total Premium (A+B) / Premium before GST", min_value=0.0, value=float(money(d.get("premium"))), step=0.01, format="%.2f", key=f"{file_key}_premium", help="For motor policies with Own Damage (A) and Liability (B), use Total Premium (A+B), not Net Liability Premium (B).")
                st.number_input("Invoice total", min_value=0.0, value=float(money(d.get("invoice_total"))), step=0.01, format="%.2f", key=f"{file_key}_invoice_total")
            tax1, tax2, tax3, tax4 = st.columns(4)
            with tax1:
                st.number_input("CGST", min_value=0.0, value=float(money(d.get("cgst"))), step=0.01, format="%.2f", key=f"{file_key}_cgst")
            with tax2:
                st.number_input("SGST", min_value=0.0, value=float(money(d.get("sgst"))), step=0.01, format="%.2f", key=f"{file_key}_sgst")
            with tax3:
                st.number_input("IGST", min_value=0.0, value=float(money(d.get("igst"))), step=0.01, format="%.2f", key=f"{file_key}_igst")
            with tax4:
                st.number_input("Total GST (optional)", min_value=0.0, value=float(money(d.get("total_gst"))), step=0.01, format="%.2f", key=f"{file_key}_total_gst")
            with st.expander("Show extracted text / OCR output"):
                st.text_area("Invoice text", value=str(d.get("_text", "")), height=220, key=f"{file_key}_raw_text", disabled=True)

    st.divider()
    st.subheader("3. Generate journal")
    st.caption("The voucher date is set to the date the journal is generated. Invoice dates are used as Reference Date.")
    if st.button("Generate journal vouchers", type="primary", use_container_width=True):
        rules = _current_rules()
        all_rows = []
        check_summary = []
        for file_key, entry in batch.items():
            if entry["error"] or entry["data"] is None:
                continue
            d = copy.deepcopy(entry["data"])
            for field in ("insurer", "insurer_ledger", "insured", "policy_type", "policy_number", "invoice_number", "invoice_date", "period_from", "period_to", "customer_gstin"):
                d[field] = st.session_state.get(f"{file_key}_{field}", d.get(field, ""))
            for field in ("premium", "invoice_total", "cgst", "sgst", "igst", "total_gst"):
                d[field] = money(st.session_state.get(f"{file_key}_{field}", d.get(field, 0)))
            rows, flags = build_voucher_rows(d, rules)
            all_rows.extend(rows)
            check_summary.append({"File": entry["name"], "Voucher lines": len(rows), "Checks": "; ".join(flags) if flags else "No checks"})
        st.session_state["generated_rows"] = all_rows
        st.session_state["generated_check_summary"] = check_summary
        st.session_state["generated_rules"] = rules

if "generated_rows" in st.session_state:
    rows = st.session_state["generated_rows"]
    if rows:
        st.subheader("4. Journal output")
        st.dataframe(pd.DataFrame([
            {**row, "Ledger Amount": float(money(row["Ledger Amount"]))}
            for row in rows
        ], columns=COLUMNS), use_container_width=True, hide_index=True)

        check_summary = st.session_state.get("generated_check_summary", [])
        if any(item["Checks"] != "No checks" for item in check_summary):
            with st.expander("Files requiring attention", expanded=True):
                st.dataframe(pd.DataFrame(check_summary), use_container_width=True, hide_index=True)
        else:
            st.success("All invoices processed without rule checks.")

        plain_rows = rows_to_plain(rows)
        csv_buffer = io.StringIO(newline="")
        import csv
        writer = csv.DictWriter(csv_buffer, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(plain_rows)
        csv_bytes = csv_buffer.getvalue().encode("utf-8-sig")

        tsv_buffer = io.StringIO(newline="")
        tsv_writer = csv.DictWriter(tsv_buffer, fieldnames=COLUMNS, delimiter="\t", lineterminator="\n")
        tsv_writer.writeheader()
        tsv_writer.writerows(plain_rows)
        tsv = tsv_buffer.getvalue()
        encoded_tsv = base64.b64encode(tsv.encode("utf-8")).decode("ascii")
        copy_html = f"""
        <div style="font-family:Arial,sans-serif;display:flex;align-items:center;gap:12px;">
          <button id="copyJournal" style="background:#1769aa;color:white;border:0;padding:10px 16px;border-radius:7px;font-size:14px;cursor:pointer;">
            Copy journal (tab-separated)
          </button>
          <span id="copyStatus" style="font-size:13px;color:inherit;"></span>
        </div>
        <script>
        const b64 = "{encoded_tsv}";
        const raw = atob(b64);
        const bytes = Uint8Array.from(raw, c => c.charCodeAt(0));
        const content = new TextDecoder("utf-8").decode(bytes);
        document.getElementById("copyJournal").addEventListener("click", async () => {{
          const status = document.getElementById("copyStatus");
          try {{
            await navigator.clipboard.writeText(content);
            status.textContent = "Copied. Paste directly into Excel.";
          }} catch (e) {{
            const area = document.createElement("textarea");
            area.value = content; document.body.appendChild(area); area.select();
            const ok = document.execCommand("copy"); area.remove();
            status.textContent = ok ? "Copied. Paste directly into Excel." : "Copy blocked by browser; use the TSV preview below.";
          }}
        }});
        </script>
        """
        components.html(copy_html, height=55)
        st.caption(
            "Vendor-wise Excel includes an All Vendors master sheet, a Vendor Index "
            "showing exact Tally ledger names, and a separate complete voucher sheet "
            "for each vendor. New vendor ledger names automatically get new sheets."
        )
        col_a, col_b = st.columns(2)
        with col_a:
            st.download_button("Download CSV (all vendors)", data=csv_bytes, file_name="journal.csv", mime="text/csv", use_container_width=True)
        with col_b:
            with tempfile.TemporaryDirectory() as temp_dir:
                xlsx_path = Path(temp_dir) / "insurance_journal_vendor_wise.xlsx"
                write_xlsx(rows, xlsx_path)
                xlsx_bytes = xlsx_path.read_bytes()
            st.download_button("Download Vendor-wise Excel", data=xlsx_bytes, file_name="insurance_journal_vendor_wise.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with st.expander("Show tab-separated copy preview"):
            st.code(tsv, language="text")
    else:
        st.warning("No journal lines were produced. Review the extraction errors and required amounts.")

with st.sidebar:
    st.header("About this tool")
    st.write("• Text PDF: PyMuPDF\n\n• Scanned PDF and images: local Tesseract OCR\n\n• Journal rules: Python + editable JSON\n\n• No external AI or paid token usage")
    st.info("OCR results should be checked before posting to your accounting system.")
    if st.button("Clear uploaded batch and results", use_container_width=True):
        st.session_state["invoice_batch"] = {}
        st.session_state["invoice_upload_version"] = st.session_state.get("invoice_upload_version", 0) + 1
        st.session_state.pop("generated_rows", None)
        st.session_state.pop("generated_check_summary", None)
        st.rerun()
