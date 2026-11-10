# Insurance Invoice → Journal Voucher Tool

A Streamlit web app and standalone Python folder tool that read multiple insurance invoices and create journal voucher rows in CSV and Excel.

**No AI model, Anthropic/OpenAI API key, token balance, or paid AI service is used.** The program uses normal Python accounting rules, PyMuPDF for PDF text extraction, and local Tesseract OCR for scanned pages and image files. Invoice data is not sent to an AI provider.

> Important: OCR and rule-based extraction are not infallible. Review the extracted values and every `[CHECK ...]` note before posting vouchers into accounting software.

## Supported files

- Text/searchable PDFs
- Scanned/image-only PDFs (each page is rendered and OCR'd locally)
- JPG / JPEG / PNG
- TIFF / BMP

## What the app does

1. Upload several invoices in one batch.
2. Show a per-file processing status and OCR/extraction log.
3. Extract insurer, insured, policy type/number, invoice number/date, policy period, customer GSTIN, premium, CGST, SGST, IGST and invoice total.
4. Let you correct extracted invoice fields before generating journal rows.
5. Generate a copyable tab-separated journal with headers, CSV, and Excel output.
6. Let you edit voucher type, expense/fallback GST/rounding ledgers, GST state mappings and insurer ledger mappings in the app.

## Accounting rules

The CSV/XLSX columns are output in this exact order:

`Voucher Date | Voucher Type Name | Voucher Number | Ledger Name | Ledger Amount | Ledger Amount Dr/Cr | Voucher Narration | Reference No. | Reference Date`

- Voucher Date: current India-local date, `dd-MM-yyyy`.
- Reference Date: invoice date, converted to `dd-MM-yyyy`. If the invoice has no separate invoice-date field, the parser uses the printed `Policy Issued On` date and adds a `[CHECK ...]` note.
- Voucher Type Name: `Journal` by default; editable.
- Voucher Number: invoice number, or policy number if invoice number is missing.
- Reference No.: invoice number.
- Voucher Date, Voucher Number and Reference No. have a leading apostrophe for text handling in Excel.
- Prepaid Insurance Vehicles (Dr): **Total Premium (A+B)** / the explicit pre-GST total. The parser must not use `Net Liability Premium (B)` or `Net Premium (B)` as the full premium. If a liability-component layout is detected but the A+B total cannot be read, the premium is left at zero and a `[CHECK ...]` note is added rather than booking the wrong component.
- Input CGST and SGST (Dr): selected using the first two digits of the **customer's GSTIN**, not the insurer's GSTIN.
- Universal Sompo supplier ledger: selected using the customer GSTIN state code.
- SBI General and HDFC ERGO supplier ledgers: fixed defaults from `rules.json`.
- Insurer (Cr): full invoice total.
- Rounding OFF: the difference between invoice total and (premium + extracted GST). Dr when invoice total is higher; Cr when it is lower. A difference of ₹1.00 or more adds a `[CHECK ...]` flag. If both CGST and SGST are explicitly read, a difference is treated as rounding, not silently reclassified as additional GST.
- If a GSTIN state or insurer mapping is missing, the fallback mapping/raw insurer ledger is used and a `[CHECK ...]` note is added.
- When both CGST and SGST labels are present, no IGST is shown, and OCR reads only one tax amount, the missing counterpart is inferred equal to the readable amount and flagged for review. If both explicit amounts differ, neither is overwritten; a check flag is added.

### Default state mappings

| Customer GSTIN starts with | CGST ledger | SGST ledger |
|---|---|---|
| 29 | `INPUT CGST -KAR` | `INPUT SGST-KAR` |
| 27 | `Input CGST_MUM` | `Input SGST_MUM` |
| 36 | `Input CGST-HYD` | `Input SGST-HYD` |

### Default insurer mappings

| Insurer | Ledger |
|---|---|
| Universal Sompo, GSTIN 29 | `Universal Sompo General Insurance Company Ltd- BNG` |
| Universal Sompo, GSTIN 27 | `Universal Sompo General Insurance Company -Mum` |
| Universal Sompo, GSTIN 36 | `Universal Sompo General Insurance Company Ltd-HYD` |
| SBI General | `SBI General Insurance Company Limited` |
| HDFC ERGO | `HDFC ERGO GENERAL INSURANCE COMPANY LIMITED` |

## 1. Run locally in VS Code (Windows)

Use **Python 3.12** (recommended for this project).

### Step A — Install Tesseract OCR

Tesseract is required only for scanned PDFs and image files. Text PDFs can be read without OCR if their text layer is usable.

1. Install the Windows Tesseract OCR application.
2. Add the folder containing `tesseract.exe` to Windows `PATH`.
3. Close and reopen VS Code.
4. Check it in the VS Code terminal:

```powershell
tesseract --version
```

If it is installed but not on PATH, set the path for the current terminal before launching the app:

```powershell
$env:TESSERACT_CMD = "C:\Program Files\Tesseract-OCR\tesseract.exe"
```

### Step B — Open the project and install packages

Extract the ZIP, open the `Insurance_Journal_Tool` folder in VS Code, then run:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

If PowerShell blocks activation, run this in the same terminal:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned
.\.venv\Scripts\Activate.ps1
```

### Step C — Start the Streamlit app

```powershell
python -m streamlit run app.py
```

Open the local URL printed in the terminal, normally `http://localhost:8501`.

## 2. Run the Python folder tool (`journal_tool.py`)

Create an `invoices` folder in the project and put your PDFs/JPG/PNG files inside it:

```text
Insurance_Journal_Tool/
  invoices/
    invoice_1.pdf
    invoice_2.jpg
    invoice_3.png
```

Run:

```powershell
python journal_tool.py --input .\invoices --output .\output
```

The program writes:

```text
output/
  journal.csv
  journal.xlsx
```


### Vendor-wise Excel export

The Excel workbook contains:

- **All Vendors**: every journal line from the batch.
- **Vendor Index**: the exact Tally vendor ledger name, its Excel sheet tab name, voucher count and journal row count.
- **One vendor sheet per Tally ledger name**: every complete voucher is kept together on the vendor sheet, including the expense debit, GST debits, rounding entry and insurer credit.

The vendor is identified from the main insurer credit ledger generated by the invoice-to-Tally mapping in `rules.json`. The five current mappings are already configured: Universal Sompo BNG, HYD and Mumbai, SBI General, and HDFC ERGO. The workbook does not hard-code a closed vendor list for grouping: when you add a new insurer mapping in `rules.json` and its Tally ledger appears on generated vouchers, a new vendor sheet is created automatically.

**Excel worksheet-name limit:** Excel allows up to 31 characters per sheet tab. Therefore, the five current long Tally names use short, recognizable tab names (for example, `Universal Sompo - BNG`). The Vendor Index maps each tab back to the exact full Tally ledger name, and the journal rows retain the full ledger name in the `Ledger Name` column. Future vendor tabs are created from the new Tally name, safely shortened if necessary.

The CSV and Copy Journal outputs remain combined across all vendors. Only the Excel download is split into vendor tabs.

### Demo mode (no invoices needed)

```powershell
python journal_tool.py --demo --output .\demo_output
```

The demo uses the verified motor-insurance arithmetic from the sample layout: Total Premium (A+B) ₹16,813.00, CGST ₹1,513.17, SGST ₹1,513.17 and invoice total ₹19,839.00. It should create a ₹0.34 Cr entry in Rounding OFF and balance the voucher.

## 3. Where to edit the rules

### In the Streamlit app

At the top under **Ledger names and mapping rules**, edit the voucher type, expense ledger, fallback GST ledger, rounding ledger, customer-GSTIN state mappings and insurer mappings. Download the updated `rules.json` from the app so you can reuse it on another device or upload it after a Streamlit Cloud restart.

### In VS Code

Edit `rules.json` directly. It is the shared defaults file used by both `app.py` and `journal_tool.py`.

For example, add a state mapping under `state_gst_mappings`:

```json
"33": {"cgst": "Input CGST-CHN", "sgst": "Input SGST-CHN"}
```

Keep your `rules.json` in your GitHub repository if these mappings are safe to share with your project team. Do not commit confidential invoice PDFs or generated company vouchers to a public repository.

## 4. Run the automated tests

```powershell
python -m pytest -q
```

The tests cover A+B premium priority, refusal to use a missing A+B in place of Net Premium (B), CGST/SGST extraction and equality inference, GSTIN-based ledger mapping, rounding direction, large-difference flags, date formatting and balanced vouchers.

## 5. Deploy with GitHub + Streamlit Community Cloud

1. Create a new GitHub repository, for example `insurance-journal-voucher-tool`.
2. Push the contents of this project folder. Do **not** push `.venv`, local invoices, `.env`, or generated company output files.
3. Open Streamlit Community Cloud and create a new app.
4. Select your GitHub repository, branch (usually `main`), and set the main file path to `app.py`.
5. Deploy. `requirements.txt` installs Python libraries and `packages.txt` installs the system Tesseract OCR package on the hosted Linux runtime.
6. Open the Streamlit URL from any device and sign in to the hosting account as required.

**No Anthropic/OpenAI key and no `.env` file are needed.** There is no token-based runtime dependency.

### Important deployment/privacy note

When used locally, invoice files are processed on your computer. When deployed to Streamlit Cloud, uploaded files are sent to the hosted Streamlit app runtime for processing (not to an AI provider). The app does not add a database or permanent invoice storage. For confidential company invoices, follow your organisation's hosting and data-handling rules; keep the app private or run it locally/on an approved server.

## 6. How extraction works / known limitations

- Searchable PDFs: read the embedded text first; a local OCR pass is used for selected incomplete invoice tables.
- Scanned PDFs: render each PDF page to an image and run local Tesseract OCR. The Universal Sompo motor schedule uses focused OCR crops for policy details and its premium table. HDFC ERGO invoices with customer GSTIN state code 27 (Maharashtra/Mumbai) use a dedicated high-resolution crop of the right-hand premium table; this avoids full-page OCR misreading SGST as `scst` or Gross Premium Paid `19596` as `19506`. The HDFC-specific override is restricted to that insurer/state combination, so Karnataka and Hyderabad parsing paths are unchanged.
- Images: run local Tesseract OCR after contrast/size preprocessing.
- Accounting calculations: Python `Decimal` values, not floating-point arithmetic.

No rule-based parser can guarantee perfect extraction from every insurer layout. OCR may misread an amount, date, or GSTIN. The review form is intentional: verify each invoice before generating the journal, especially when a `[CHECK ...]` note appears.
