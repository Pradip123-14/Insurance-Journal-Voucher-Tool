"""Rule-based insurance invoice extraction and journal voucher generation.

No generative AI or external API is used. PDFs are read with PyMuPDF; scanned
pages and image files are read with the local Tesseract OCR executable.
"""
from __future__ import annotations

import csv
import json
import os
import re
from datetime import date, datetime, timezone, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable

import fitz  # PyMuPDF
import pytesseract
from PIL import Image, ImageOps, ImageEnhance, ImageFilter
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

COLUMNS = [
    "Voucher Date",
    "Voucher Type Name",
    "Voucher Number",
    "Ledger Name",
    "Ledger Amount",
    "Ledger Amount Dr/Cr",
    "Voucher Narration",
    "Reference No.",
    "Reference Date",
]

SUPPORTED_EXTENSIONS = {".pdf", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}
CENT = Decimal("0.01")
ZERO = Decimal("0.00")


def money(value: Any) -> Decimal:
    """Convert a UI/parser value to a 2-decimal Decimal safely."""
    if value is None or value == "":
        return ZERO
    if isinstance(value, Decimal):
        return value.quantize(CENT, rounding=ROUND_HALF_UP)
    text = str(value).strip().replace("₹", "").replace("INR", "").replace("Rs.", "").replace("Rs", "")
    text = text.replace(",", "").replace(" ", "")
    if not text:
        return ZERO
    try:
        return Decimal(text).quantize(CENT, rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return ZERO


def load_rules(path: str | Path | None = None) -> dict[str, Any]:
    path = Path(path) if path else Path(__file__).with_name("rules.json")
    with path.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    return normalize_rules(data)


def normalize_rules(data: dict[str, Any]) -> dict[str, Any]:
    """Fill missing keys while retaining user edits from a rules JSON file."""
    defaults = {
        "version": 1,
        "voucher_type": "Journal",
        "ledgers": {
            "expense": "Prepaid Insurance Vehicles",
            "fallback_gst": "Input GST - CHECK",
            "rounding": "Rounding OFF",
        },
        "state_gst_mappings": {
            "29": {"cgst": "INPUT CGST -KAR", "sgst": "INPUT SGST-KAR"},
            "27": {"cgst": "Input CGST_MUM", "sgst": "Input SGST_MUM"},
            "36": {"cgst": "Input CGST-HYD", "sgst": "Input SGST-HYD"},
        },
        "insurer_mappings": {
            "universal_sompo": {
                "name_match": "Universal Sompo General Insurance Company Limited",
                "29": "Universal Sompo General Insurance Company Ltd- BNG",
                "27": "Universal Sompo General Insurance Company -Mum",
                "36": "Universal Sompo General Insurance Company Ltd-HYD",
            },
            "sbi_general": {
                "name_match": "SBI General Insurance Company Limited",
                "default": "SBI General Insurance Company Limited",
            },
            "hdfc_ergo": {
                "name_match": "HDFC ERGO General Insurance Company Limited",
                "default": "HDFC ERGO GENERAL INSURANCE COMPANY LIMITED",
            },
        },
    }
    out = json.loads(json.dumps(defaults))
    if isinstance(data, dict):
        out["voucher_type"] = str(data.get("voucher_type", out["voucher_type"]))
        for key in ("expense", "fallback_gst", "rounding"):
            val = data.get("ledgers", {}).get(key) if isinstance(data.get("ledgers"), dict) else None
            if val is not None:
                out["ledgers"][key] = str(val)
        if isinstance(data.get("state_gst_mappings"), dict):
            for code, mapping in data["state_gst_mappings"].items():
                if isinstance(mapping, dict):
                    out["state_gst_mappings"][str(code)] = {
                        "cgst": str(mapping.get("cgst", "")),
                        "sgst": str(mapping.get("sgst", "")),
                    }
        if isinstance(data.get("insurer_mappings"), dict):
            for key, mapping in data["insurer_mappings"].items():
                if isinstance(mapping, dict):
                    out["insurer_mappings"].setdefault(key, {}).update(
                        {str(k): str(v) for k, v in mapping.items()}
                    )
    return out


def save_rules(rules: dict[str, Any], path: str | Path) -> None:
    with Path(path).open("w", encoding="utf-8") as handle:
        json.dump(normalize_rules(rules), handle, indent=2, ensure_ascii=False)


def _configure_tesseract() -> None:
    cmd = os.environ.get("TESSERACT_CMD", "").strip()
    if cmd:
        pytesseract.pytesseract.tesseract_cmd = cmd
        return
    # Common Windows installer path. This allows the app to find Tesseract
    # even when its installer did not add the executable to PATH.
    candidates = [
        Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
        Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
    ]
    program_files = os.environ.get("ProgramFiles")
    if program_files:
        candidates.append(Path(program_files) / "Tesseract-OCR" / "tesseract.exe")
    for candidate in candidates:
        if candidate.is_file():
            pytesseract.pytesseract.tesseract_cmd = str(candidate)
            return


def _preprocess_image(image: Image.Image) -> Image.Image:
    image = ImageOps.exif_transpose(image).convert("L")
    # Upscaling and contrast help Tesseract on small scans without requiring
    # OpenCV or downloading any model files.
    if image.width < 1800:
        scale = min(2.0, 1800 / max(image.width, 1))
        image = image.resize((int(image.width * scale), int(image.height * scale)))
    image = ImageEnhance.Contrast(image).enhance(1.6)
    image = image.filter(ImageFilter.SHARPEN)
    return image



def _insurance_text_score(text: str) -> int:
    """Heuristic score for insurance invoice fields; higher means more useful."""
    normalized = re.sub(r"\s+", "", (text or "")).lower()
    score = 0
    if re.search(r"tota[l1i]premium", normalized):
        score += 5
    if re.search(r"a\+b|a\+8|aandb", normalized):
        score += 1
    if re.search(r"cgs?t", normalized) or "cgst" in normalized:
        score += 2
    if "sgst" in normalized or "sgs t" in normalized:
        score += 2
    if "igst" in normalized:
        score += 1
    if "grosspremiumpaid" in normalized or "grosspremium" in normalized:
        score += 2
    if "gstin" in normalized or "gstno" in normalized:
        score += 2
    if "invoiceno" in normalized or "invoicenumber" in normalized:
        score += 1
    if "invoicedate" in normalized:
        score += 1
    return score


def _invoice_quality_score(text: str) -> int:
    """Score extracted invoice text by usable accounting fields, not label count.

    PDF text layers can contain every heading but omit or misorder the values in
    a two-column table. A label-only score can therefore prefer a broken text
    layer over perfectly readable OCR. This score rewards complete, arithmetically
    consistent accounting fields.
    """
    if not (text or "").strip():
        return -100
    normalized_text = _clean_text(text)
    parsed = parse_invoice_text(normalized_text)
    score = _insurance_text_score(normalized_text)
    premium = money(parsed.get("premium"))
    cgst = money(parsed.get("cgst"))
    sgst = money(parsed.get("sgst"))
    igst = money(parsed.get("igst"))
    invoice_total = money(parsed.get("invoice_total"))

    if premium > ZERO:
        score += 5
    if re.search(r"total\s+premium[\s\S]{0,80}?(?:A\s*[+&]\s*B|A\s+AND\s+B)", normalized_text, re.I):
        score += 4
    if cgst > ZERO:
        score += 4
    if sgst > ZERO:
        score += 4
    if igst > ZERO:
        score += 3
    if invoice_total > ZERO:
        score += 5
    if parsed.get("customer_gstin"):
        score += 2
    if parsed.get("invoice_number"):
        score += 1

    # A reliable invoice table should reconcile premium + tax to gross paid,
    # allowing sub-rupee rounding. This is the strongest signal for this tool.
    if invoice_total > ZERO and premium > ZERO and (cgst + sgst + igst) > ZERO:
        difference = abs(invoice_total - premium - cgst - sgst - igst)
        if difference < Decimal("1.00"):
            score += 10
        elif difference < Decimal("5.00"):
            score += 3
        else:
            score -= 6

    if re.search(r"vehicle\s+identification\s+no\.?", parsed.get("policy_number", ""), re.I):
        score -= 8
    return score


def _ocr_words_as_rows(data: dict[str, Any]) -> str:
    """Rebuild OCR rows by vertical position, not Tesseract's column order.

    Some insurance PDFs are laid out as tables. Tesseract PSM 3 can return all
    left-column labels first and all right-column amounts afterwards. The word
    boxes still contain coordinates, so group words with the same baseline and
    sort them left-to-right. This restores rows such as "CGST @9% ₹1,513.17".
    """
    words: list[tuple[float, int, int, str]] = []
    texts = data.get("text", [])
    for i, raw in enumerate(texts):
        word = str(raw or "").strip()
        if not word:
            continue
        try:
            confidence = float(data.get("conf", [0] * len(texts))[i])
            left = int(data.get("left", [0] * len(texts))[i])
            top = int(data.get("top", [0] * len(texts))[i])
            height = max(1, int(data.get("height", [1] * len(texts))[i]))
        except (TypeError, ValueError, IndexError):
            continue
        if confidence < 0:
            continue
        center_y = top + height / 2
        words.append((center_y, left, height, word))

    words.sort(key=lambda item: (item[0], item[1]))
    rows: list[dict[str, Any]] = []
    for center_y, left, height, word in words:
        matched = None
        for row in reversed(rows[-10:]):
            tolerance = max(7.0, min(float(height), row["height"]) * 0.60)
            if abs(center_y - row["center_y"]) <= tolerance:
                matched = row
                break
        if matched is None:
            rows.append({"center_y": center_y, "height": float(height), "words": [(left, word)]})
        else:
            count = len(matched["words"])
            matched["center_y"] = (matched["center_y"] * count + center_y) / (count + 1)
            matched["height"] = (matched["height"] * count + height) / (count + 1)
            matched["words"].append((left, word))

    return "\n".join(
        " ".join(word for _, word in sorted(row["words"], key=lambda item: item[0]))
        for row in sorted(rows, key=lambda item: item["center_y"])
    )


def _ocr_image_best_text(image: Image.Image) -> str:
    """Evaluate multiple OCR layouts and prefer the most complete journal data."""
    prepared = _preprocess_image(image)
    candidates: list[str] = []
    for psm in (6, 3, 4, 11):
        config = f"--oem 3 --psm {psm}"
        try:
            candidates.append(pytesseract.image_to_string(prepared, config=config))
            word_data = pytesseract.image_to_data(
                prepared, config=config, output_type=pytesseract.Output.DICT
            )
            row_text = _ocr_words_as_rows(word_data)
            if row_text:
                candidates.append(row_text)
        except pytesseract.pytesseract.TesseractNotFoundError:
            raise
    return max(
        candidates,
        key=lambda value: (_invoice_quality_score(value), _insurance_text_score(value)),
        default="",
    )


def _ocr_pdf_clip(page: fitz.Page, rect: fitz.Rect, scale: float = 4.0) -> str:
    """OCR a high-resolution page region. Useful for dense policy tables."""
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=rect, alpha=False)
    image = Image.open(BytesIO(pix.tobytes("png"))).convert("L")
    image = ImageEnhance.Contrast(image).enhance(1.7).filter(ImageFilter.SHARPEN)
    return pytesseract.image_to_string(image, config="--oem 3 --psm 6")


def _ocr_pdf_page(page: fitz.Page) -> str:
    """OCR scanned insurance pages with insurer/layout-specific table crops.

    These one-page policies are scans. Full-page OCR can merge table columns,
    lose GST labels, and confuse digits. Focused policy-details and premium-table
    crops preserve row/amount relationships for supported insurer templates.
    """
    rect = page.rect
    details_rect = fitz.Rect(
        rect.x0 + rect.width * 0.07,
        rect.y0 + rect.height * 0.05,
        rect.x0 + rect.width * 0.95,
        rect.y0 + rect.height * 0.35,
    )
    table_rect = fitz.Rect(
        rect.x0 + rect.width * 0.07,
        rect.y0 + rect.height * 0.40,
        rect.x0 + rect.width * 0.95,
        rect.y0 + rect.height * 0.59,
    )

    details_text = _ocr_pdf_clip(page, details_rect, scale=4.0)
    # The crop contains the insurer name and policy schedule heading. For this
    # known template, retain the focused crops only; a second full-page OCR pass
    # would reintroduce noisy duplicates of the tax rows.
    normalized = re.sub(r"\s+", " ", details_text).lower()
    compact_details = re.sub(r"\s+", "", details_text).upper()

    if "universal sompo" in normalized and "policy" in normalized:
        table_text = _ocr_pdf_clip(page, table_rect, scale=5.0)
        return details_text + "\n\n[ACCOUNTING PREMIUM TABLE]\n" + table_text

    # HDFC ERGO Mumbai motor schedules have a two-column premium table. Full-
    # page OCR is prone to two specific errors on this template: SGST may be
    # read as "scst", and Gross Premium Paid 19596 may be read as 19506. The
    # focused right-hand crop preserves each label beside its amount. Apply
    # this override ONLY when the customer GSTIN starts with state code 27 and
    # the insurer printed in the footer is HDFC ERGO, leaving KAR/HYD and other
    # insurers on their existing extraction path.
    customer_is_mumbai = bool(re.search(r"27[A-Z0-9]{13}", compact_details))
    if customer_is_mumbai:
        footer_rect = fitz.Rect(
            rect.x0 + rect.width * 0.35,
            rect.y0 + rect.height * 0.82,
            rect.x0 + rect.width * 0.99,
            rect.y0 + rect.height * 0.98,
        )
        footer_text = _ocr_pdf_clip(page, footer_rect, scale=4.0)
        if re.search(r"hdfc\s*ergo", footer_text, re.I):
            mumbai_table_rect = fitz.Rect(
                rect.x0 + rect.width * 0.48,
                rect.y0 + rect.height * 0.405,
                rect.x0 + rect.width * 0.99,
                rect.y0 + rect.height * 0.59,
            )
            mumbai_table_text = _ocr_pdf_clip(page, mumbai_table_rect, scale=5.0)

            # Keep a full-page OCR pass for identifiers/dates/address fields,
            # but remove its noisy premium-table block. The focused premium
            # crop above is authoritative for premium, CGST, SGST and gross.
            full_pix = page.get_pixmap(matrix=fitz.Matrix(3.0, 3.0), alpha=False)
            full_image = Image.open(BytesIO(full_pix.tobytes("png")))
            full_page_text = pytesseract.image_to_string(
                _preprocess_image(full_image), config="--oem 3 --psm 6"
            )
            metadata_text, table_removed = re.subn(
                r"schedule\s+of\s+premium[\s\S]*?point\s+of\s+sale",
                "[MUMBAI HDFC ACCOUNTING TABLE OMITTED FROM FULL PAGE OCR]",
                full_page_text,
                count=1,
                flags=re.I,
            )
            if not table_removed:
                # Fail-safe: omit individual accounting rows if OCR did not
                # retain the table boundaries. The focused crop still supplies
                # the amounts without duplicate tax entries being summed.
                metadata_text = "\n".join(
                    line for line in full_page_text.splitlines()
                    if not re.search(
                        r"total\s+premium|cgst|sgst|scst|gross\s+premium\s+paid|"
                        r"net\s+liability\s+premium",
                        line,
                        re.I,
                    )
                )
            return (
                mumbai_table_text
                + "\n\n[MUMBAI HDFC POLICY DETAILS]\n"
                + metadata_text
                + "\n\n[MUMBAI HDFC DETAILS CROP]\n"
                + details_text
                + "\n\n[MUMBAI HDFC FOOTER]\n"
                + footer_text
            )

    # Fallback for other scanned PDF layouts: one full-page OCR pass.
    pix = page.get_pixmap(matrix=fitz.Matrix(3.0, 3.0), alpha=False)
    image = Image.open(BytesIO(pix.tobytes("png")))
    prepared = _preprocess_image(image)
    return pytesseract.image_to_string(prepared, config="--oem 3 --psm 6")

def extract_document_text(file_name: str, file_bytes: bytes) -> tuple[str, str, int, list[str]]:
    """Return text, extraction method, page count, and a per-file log."""
    _configure_tesseract()
    suffix = Path(file_name).suffix.lower()
    log: list[str] = []
    chunks: list[str] = []

    if suffix == ".pdf":
        try:
            document = fitz.open(stream=file_bytes, filetype="pdf")
        except Exception as exc:
            raise ValueError(f"Could not open PDF: {exc}") from exc
        page_count = len(document)
        used_ocr = False
        used_text = False
        for page_index, page in enumerate(document):
            # sort=True is essential for multi-column insurance tables: default
            # PDF object order can place labels and amount cells in separate runs.
            page_text = page.get_text("text", sort=True) or ""
            compact_text_len = len(re.sub(r"\s", "", page_text))
            if compact_text_len >= 120:
                used_text = True
                compact = re.sub(r"\s+", "", page_text).lower()
                has_accounting_table = bool(re.search(
                    r"premium|liability|c\s*g\s*s\s*t|s\s*g\s*s\s*t|i\s*g\s*s\s*t|gross\s+premium",
                    page_text, re.I,
                ))
                parsed_text = parse_invoice_text(page_text, "PDF page")
                premium_value = money(parsed_text.get("premium"))
                cgst_value = money(parsed_text.get("cgst"))
                sgst_value = money(parsed_text.get("sgst"))
                igst_value = money(parsed_text.get("igst"))
                invoice_total_value = money(parsed_text.get("invoice_total"))
                tax_value = cgst_value + sgst_value + igst_value
                arithmetic_bad = (
                    premium_value <= ZERO
                    or invoice_total_value <= ZERO
                    or (tax_value <= ZERO and not igst_value > ZERO)
                    or (tax_value > ZERO and abs(invoice_total_value - premium_value - tax_value) >= Decimal("1.00"))
                )
                needs_table_ocr = has_accounting_table and arithmetic_bad
                if needs_table_ocr:
                    try:
                        ocr_text = _ocr_pdf_page(page)
                        embedded_score = _invoice_quality_score(page_text)
                        ocr_score = _invoice_quality_score(ocr_text)
                        if ocr_score > embedded_score:
                            chunks.append(ocr_text)
                            used_ocr = True
                            log.append(
                                f"Page {page_index + 1}: local OCR selected (accounting-field score {ocr_score} > embedded-text score {embedded_score})"
                            )
                        else:
                            chunks.append(page_text)
                            log.append(
                                f"Page {page_index + 1}: embedded text selected (accounting-field score {embedded_score} >= OCR score {ocr_score})"
                            )
                    except pytesseract.pytesseract.TesseractNotFoundError:
                        chunks.append(page_text)
                        log.append(f"Page {page_index + 1}: embedded PDF text used; optional OCR assist unavailable")
                    except Exception as exc:
                        chunks.append(page_text)
                        log.append(f"Page {page_index + 1}: embedded PDF text used; OCR assist failed ({exc})")
                else:
                    chunks.append(page_text)
                    log.append(f"Page {page_index + 1}: extracted embedded PDF text (accounting fields consistent or no accounting table)")
            else:
                used_ocr = True
                try:
                    ocr_text = _ocr_pdf_page(page)
                    chunks.append(ocr_text)
                    if "[MUMBAI HDFC" in ocr_text:
                        log.append(
                            f"Page {page_index + 1}: HDFC ERGO Maharashtra (customer GSTIN 27) focused premium-table OCR applied"
                        )
                    else:
                        log.append(f"Page {page_index + 1}: OCR scanned page")
                except pytesseract.pytesseract.TesseractNotFoundError as exc:
                    raise RuntimeError(
                        "Tesseract OCR is not installed or not on PATH. Install Tesseract OCR, "
                        "then restart the app. See README.md."
                    ) from exc
                except Exception as exc:
                    log.append(f"Page {page_index + 1}: OCR failed ({exc})")
        document.close()
        method = "PDF text + OCR" if used_ocr and used_text else ("OCR scanned PDF" if used_ocr else "PDF text")
        return "\n".join(chunks), method, page_count, log

    if suffix in {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}:
        try:
            image = Image.open(BytesIO(file_bytes))
            page_count = 1
            text = _ocr_image_best_text(image)
            log.append("Image processed with local Tesseract OCR")
            return text, "Image OCR", page_count, log
        except pytesseract.pytesseract.TesseractNotFoundError as exc:
            raise RuntimeError(
                "Tesseract OCR is not installed or not on PATH. Install Tesseract OCR, "
                "then restart the app. See README.md."
            ) from exc
        except Exception as exc:
            raise ValueError(f"Could not read image: {exc}") from exc

    raise ValueError(f"Unsupported file type: {suffix}. Use PDF, JPG, PNG, TIFF, or BMP.")


def _clean_text(text: str) -> str:
    text = text.replace("\u00a0", " ").replace("\u2013", "-").replace("\u2014", "-")
    text = text.replace("\ufeff", "")
    # Common local-Tesseract confusions in the word PREMIUM. These occur in
    # the uploaded Universal Sompo table as "Premiom" / "Premiim" and prevent
    # the highest-priority Total Premium (A+B) rule from matching.
    text = re.sub(r"\bpremi[o0i]m\b", "Premium", text, flags=re.I)
    text = re.sub(r"\bpremiim\b|\bpremiam\b|\bpremiun\b", "Premium", text, flags=re.I)
    return "\n".join(re.sub(r"[ \t]+", " ", line).strip() for line in text.splitlines())


def _find_amounts(text: str) -> list[Decimal]:
    """Get plausible amounts from one short line, skipping percentage values."""
    found: list[Decimal] = []
    pattern = re.compile(r"(?<![A-Z0-9])(?:₹\s*|INR\s*|Rs\.?\s*)?(-?\d[\d,]*(?:\.\d{1,2})?)", re.I)
    for match in pattern.finditer(text):
        tail = text[match.end():]
        if re.match(r"\s*%", tail):
            continue
        token = match.group(1).replace(",", "")
        try:
            value = Decimal(token)
        except InvalidOperation:
            continue
        # Skip likely years / long policy IDs when OCR found a labelled row.
        if abs(value) >= Decimal("100000000000"):
            continue
        found.append(value.quantize(CENT, rounding=ROUND_HALF_UP))
    return found


def _amount_after_labels(text: str, labels: Iterable[str], *, sum_matches: bool = False) -> Decimal | None:
    """Find amount(s) on a labelled row, allowing OCR to split label/value lines.

    Tesseract and PDF text layers sometimes place a table label and its right-hand
    amount on separate lines. Look ahead a few lines, but stop at the next row
    label so we never borrow an amount from a neighbouring tax/premium row.
    """
    lines = [line.strip() for line in text.splitlines()]
    patterns = [re.compile(label, re.I) for label in labels]
    row_start = re.compile(
        r"^(?:basic|third party|net|total|gross|own damage|liability|cgst|sgst|igst|"
        r"cess|stamp duty|premium|gst|tax|amount|grand total|invoice total|"
        r"c\s*g\s*s\s*t|s\s*g\s*s\s*t|i\s*g\s*s\s*t)\b",
        re.I,
    )
    matches: list[Decimal] = []
    for index, line in enumerate(lines):
        if not any(pattern.search(line) for pattern in patterns):
            continue
        candidates = _find_amounts(line)
        if candidates:
            matches.append(candidates[-1])
            continue
        # OCR can put the amount cell on the next one or two lines. Accept only
        # a value-like line (possibly with a currency symbol), never another
        # labelled row or a GSTIN/policy identifier.
        for next_index in range(index + 1, min(index + 4, len(lines))):
            candidate_line = lines[next_index]
            if not candidate_line:
                continue
            if row_start.search(candidate_line):
                break
            amount_match = re.fullmatch(
                r"(?:(?:₹|INR|Rs\.?)\s*)?-?\d[\d,]*(?:\.\d{1,2})?",
                candidate_line,
                re.I,
            )
            if amount_match:
                found = _find_amounts(candidate_line)
                if found:
                    matches.append(found[-1])
                break
            # Some OCR engines retain a currency sign or table separator.
            cleaned = candidate_line.strip(" |¦:;\t")
            if re.fullmatch(r"(?:(?:₹|INR|Rs\.?)\s*)?-?\d[\d,]*(?:\.\d{1,2})?", cleaned, re.I):
                found = _find_amounts(cleaned)
                if found:
                    matches.append(found[-1])
                break
    if not matches:
        return None
    if sum_matches:
        return sum(matches, ZERO).quantize(CENT, rounding=ROUND_HALF_UP)
    return matches[0]


def _extract_premium_amount(text: str) -> Decimal | None:
    """Find the pre-GST premium without mistaking a liability component for it.

    For motor policies with an Own Damage (A) + Liability (B) split, the
    accounting base is the explicit Total Premium (A+B) row. A Net Premium (B)
    / Net Liability Premium (B) row is only one component and must never be
    selected as the prepaid-insurance amount.
    """
    # First choice: the explicit A+B total, including OCR spacing variants.
    a_plus_b_patterns = [
        r"\btota[l1i]\s+premium\b[^\n]{0,45}(?:\(?\s*A\s*[+&]\s*B\s*\)?|A\s+AND\s+B|A\s*[+&]\s*8)",
        r"\btotal\s+premium\s*[-:]?\s*A\s*[+&]\s*B\b",
    ]
    value = _amount_after_labels(text, a_plus_b_patterns)
    if value is not None:
        return value

    # A non-gross Total Premium row is also a valid pre-GST amount where the
    # insurer does not print the A+B notation explicitly.
    generic_total = _amount_after_labels(text, [
        r"\btota[l1i]\s+premium\b(?![^\n]*(?:including|incl\.?|gross).{0,12}(?:gst|tax))"
    ])
    if generic_total is not None:
        return generic_total

    # Some insurers call the pre-tax value Total Basic Premium / Premium
    # Before GST. These are safe fallbacks; a net-liability component is not.
    explicit_pretax = _amount_after_labels(text, [
        r"\bpremium\s*(?:before|excluding|excl\.?)\s*(?:gst|tax)\b",
        r"\btotal\s+basic\s+premium\b",
    ])
    if explicit_pretax is not None:
        return explicit_pretax

    # A liability table means a bare Net Premium (B) is a component, not the
    # full insurance premium. Refuse to guess; the review screen will show 0
    # and the narration will carry a CHECK flag instead of posting a wrong Dr.
    has_component_b = bool(re.search(
        r"\bnet\s+(?:liability\s+)?premium\s*\(?\s*B\s*\)?|"
        r"\bliability\s+section\s*\(?\s*B\s*\)?",
        text,
        re.I,
    ))
    if has_component_b:
        return None

    # Last-resort fallback for simpler non-split policies only.
    return _amount_after_labels(text, [
        r"\bnet\s+premium\b(?![^\n]{0,25}(?:liability|\(\s*B\s*\)|\bB\b))",
        r"\bbasic\s+premium\b",
    ])


def _extract_text_field(text: str, labels: Iterable[str]) -> str:
    lines = text.splitlines()
    for index, line in enumerate(lines):
        for label in labels:
            m = re.search(label, line, re.I)
            if not m:
                continue
            rest = line[m.end():].strip(" \t:.-#")
            # Remove unrelated trailing labels if the invoice uses a dense row.
            rest = re.split(r"\s{2,}(?:Date|GSTIN|Address|Place of Supply)\b", rest, maxsplit=1, flags=re.I)[0]
            if rest and not re.fullmatch(r"[|/\- ]+", rest):
                return rest.strip()
            for next_line in lines[index + 1:index + 3]:
                candidate = next_line.strip()
                if candidate and not re.search(r"^(?:GSTIN|PAN|Date|Address)\b", candidate, re.I):
                    return candidate
    return ""


def _extract_policy_number(text: str) -> str:
    """Extract a real policy identifier, tolerating OCR errors in 'Policy No'."""
    lines = text.splitlines()
    # OCR on these scans sometimes drops the initial P ("olicyNo") or reads it
    # as R/B ("Rolicy/No", "Bolicy:No"). The value remains strongly delimited.
    policy_label = re.compile(
        r"\b(?:[PRB])?olicy\s*[/.:]?\s*(?:no\.?|number|#)|"
        r"\b(?:[PRB])?olicy\s*reference\s*(?:no\.?|number)", re.I
    )
    non_value_labels = re.compile(
        r"^(?:vehicle\s*identification\s*no\.?|vehicle\s*identification|registration|regn?\.?|chassis|engine|vin|gstin|gst\s*no|"
        r"policy\s*period|policy\s*type|insured|customer|invoice|date|period|address)\b",
        re.I,
    )
    next_field = re.compile(
        r"\s+(?=(?:vehicle\s*identification\s*no\.?|vehicle\s*identification|registration|regn?\.?|chassis|engine|vin|"
        r"geographical|insured|invoice|policy|date|period|address)\b)", re.I
    )
    for index, line in enumerate(lines):
        match = policy_label.search(line)
        if not match:
            continue
        candidates = [line[match.end():].strip(" \t:.-#")]
        candidates.extend(candidate.strip() for candidate in lines[index + 1:index + 3])
        for candidate in candidates:
            candidate = next_field.split(candidate, maxsplit=1)[0]
            candidate = re.split(r"\s{2,}", candidate, maxsplit=1)[0].strip(" :#-|")
            if not candidate:
                continue
            if non_value_labels.search(candidate):
                break
            compact = re.sub(r"\s+", "", candidate)
            if len(compact) >= 5 and re.fullmatch(r"[A-Z0-9][A-Z0-9./_-]*", compact, re.I) and re.search(r"\d", compact):
                return compact
    return ""


_DATE_TOKEN_PATTERNS = [
    re.compile(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b"),
    re.compile(r"\b\d{4}[/-]\d{1,2}[/-]\d{1,2}\b"),
    re.compile(r"\b\d{1,2}[ .-][A-Za-z]{3,9}[ .-]\d{2,4}\b"),
    re.compile(r"\b[A-Za-z]{3,9}[ .-]\d{1,2},?[ .-]\d{2,4}\b"),
]


def _date_tokens(text: str) -> list[str]:
    found: list[str] = []
    for pattern in _DATE_TOKEN_PATTERNS:
        found.extend(m.group(0) for m in pattern.finditer(text))
    # Keep source order and remove duplicates.
    ordered: list[str] = []
    for token in sorted(found, key=lambda s: min((text.find(s), len(s)))):
        if token not in ordered:
            ordered.append(token)
    return ordered


def parse_date(value: Any) -> date | None:
    if value is None:
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    raw = str(value).strip().replace(",", "")
    if not raw:
        return None
    formats = [
        "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y",
        "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
        "%d %b %Y", "%d-%b-%Y", "%d %B %Y", "%d-%B-%Y",
        "%b %d %Y", "%B %d %Y", "%b %d-%Y", "%B %d-%Y",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            pass
    return None


def format_date_ddmmyyyy(value: Any) -> str:
    parsed = parse_date(value)
    return parsed.strftime("%d-%m-%Y") if parsed else ""


def _extract_labeled_date(text: str, labels: Iterable[str]) -> str:
    for line in text.splitlines():
        if any(re.search(label, line, re.I) for label in labels):
            for token in _date_tokens(line):
                parsed = parse_date(token)
                if parsed:
                    return parsed.strftime("%d-%m-%Y")
    return ""


def _extract_customer_gstin(text: str) -> str:
    gstin_re = re.compile(r"\d{2}[A-Z]{5}\d{4}[A-Z][A-Z0-9]Z[A-Z0-9]", re.I)
    lines = text.splitlines()
    priority_terms = ("recipient", "customer", "insured", "bill to", "buyer", "policyholder", "proposer")
    reject_terms = ("supplier", "insurer", "issued by", "from:", "seller", "insurance company")

    # Strongest signal: a GSTIN on the same line as an explicit customer label.
    for index, line in enumerate(lines):
        compact = re.sub(r"\s+", "", line)
        values = gstin_re.findall(compact)
        context = line.lower()
        if values and any(term in context for term in priority_terms) and not any(term in context for term in reject_terms):
            return values[0].upper()
        if not values and any(term in context for term in priority_terms) and not any(term in context for term in reject_terms):
            for nearby in lines[index + 1:index + 4]:
                nearby_lower = nearby.lower()
                if any(term in nearby_lower for term in reject_terms):
                    continue
                values = gstin_re.findall(re.sub(r"\s+", "", nearby))
                if values:
                    return values[0].upper()

    # Next, look in the local block around an insured/customer heading.
    for index, line in enumerate(lines):
        if re.search(r"customer|insured|bill\s*to|recipient|proposer|policyholder", line, re.I):
            # Prefer an explicitly-recipient-labelled GSTIN in this block over
            # a supplier GSTIN that may also appear nearby on the invoice.
            block_lines = lines[index:index + 10]
            for offset, block_line in enumerate(block_lines):
                if re.search(r"recipient|customer|buyer|bill\s*to", block_line, re.I):
                    values = gstin_re.findall(re.sub(r"\s+", "", block_line))
                    if values:
                        return values[0].upper()
                    for nearby in block_lines[offset + 1:offset + 3]:
                        if any(term in nearby.lower() for term in reject_terms):
                            continue
                        values = gstin_re.findall(re.sub(r"\s+", "", nearby))
                        if values:
                            return values[0].upper()
            # Otherwise only accept a unique GSTIN from non-supplier lines.
            block_values = []
            for block_line in block_lines:
                if any(term in block_line.lower() for term in reject_terms):
                    continue
                block_values.extend(gstin_re.findall(re.sub(r"\s+", "", block_line)))
            unique_values = list(dict.fromkeys(value.upper() for value in block_values))
            if len(unique_values) == 1:
                return unique_values[0]

    # Last resort: if the document contains exactly one non-supplier GSTIN,
    # use it with a review warning added by parse_invoice_text. This handles
    # invoices where OCR separated the GSTIN label from the customer block.
    candidates: list[str] = []
    for index, line in enumerate(lines):
        compact = re.sub(r"\s+", "", line)
        values = gstin_re.findall(compact)
        if not values:
            continue
        context = " ".join(lines[max(0, index - 2): min(len(lines), index + 3)]).lower()
        if any(term in context for term in reject_terms):
            continue
        candidates.extend(value.upper() for value in values)
    unique = list(dict.fromkeys(candidates))
    return unique[0] if len(unique) == 1 else ""


def _detect_insurer(text: str) -> str:
    lower = text.lower()
    known = [
        ("universal sompo", "Universal Sompo General Insurance Company Limited"),
        ("sbi general", "SBI General Insurance Company Limited"),
        ("hdfc ergo", "HDFC ERGO General Insurance Company Limited"),
    ]
    for needle, display in known:
        if needle in lower:
            return display
    for line in text.splitlines()[:45]:
        if re.search(r"\b(?:General|Health|Life) Insurance Company\b", line, re.I):
            candidate = line.strip(" :|-")
            if len(candidate) <= 120:
                return candidate
    return ""


def _detect_policy_type(text: str) -> str:
    checks = [
        (r"commercial\s+vehicle", "Commercial Vehicle"),
        (r"private\s+car", "Private Car"),
        (r"two[- ]?wheeler|two\s+wheel", "Two Wheeler"),
        (r"motor\s+package|motor\s+insurance|motor\s+policy", "Motor"),
        (r"vehicle\s+insurance", "Vehicle"),
        (r"health\s+insurance|mediclaim", "Health"),
        (r"fire\s+insurance", "Fire"),
        (r"marine\s+insurance", "Marine"),
        (r"travel\s+insurance", "Travel"),
    ]
    for pattern, label in checks:
        if re.search(pattern, text, re.I):
            return label
    return ""


def _extract_period(text: str) -> tuple[str, str]:
    from_date = _extract_labeled_date(text, [
        r"period\s*(?:of\s*)?(?:insurance\s*)?from", r"policy\s*start\s*date",
        r"risk\s*commencement\s*date", r"effective\s*date", r"from\s*date",
    ])
    to_date = _extract_labeled_date(text, [
        r"period\s*(?:of\s*)?(?:insurance\s*)?to", r"policy\s*end\s*date",
        r"expiry\s*date", r"expiration\s*date", r"to\s*date",
    ])
    if not from_date or not to_date:
        for line in text.splitlines():
            if re.search(r"policy\s*period|period\s*of\s*insurance|insurance\s*period", line, re.I):
                dates = [format_date_ddmmyyyy(x) for x in _date_tokens(line)]
                dates = [x for x in dates if x]
                if len(dates) >= 2:
                    from_date = from_date or dates[0]
                    to_date = to_date or dates[1]
                    break
    return from_date, to_date


def parse_invoice_text(text: str, file_name: str = "invoice") -> dict[str, Any]:
    clean = _clean_text(text)

    invoice_number = _extract_text_field(clean, [
        r"tax\s*invoice\s*(?:no\.?|number|#)", r"invoice\s*(?:no\.?|number|#|num\.? )",
        r"document\s*(?:no\.?|number)",
    ])
    # Keep only a compact identifier if line contains multiple fields.
    invoice_number = re.split(r"\s{2,}|\s+(?:Date|Dated|Invoice Date)\b", invoice_number, maxsplit=1, flags=re.I)[0].strip()
    invoice_number = invoice_number.strip(" :#-")
    # OCR may insert spaces into the insurer's alphanumeric invoice identifier.
    invoice_number = re.sub(r"\s+", "", invoice_number)

    policy_number = _extract_policy_number(clean)

    invoice_date = _extract_labeled_date(clean, [
        r"invoice\s*date", r"date\s*of\s*invoice", r"tax\s*invoice\s*date", r"document\s*date",
    ])
    invoice_date_from_policy_issue = False
    if not invoice_date:
        # This Maruti/Universal Sompo certificate has no separately labelled
        # invoice date. Use the printed Policy Issued On date as a transparent
        # reference-date fallback and flag it for review in the narration.
        invoice_date = _extract_labeled_date(clean, [r"policy\s*issued\s*on", r"issued\s*on"])
        invoice_date_from_policy_issue = bool(invoice_date)

    insured = _extract_text_field(clean, [
        r"name\s*of\s*insured", r"insured\s*name", r"policy\s*holder\s*name",
        r"policyholder", r"proposer\s*name", r"customer\s*name", r"insured\s*:",
    ])
    # Some motor invoice templates place "GST No & State" immediately after
    # the customer's legal name. Do not include that heading in the insured
    # name that is placed in the narration.
    insured = re.split(
        r"\s+(?:GST\s*(?:No\.?|IN|Number)|GSTIN\b|GST\s+No\s*&\s*State\b|"
        r"PAN\b|PORNO\b).*?$",
        insured,
        maxsplit=1,
        flags=re.I,
    )[0].strip(" :|-|")
    insurer = _detect_insurer(clean)
    policy_type = _detect_policy_type(clean)
    customer_gstin = _extract_customer_gstin(clean)
    period_from, period_to = _extract_period(clean)

    # Select the premium row by semantic priority, not by the order in which
    # PDF text happens to be returned. Total Premium (A+B) must beat Net
    # Liability Premium (B) even when the latter appears first.
    premium = _extract_premium_amount(clean)
    cgst = _amount_after_labels(clean, [r"\bC\s*G\s*S\s*T\b"], sum_matches=True)
    sgst = _amount_after_labels(clean, [r"\bS\s*G\s*S\s*T\b"], sum_matches=True)
    igst = _amount_after_labels(clean, [r"\bI\s*G\s*S\s*T\b"], sum_matches=True)
    total_gst = _amount_after_labels(clean, [
        r"total\s+(?:GST|tax)", r"GST\s+amount", r"total\s+tax\s+amount",
    ])
    invoice_total = _amount_after_labels(clean, [
        r"total\s+invoice\s+amount", r"invoice\s+total", r"amount\s+payable",
        r"total\s+amount\s+payable", r"grand\s+total", r"total\s+premium\s*\(?\s*(?:including|incl\.?)\s*(?:gst|tax)\s*\)?",
        r"gross\s+premium", r"net\s+payable", r"total\s+payable",
    ])

    warnings: list[str] = []
    if invoice_date_from_policy_issue:
        warnings.append("INVOICE DATE NOT SEPARATELY PRINTED; POLICY ISSUE DATE USED AS REFERENCE DATE")
    if premium is None and re.search(
        r"\bnet\s+(?:liability\s+)?premium\s*\(?\s*B\s*\)?|\bliability\s+section\s*\(?\s*B\s*\)?",
        clean,
        re.I,
    ):
        warnings.append("TOTAL PREMIUM (A+B) NOT READ; DO NOT USE NET PREMIUM (B) - CORRECT PREMIUM BEFORE GENERATING")

    # On a local Indian invoice, CGST and SGST are equal components. OCR may
    # miss one amount cell even when both row labels are visible. Infer only
    # the missing counterpart (never overwrite two amounts that were read),
    # and add a CHECK note so the inferred value is reviewed by the user.
    has_cgst_label = bool(re.search(r"\bC\s*G\s*S\s*T\b", clean, re.I))
    has_sgst_label = bool(re.search(r"\bS\s*G\s*S\s*T\b", clean, re.I))
    has_igst_label = bool(re.search(r"\bI\s*G\s*S\s*T\b", clean, re.I))
    if igst in (None, ZERO) and not has_igst_label and has_cgst_label and has_sgst_label:
        if (cgst is None or cgst == ZERO) and sgst is not None and sgst > ZERO:
            cgst = sgst
            warnings.append("CGST AMOUNT INFERRED TO MATCH SGST - VERIFY SOURCE INVOICE")
        elif (sgst is None or sgst == ZERO) and cgst is not None and cgst > ZERO:
            sgst = cgst
            warnings.append("SGST AMOUNT INFERRED TO MATCH CGST - VERIFY SOURCE INVOICE")

    result = {
        "file_name": file_name,
        "insurer": insurer,
        "insurer_ledger": "",
        "insured": insured,
        "policy_type": policy_type,
        "policy_number": policy_number,
        "invoice_number": invoice_number,
        "invoice_date": invoice_date,
        "period_from": period_from,
        "period_to": period_to,
        "customer_gstin": customer_gstin,
        "premium": premium if premium is not None else ZERO,
        "cgst": cgst if cgst is not None else ZERO,
        "sgst": sgst if sgst is not None else ZERO,
        "igst": igst if igst is not None else ZERO,
        "total_gst": total_gst if total_gst is not None else ZERO,
        "invoice_total": invoice_total if invoice_total is not None else ZERO,
        "_text": clean,
        "_warnings": warnings,
    }
    return result


def parse_invoice_file(file_name: str, file_bytes: bytes) -> tuple[dict[str, Any], str, int, list[str]]:
    text, method, pages, log = extract_document_text(file_name, file_bytes)
    data = parse_invoice_text(text, file_name)
    data["_extraction_method"] = method
    data["_page_count"] = pages
    data["_log"] = log
    if not text.strip():
        data["_warnings"].append("No readable text found")
    return data, method, pages, log


def _detect_insurer_ledger(insurer: str, state_code: str, rules: dict[str, Any]) -> tuple[str, bool]:
    low = (insurer or "").lower()
    mappings = rules.get("insurer_mappings", {})
    for key, mapping in mappings.items():
        name_match = str(mapping.get("name_match", key)).lower()
        if name_match in low or key.replace("_", " ") in low:
            if key == "universal_sompo":
                ledger = mapping.get(state_code, "")
                return str(ledger or insurer), bool(ledger)
            ledger = mapping.get("default", "")
            return str(ledger or insurer), bool(ledger)
    return (insurer or "INSURER - CHECK"), False


def _add_row(rows: list[dict[str, Any]], voucher_date: str, voucher_type: str,
             voucher_number: str, ledger_name: str, amount: Decimal, drcr: str,
             narration: str, reference_no: str, reference_date: str) -> None:
    amount = money(amount)
    if amount <= ZERO:
        return
    rows.append({
        "Voucher Date": voucher_date,
        "Voucher Type Name": voucher_type,
        "Voucher Number": voucher_number,
        "Ledger Name": ledger_name,
        "Ledger Amount": amount,
        "Ledger Amount Dr/Cr": drcr,
        "Voucher Narration": narration,
        "Reference No.": reference_no,
        "Reference Date": reference_date,
    })


def build_voucher_rows(invoice: dict[str, Any], rules: dict[str, Any], voucher_day: date | None = None) -> tuple[list[dict[str, Any]], list[str]]:
    """Build balanced ledger lines for a single invoice and return check flags."""
    rules = normalize_rules(rules)
    # Use India local date even when the app host/server itself runs in UTC.
    india_tz = timezone(timedelta(hours=5, minutes=30))
    voucher_day = voucher_day or datetime.now(india_tz).date()
    voucher_date = "'" + voucher_day.strftime("%d-%m-%Y")
    invoice_date = format_date_ddmmyyyy(invoice.get("invoice_date", ""))
    # Reference Date is formatted dd-MM-yyyy without a literal apostrophe.
    # Only Voucher Date, Voucher Number and Reference No. use the requested
    # leading apostrophe for Excel text handling.
    reference_date = invoice_date if invoice_date else ""

    invoice_number = str(invoice.get("invoice_number") or "").strip()
    policy_number = str(invoice.get("policy_number") or "").strip()
    voucher_number_value = invoice_number or policy_number
    voucher_number = "'" + voucher_number_value if voucher_number_value else ""
    reference_no = "'" + invoice_number if invoice_number else ""

    insurer = str(invoice.get("insurer") or "").strip()
    insured = str(invoice.get("insured") or "").strip()
    policy_type = str(invoice.get("policy_type") or "").strip()
    period_from = format_date_ddmmyyyy(invoice.get("period_from", ""))
    period_to = format_date_ddmmyyyy(invoice.get("period_to", ""))
    gstin = re.sub(r"\s", "", str(invoice.get("customer_gstin") or "")).upper()
    gstin_is_valid = bool(re.fullmatch(r"\d{2}[A-Z]{5}\d{4}[A-Z][A-Z0-9]Z[A-Z0-9]", gstin))
    state_code = gstin[:2] if gstin_is_valid else ""

    flags: list[str] = [str(warning).upper() for warning in invoice.get("_warnings", []) if str(warning).strip()]
    if not invoice_number:
        flags.append("INVOICE NUMBER MISSING")
    if not policy_number:
        flags.append("POLICY NUMBER MISSING")
    if not invoice_date:
        flags.append("INVOICE DATE MISSING")
    if not insured:
        flags.append("INSURED NAME MISSING")
    if not insurer:
        flags.append("INSURER NOT IDENTIFIED")
    if not gstin:
        flags.append("CUSTOMER GSTIN MISSING")
    elif not gstin_is_valid:
        flags.append("CUSTOMER GSTIN FORMAT NEEDS REVIEW")

    state_mapping = rules.get("state_gst_mappings", {}).get(state_code)
    if not state_mapping or not state_mapping.get("cgst") or not state_mapping.get("sgst"):
        flags.append(f"GSTIN STATE {state_code or 'UNKNOWN'} NOT MAPPED")

    mapped_insurer_ledger, insurer_mapped = _detect_insurer_ledger(insurer, state_code, rules)
    insurer_ledger = str(invoice.get("insurer_ledger") or mapped_insurer_ledger).strip()
    if not insurer_mapped:
        flags.append("INSURER LEDGER NOT MAPPED - REVIEW LEDGER")

    premium = money(invoice.get("premium"))
    cgst = money(invoice.get("cgst"))
    sgst = money(invoice.get("sgst"))
    igst = money(invoice.get("igst"))
    explicit_total_gst = money(invoice.get("total_gst"))
    invoice_total = money(invoice.get("invoice_total"))
    if premium <= ZERO:
        flags.append("PREMIUM AMOUNT MISSING OR ZERO")

    if cgst > ZERO and sgst > ZERO and cgst != sgst:
        flags.append("CGST AND SGST AMOUNTS DIFFER - VERIFY SOURCE INVOICE")

    explicit_components = (cgst + sgst + igst).quantize(CENT)
    tax_total = max(explicit_total_gst, explicit_components).quantize(CENT)
    tax_breakup_missing = False
    if explicit_total_gst > ZERO and explicit_components > ZERO and abs(explicit_total_gst - explicit_components) >= CENT:
        tax_breakup_missing = True
        flags.append("GST TOTAL DOES NOT MATCH TAX COMPONENTS")
    if (cgst > ZERO) != (sgst > ZERO) and explicit_total_gst <= ZERO:
        flags.append("ONLY ONE OF CGST OR SGST WAS READ")

    # If tax breakup is not available, infer total tax from total less premium,
    # but book it to the fallback GST ledger and clearly flag it for review.
    if tax_total <= ZERO and invoice_total > ZERO and premium > ZERO:
        inferred_tax = (invoice_total - premium).quantize(CENT)
        if inferred_tax >= Decimal("1.00"):
            tax_total = inferred_tax
            tax_breakup_missing = True
            flags.append("GST BREAKUP NOT READ - TOTAL GST SENT TO FALLBACK LEDGER")
    elif explicit_total_gst > explicit_components and explicit_total_gst > ZERO:
        tax_breakup_missing = True
    elif explicit_components > ZERO and explicit_total_gst <= ZERO:
        # Only infer missing tax from the invoice total when the tax split is
        # genuinely incomplete. If CGST and SGST are both explicitly present,
        # the remaining difference is a rounding difference by definition for
        # this journal rule; it must not be silently reclassified as GST.
        tax_split_incomplete = (
            igst <= ZERO and (cgst <= ZERO or sgst <= ZERO)
        )
        residual = ZERO
        if tax_split_incomplete and invoice_total > ZERO and premium > ZERO:
            residual = (invoice_total - premium - explicit_components).quantize(CENT)
        if residual >= Decimal("1.00"):
            tax_total = explicit_components + residual
            tax_breakup_missing = True
            flags.append("PART OF GST BREAKUP NOT READ - RESIDUAL SENT TO FALLBACK LEDGER")

    if explicit_total_gst > explicit_components:
        residual = (explicit_total_gst - explicit_components).quantize(CENT)
        if residual >= Decimal("0.01"):
            tax_breakup_missing = True

    if invoice_total <= ZERO:
        invoice_total = (premium + tax_total).quantize(CENT)
        flags.append("INVOICE TOTAL MISSING - CALCULATED FROM PREMIUM AND GST")

    # Determine which tax ledgers can safely be used.
    tax_lines: list[tuple[str, Decimal]] = []
    fallback_ledger = rules["ledgers"]["fallback_gst"]
    if tax_total > ZERO:
        can_use_state_ledgers = bool(state_mapping and state_mapping.get("cgst") and state_mapping.get("sgst"))
        if not can_use_state_ledgers:
            tax_lines.append((fallback_ledger, tax_total))
        elif tax_breakup_missing or explicit_total_gst > explicit_components:
            # Preserve any clearly extracted split, and place the unknown part
            # in the configurable fallback GST ledger.
            known = explicit_components
            if cgst > ZERO:
                tax_lines.append((state_mapping["cgst"], cgst))
            if sgst > ZERO:
                tax_lines.append((state_mapping["sgst"], sgst))
            remaining = (tax_total - known).quantize(CENT)
            if igst > ZERO:
                # There is no IGST mapping in the requested rules.
                remaining = (remaining + igst).quantize(CENT)
                flags.append("IGST LEDGER NOT CONFIGURED")
            if remaining > ZERO:
                tax_lines.append((fallback_ledger, remaining))
            if not tax_lines:
                tax_lines.append((fallback_ledger, tax_total))
        elif igst > ZERO:
            if cgst > ZERO:
                tax_lines.append((state_mapping["cgst"], cgst))
            if sgst > ZERO:
                tax_lines.append((state_mapping["sgst"], sgst))
            tax_lines.append((fallback_ledger, igst))
            flags.append("IGST LEDGER NOT CONFIGURED")
        else:
            if cgst > ZERO:
                tax_lines.append((state_mapping["cgst"], cgst))
            if sgst > ZERO:
                tax_lines.append((state_mapping["sgst"], sgst))
            unallocated = (tax_total - cgst - sgst).quantize(CENT)
            if unallocated > ZERO:
                tax_lines.append((fallback_ledger, unallocated))
                flags.append("GST BALANCE SENT TO FALLBACK LEDGER")
            if not tax_lines:
                tax_lines.append((fallback_ledger, tax_total))
                flags.append("GST BREAKUP NOT AVAILABLE")

    base_plus_tax = (premium + tax_total).quantize(CENT)
    rounding_difference = (invoice_total - base_plus_tax).quantize(CENT)
    if abs(rounding_difference) >= Decimal("1.00"):
        flags.append(f"INVOICE TOTAL DIFFERENCE {rounding_difference:+.2f} - REVIEW")

    # Ensure the user sees a note for incomplete policy-period extraction.
    if not period_from or not period_to:
        flags.append("POLICY PERIOD NOT FULLY READ")

    flags = list(dict.fromkeys(flags))
    period_text = f"{period_from or 'N/A'} to {period_to or 'N/A'}"
    narration = (
        f"Being {policy_type or 'insurance'} premium of {insurer or 'unknown insurer'} "
        f"for {insured or 'unknown insured'}, Policy No {policy_number or 'N/A'}, "
        f"Inv No {invoice_number or 'N/A'}, period {period_text}"
    )
    if flags:
        narration += " [CHECK " + "; ".join(flags) + "]"

    rows: list[dict[str, Any]] = []
    voucher_type = str(rules.get("voucher_type") or "Journal")
    ledgers = rules["ledgers"]
    _add_row(rows, voucher_date, voucher_type, voucher_number, ledgers["expense"], premium,
             "Dr", narration, reference_no, reference_date)
    for ledger, amount in tax_lines:
        _add_row(rows, voucher_date, voucher_type, voucher_number, ledger, amount,
                 "Dr", narration, reference_no, reference_date)
    if rounding_difference > ZERO:
        _add_row(rows, voucher_date, voucher_type, voucher_number, ledgers["rounding"], rounding_difference,
                 "Dr", narration, reference_no, reference_date)
    elif rounding_difference < ZERO:
        _add_row(rows, voucher_date, voucher_type, voucher_number, ledgers["rounding"], abs(rounding_difference),
                 "Cr", narration, reference_no, reference_date)
    _add_row(rows, voucher_date, voucher_type, voucher_number, insurer_ledger, invoice_total,
             "Cr", narration, reference_no, reference_date)

    return rows, flags


def rows_to_plain(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Return ordered export-ready rows with amount as a 2-decimal string."""
    result = []
    for row in rows:
        result.append({
            col: (f"{money(row[col]):.2f}" if col == "Ledger Amount" else row.get(col, ""))
            for col in COLUMNS
        })
    return result


def write_csv(rows: list[dict[str, Any]], output_path: str | Path) -> None:
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=COLUMNS, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows_to_plain(rows))


def _vendor_sheet_title(ledger_name: str, used_titles: set[str]) -> str:
    """Return a readable, Excel-safe worksheet title for a Tally vendor ledger.

    The five currently used Tally ledger names get concise, stable tab names.
    Any future vendor automatically receives a tab based on its ledger name.
    """
    ledger = str(ledger_name or "").strip() or "Unassigned Vendor"
    normalized = re.sub(r"[^a-z0-9]+", "", ledger.lower())
    known_titles = {
        "universalsompogeneralinsurancecompanyltdbng": "Universal Sompo - BNG",
        "universalsompogeneralinsurancecompanyltdhyd": "Universal Sompo - HYD",
        "universalsompogeneralinsurancecompanymum": "Universal Sompo - Mum",
        "sbigeneralinsurancecompanylimited": "SBI General",
        "hdfcergogeneralinsurancecompanylimited": "HDFC ERGO",
    }
    base = known_titles.get(normalized)
    if not base:
        # Excel forbids : \ / ? * [ ] in worksheet titles and limits names to 31 chars.
        base = re.sub(r"[:\\/?*\[\]]", "-", ledger).strip(" '\t") or "Unassigned Vendor"
        base = base[:31]

    candidate = base[:31]
    suffix_number = 2
    while candidate.casefold() in used_titles:
        suffix = f" ({suffix_number})"
        candidate = f"{base[:31 - len(suffix)]}{suffix}"
        suffix_number += 1
    used_titles.add(candidate.casefold())
    return candidate


def _voucher_group_key(row: dict[str, Any]) -> tuple[str, ...]:
    """Identify the voucher shared by all of its ledger rows."""
    return tuple(str(row.get(column, "") or "") for column in (
        "Voucher Date", "Voucher Type Name", "Voucher Number", "Reference No.",
        "Reference Date", "Voucher Narration",
    ))


def _append_xlsx_rows(sheet: Any, rows: list[dict[str, Any]]) -> None:
    """Write and format journal rows consistently on each workbook sheet."""
    sheet.append(COLUMNS)
    for row in rows:
        values = []
        for col in COLUMNS:
            value = row.get(col, "")
            if col == "Ledger Amount":
                value = float(money(value))
            else:
                value = "" if value is None else str(value)
            values.append(value)
        sheet.append(values)

    header_fill = PatternFill("solid", fgColor="17365D")
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    sheet.freeze_panes = "A2"
    if sheet.max_row >= 2:
        sheet.auto_filter.ref = sheet.dimensions
    sheet.row_dimensions[1].height = 32
    widths = [16, 20, 22, 48, 18, 20, 80, 24, 18]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(index)].width = width
    for row_index in range(2, sheet.max_row + 1):
        sheet.cell(row_index, 5).number_format = '#,##0.00'
        for col_index in (1, 3, 8, 9):
            sheet.cell(row_index, col_index).number_format = "@"
        sheet.cell(row_index, 7).alignment = Alignment(wrap_text=True, vertical="top")


def write_xlsx(rows: list[dict[str, Any]], output_path: str | Path) -> None:
    """Export a master journal plus one complete journal sheet per vendor.

    Vendor identity comes from the insurer's main credit ledger row in each
    voucher (the largest credit line, excluding the much smaller rounding
    adjustment by amount). The full voucher is copied to that vendor's sheet,
    so its expense, GST, rounding and insurer lines stay together. New ledger
    names automatically create new tabs; no vendor list is hard-coded for
    grouping.
    """
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()

    # Preserve one complete sheet for users who need the combined journal.
    master = workbook.active
    master.title = "All Vendors"
    _append_xlsx_rows(master, rows)

    # First assemble complete voucher groups so every ledger line from an
    # invoice is placed together on its insurer/vendor sheet.
    vouchers: dict[tuple[str, ...], list[dict[str, Any]]] = {}
    for row in rows:
        vouchers.setdefault(_voucher_group_key(row), []).append(row)

    vendor_rows: dict[str, list[dict[str, Any]]] = {}
    vouchers_by_vendor: dict[str, list[list[dict[str, Any]]]] = {}
    for voucher_rows in vouchers.values():
        credit_rows = [
            row for row in voucher_rows
            if str(row.get("Ledger Amount Dr/Cr", "")).strip().casefold() == "cr"
        ]
        if credit_rows:
            # The insurer is credited for the full invoice total; a rounding
            # credit, if present, is only a small adjustment. Taking the
            # largest credit amount therefore selects the vendor ledger.
            vendor_row = max(credit_rows, key=lambda item: money(item.get("Ledger Amount")))
            vendor_name = str(vendor_row.get("Ledger Name") or "").strip() or "Unassigned Vendor"
        else:
            vendor_name = "Unassigned Vendor"
        vendor_rows.setdefault(vendor_name, []).extend(voucher_rows)
        vouchers_by_vendor.setdefault(vendor_name, []).append(voucher_rows)

    known_order = [
        "Universal Sompo General Insurance Company Ltd- BNG",
        "Universal Sompo General Insurance Company Ltd-HYD",
        "Universal Sompo General Insurance Company -Mum",
        "SBI General Insurance Company Limited",
        "HDFC ERGO GENERAL INSURANCE COMPANY LIMITED",
    ]
    order = {name.casefold(): index for index, name in enumerate(known_order)}
    sorted_vendors = sorted(
        vendor_rows,
        key=lambda name: (order.get(name.casefold(), len(order)), name.casefold()),
    )

    # Add an index so the exact Tally ledger name can always be matched to
    # its Excel tab. Excel limits worksheet titles to 31 characters, therefore
    # known long Tally names use readable short tab labels and future names are
    # safely truncated with collision handling.
    used_titles = {"all vendors", "vendor index"}
    vendor_tab_names: dict[str, str] = {}
    for vendor_name in sorted_vendors:
        tab_name = _vendor_sheet_title(vendor_name, used_titles)
        vendor_tab_names[vendor_name] = tab_name
        vendor_sheet = workbook.create_sheet(tab_name)
        _append_xlsx_rows(vendor_sheet, vendor_rows[vendor_name])

    index_sheet = workbook.create_sheet("Vendor Index", index=1)
    index_sheet.append(["Tally Vendor Ledger Name", "Excel Sheet Name", "Voucher Count", "Journal Row Count"])
    for vendor_name in sorted_vendors:
        index_sheet.append([
            vendor_name,
            vendor_tab_names[vendor_name],
            len(vouchers_by_vendor[vendor_name]),
            len(vendor_rows[vendor_name]),
        ])
    index_header_fill = PatternFill("solid", fgColor="17365D")
    for cell in index_sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = index_header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    index_sheet.freeze_panes = "A2"
    if index_sheet.max_row >= 2:
        index_sheet.auto_filter.ref = index_sheet.dimensions
    index_sheet.column_dimensions["A"].width = 58
    index_sheet.column_dimensions["B"].width = 30
    index_sheet.column_dimensions["C"].width = 18
    index_sheet.column_dimensions["D"].width = 20
    index_sheet.row_dimensions[1].height = 32

    workbook.save(path)


def demo_invoice() -> dict[str, Any]:
    """Deterministic demo based on the verified Universal Sompo motor layout."""
    return {
        "file_name": "DEMO_Universal_Sompo_motor_invoice.pdf",
        "insurer": "Universal Sompo General Insurance Company Limited",
        "insurer_ledger": "",
        "insured": "M/S SAMVREEDDHI MOBILITY PRIVATE LIMITED",
        "policy_type": "Commercial Vehicle",
        "policy_number": "2314/86957518/00/000",
        "invoice_number": "29AL260275390",
        "invoice_date": "07-09-2026",
        "period_from": "07-09-2026",
        "period_to": "06-09-2027",
        "customer_gstin": "29ABLCS9973P1ZW",
        "premium": Decimal("16813.00"),
        "cgst": Decimal("1513.17"),
        "sgst": Decimal("1513.17"),
        "igst": ZERO,
        "total_gst": ZERO,
        "invoice_total": Decimal("19839.00"),
        "_warnings": [
            "INVOICE DATE NOT SEPARATELY PRINTED; POLICY ISSUE DATE USED AS REFERENCE DATE"
        ],
    }
