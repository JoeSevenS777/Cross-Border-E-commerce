# -*- coding: utf-8 -*-
"""Card withdrawal OCR automation

- Process all .xlsx files in the same folder as this script
- For each workbook:
  - Read embedded screenshots in column `提款截圖`
  - OCR them with Tesseract
  - Extract 日期, 提款編號, 提款金額
  - Fill only missing cells in columns 日期 / 提款編號 / 提款金額
  - Summarize newest date + total 提款金額
  - Rename file to: yyMMdd + core name + total amount

This version includes improved amount parsing for:
- slips where withdrawal amount is a negative value (-NT$xxxx or -xx,xxx)
- ignores NT$0 (可提領金額)
- ignores dates and long IDs
"""

import os
import re
import datetime as dt

import openpyxl
from PIL import Image
import pytesseract


# === CONFIGURATION ===

# Folder: the same folder as this script
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Sheet name (change if needed)
SHEET_NAME = "Sheet1"

# Path to Tesseract on your PC – adjust if needed
pytesseract.pytesseract.tesseract_cmd = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"

# OCR language
TESS_LANG = "chi_tra+eng"   # if chi_tra not installed, you can try "eng"


# === OCR + PARSING HELPERS ===

def ocr_image(pil_image):
    """Run OCR on a PIL image and return text."""
    return pytesseract.image_to_string(pil_image, lang=TESS_LANG)


def parse_date_from_text(text):
    """Find a date like 2025/11/06 – we use the first match (from 建立時間)."""
    m = re.search(r"(20\d{2}/\d{2}/\d{2})", text)
    if not m:
        return None
    date_str = m.group(1)
    try:
        return dt.datetime.strptime(date_str, "%Y/%m/%d").date()
    except ValueError:
        return None


def parse_withdraw_id_from_text(text):
    """Find a long number starting with 20, 16–20 digits long."""
    m = re.search(r"\b(20\d{15,20})\b", text)
    if not m:
        return None
    return m.group(1)


def parse_amount_from_text(text):
    """Extract the withdrawal amount for auto-withdrawal slips.

    Rules:
    - Target is always negative on the slip (e.g. -NT$16642, -16,642).
    - Amount is never 0 (ignore NT$0 from 可提領金額).
    - Prefer the keyword line (提領總額 / 提款金額 / 提領金額).
    - Fallback: any other negative amounts in the text.
    - Return positive integer (e.g. 16642) or None.
    """
    KEYWORDS = ("提領總額", "提款金額", "提領金額")
    lines = text.splitlines()

    # ---------- 1) Keyword line with explicit negative amount ----------
    for line in lines:
        clean = line.replace(" ", "")
        if any(k in clean for k in KEYWORDS):
            # e.g. "提領總額 -NT$16642" or "提領總額：-16,642元"
            m = re.search(r"[-−]\s*(?:NT\$|NT|新台幣)?\s*([0-9][0-9,\.]*)", line)
            if m:
                num_str = m.group(1).replace(",", "").strip()
                if "." in num_str:
                    num_str = num_str.split(".", 1)[0]
                try:
                    val = int(num_str)
                except ValueError:
                    continue
                # ignore 0 just in case
                if val != 0:
                    return abs(val)

    # ---------- helper: detect 8-digit dates like 20251204 ----------
    def looks_like_date8(s: str) -> bool:
        if len(s) != 8 or not s.isdigit():
            return False
        year = int(s[:4])
        month = int(s[4:6])
        day = int(s[6:8])
        if not (2000 <= year <= 2099):
            return False
        if not (1 <= month <= 12 and 1 <= day <= 31):
            return False
        return True

    # ---------- 2) Fallback: scan all negative numbers ----------
    candidates = []

    # pattern: minus + optional NT$ + number, anywhere in text
    for m in re.finditer(r"[-−]\s*(?:NT\$|NT|新台幣)?\s*([0-9][0-9,\.]*)", text):
        num_str = m.group(1).replace(",", "").strip()
        if not num_str:
            continue

        # skip super-long numbers → probably IDs
        if len(num_str) > 9:
            continue

        # skip 8-digit dates like 20251204
        if looks_like_date8(num_str):
            continue

        if "." in num_str:
            num_str = num_str.split(".", 1)[0]

        try:
            val = int(num_str)
        except ValueError:
            continue

        # ignore 0 and absurdly large values
        if val == 0 or val > 2_000_000:
            continue

        candidates.append(val)

    if candidates:
        # choose the one with the largest absolute value (the real withdrawal)
        best = max(candidates, key=lambda v: abs(v))
        return abs(best)

    return None


def extract_fields_from_image(pil_image):
    """Full pipeline for a screenshot: OCR -> date, id, amount."""
    text = ocr_image(pil_image)
    date_value = parse_date_from_text(text)
    withdraw_id = parse_withdraw_id_from_text(text)
    amount = parse_amount_from_text(text)
    return date_value, withdraw_id, amount


# === HELPERS FOR SUMMARY & RENAMING ===


def normalize_date_cell(value):
    """Convert cell value to a date object if possible."""
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, str):
        value = value.strip()
        for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d"):
            try:
                return dt.datetime.strptime(value, fmt).date()
            except ValueError:
                continue
    return None


def normalize_amount_cell(value):
    """Convert cell value to an integer amount if possible."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, str):
        s = value.replace(",", "").strip()
        if s == "":
            return None
        try:
            return int(float(s))
        except ValueError:
            return None
    return None


def summarize_sheet(ws, col_date, col_amount):
    """Return (newest_date, total_amount) from the sheet."""
    dates = []
    total = 0

    for r in range(2, ws.max_row + 1):
        d = normalize_date_cell(ws.cell(row=r, column=col_date).value)
        if d:
            dates.append(d)

        amt = normalize_amount_cell(ws.cell(row=r, column=col_amount).value)
        if amt is not None:
            total += amt

    newest_date = max(dates) if dates else None
    return newest_date, total


def core_name_from_filename(stem):
    """Extract the core middle name of the file.

    - remove leading 6-digit date if exists
    - remove trailing digits
    - strip spaces / dashes / underscores
    Example: '251106個人卡提(馬京瑋)133819' -> '個人卡提(馬京瑋)'
    """
    if re.match(r"^\d{6}", stem):
        stem = stem[6:]

    stem = re.sub(r"\d+$", "", stem)
    stem = stem.strip(" -_")

    return stem or "個人卡提"


# === MAIN PROCESSING FOR ONE WORKBOOK ===


def process_workbook(filepath):
    print(f"\n=== Processing: {os.path.basename(filepath)} ===")

    wb = openpyxl.load_workbook(filepath)
    if SHEET_NAME in wb.sheetnames:
        ws = wb[SHEET_NAME]
    else:
        ws = wb[wb.sheetnames[0]]

    headers = {cell.value: cell.column for cell in ws[1] if cell.value}
    col_date = headers.get("日期")
    col_id = headers.get("提款編號")
    col_amount = headers.get("提款金額")
    col_image = headers.get("提款截圖")

    if not all([col_date, col_id, col_amount, col_image]):
        print("  Skipped: Missing one of headers: 日期 / 提款編號 / 提款金額 / 提款截圖")
        return

    filled_rows = 0

    for img in getattr(ws, "_images", []):
        anchor = img.anchor
        mark = anchor._from  # 0-based row & col

        row_index = mark.row + 1
        col_index = mark.col + 1

        if col_index != col_image or row_index == 1:
            continue

        v_date = ws.cell(row=row_index, column=col_date).value
        v_id = ws.cell(row=row_index, column=col_id).value
        v_amt = ws.cell(row=row_index, column=col_amount).value
        if (v_date not in (None, "") and
            v_id not in (None, "") and
            v_amt not in (None, "")):
            continue

        pil_img = Image.open(img.ref)
        date_value, withdraw_id, amount = extract_fields_from_image(pil_img)

        if date_value and v_date in (None, ""):
            ws.cell(row=row_index, column=col_date).value = date_value
        if withdraw_id and v_id in (None, ""):
            ws.cell(row=row_index, column=col_id).value = str(withdraw_id)
        if amount is not None and v_amt in (None, ""):
            ws.cell(row=row_index, column=col_amount).value = amount

        filled_rows += 1
        print(f"  Row {row_index}: date={date_value}, id={withdraw_id}, amount={amount}")

    newest_date, total_amount = summarize_sheet(ws, col_date, col_amount)

    folder, filename = os.path.split(filepath)
    stem, ext = os.path.splitext(filename)

    if newest_date and total_amount > 0:
        date_prefix = newest_date.strftime("%y%m%d")
        core = core_name_from_filename(stem)
        new_stem = f"{date_prefix}{core}{total_amount}"
        new_name = new_stem + ext
    else:
        new_name = filename

    new_path = os.path.join(folder, new_name)

    wb.save(new_path)

    if new_name != filename:
        try:
            os.remove(filepath)
            print(f"  Saved as: {new_name} (original deleted)")
        except Exception as e:
            print(f"  Saved as: {new_name} (could not delete old file: {e})")
    else:
        print(f"  Saved (name unchanged): {filename}")

    print(f"  Filled rows: {filled_rows}, total amount: {total_amount}, newest date: {newest_date}")


# === ENTRY POINT: PROCESS ALL XLSX IN FOLDER ===


def main():
    xlsx_files = [
        f for f in os.listdir(BASE_DIR)
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    ]

    if not xlsx_files:
        print("No .xlsx files found in folder.")
        return

    print(f"Found {len(xlsx_files)} .xlsx file(s) in {BASE_DIR}")

    for fname in xlsx_files:
        full_path = os.path.join(BASE_DIR, fname)
        process_workbook(full_path)


if __name__ == "__main__":
    main()
