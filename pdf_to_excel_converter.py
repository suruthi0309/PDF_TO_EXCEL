#!/usr/bin/env python3
"""
Bank PDF -> Excel/CSV converter (Auto-column detection + formatting + CSV)
Corrected version: avoids duplicate extraction and removes header rows.
"""

import os
import re
import math
import pdfplumber
import pandas as pd
from tqdm import tqdm
from dateutil import parser as dateparser
from datetime import datetime
import time

# --- Optional OCR libs
try:
    from pdf2image import convert_from_path
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# --- Excel styling
try:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment
except ImportError:
    load_workbook = None

# Optional process control to close Excel if needed
try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

from tkinter import Tk, filedialog, simpledialog

SECTION_HEADERS = [
    "Account Summary", "Deposits & Other Credits", "Deposits",
    "ATM Withdrawals & Debits", "ATM Withdrawals",
    "Withdrawals & Other Debits", "Checks Paid", "Checks",
    "Payments", "Transactions"
]

DATE_RE = re.compile(r'(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}[/-]\d{1,2}[/-]\d{1,2})')
AMOUNT_RE = re.compile(r'(-?[\₹\$\£]?\s?[\d,]+(?:\.\d{1,2})?)')

def select_folder():
    root = Tk()
    root.withdraw()
    return filedialog.askdirectory(title="Select Folder Containing PDFs")

def get_pdf_password(filename):
    return simpledialog.askstring("Password Required", f"Enter password for: {filename}", show='*')

def try_parse_date(text):
    if not isinstance(text, str):
        return None
    text = text.strip()
    m = DATE_RE.search(text)
    if m:
        try:
            return dateparser.parse(m.group(1), dayfirst=True, fuzzy=True).date()
        except:
            pass
    try:
        return dateparser.parse(text, dayfirst=True, fuzzy=True).date()
    except:
        return None

def try_parse_amount(text):
    if text is None:
        return None
    text = str(text).replace('(', '-').replace(')', '').replace('CR', '').replace('Dr', '-')
    m = AMOUNT_RE.search(text.replace(' ', ''))
    if not m:
        m = AMOUNT_RE.search(text)
    if not m:
        return None
    try:
        return float(re.sub(r'[^\d\.\-]', '', m.group(1)))
    except:
        return None

def normalize_amount_sign(raw, amount):
    if amount is None:
        return None
    text = raw.lower()
    negatives = ['withdrawal', 'debit', 'paid', 'purchase', 'atm', 'payment', 'check']
    positives = ['deposit', 'credit', 'interest', 'refund']
    if any(k in text for k in negatives):
        return -abs(amount)
    if any(k in text for k in positives):
        return abs(amount)
    return amount

def apply_ocr(pdf_path):
    if not OCR_AVAILABLE:
        return None
    try:
        pages = convert_from_path(pdf_path)
        lines = []
        for img in pages:
            text = pytesseract.image_to_string(img)
            lines.extend([l.strip() for l in text.splitlines() if l.strip()])
        return list(set(lines))
    except:
        return None

def extract_text_rows_from_pdf(pdf_path, password=None):
    """
    ✅ Only extract text if no table exists. Remove duplicates.
    """
    lines = []
    try:
        with pdfplumber.open(pdf_path, password=password) as pdf:
            for page in pdf.pages:
                page_unique = set()

                table = page.extract_table()
                if table:
                    for row in table:
                        joined = "  ".join([(str(c) if c else "").strip() for c in row])
                        if joined.strip():
                            page_unique.add(joined.strip())
                else:
                    text = page.extract_text()
                    if text:
                        for l in text.splitlines():
                            if l.strip():
                                page_unique.add(l.strip())

                lines.extend(sorted(page_unique))

    except Exception:
        ocr = apply_ocr(pdf_path)
        return ocr if ocr else []

    return list(dict.fromkeys(lines)) or apply_ocr(pdf_path) or []

def detect_sections(lines):
    sections = {}
    current = "Uncategorized"
    for line in lines:
        for h in SECTION_HEADERS:
            if h.lower() in line.lower():
                current = h
                break
        sections.setdefault(current, []).append(line)
    return sections

def rows_from_section_lines(lines):
    rows = []
    for line in lines:
        parts = re.split(r'\s{2,}', line) or [line]
        rows.append({"raw": line, "parts": parts})
    return rows

def infer_row_columns(row):
    raw = row["raw"]
    parts = row["parts"]
    date = None
    amount = None

    for p in parts:
        if not date:
            date = try_parse_date(p)
        if amount is None:
            amount = try_parse_amount(p)

    date = date or try_parse_date(raw)
    amount = amount or try_parse_amount(raw)

    signed_amt = normalize_amount_sign(raw, amount)

    return {
        "Date": date,
        "Description": re.sub(r'\s{2,}', ' ', raw),
        "Amount": signed_amt,
        "Raw": raw
    }

# ✅ ✅ Updated function to remove header rows
def build_master_rows(sections, src):
    master = []
    skip_keywords = [
        "date", "description", "debit", "credit", "balance",
        "invoice", "amount", "customer", "check", "paid", "account",
        "statement", "summary", "opening", "closing"
    ]

    for sec, lines in sections.items():
        for r in rows_from_section_lines(lines):
            entry = infer_row_columns(r)
            raw_lower = entry["Raw"].lower()

            # ✅ Skip if no date and no amount + looks like a header
            if entry["Date"] is None and entry["Amount"] is None:
                if any(k in raw_lower for k in skip_keywords):
                    continue

            entry["Section"] = sec
            entry["Source File"] = src
            master.append(entry)

    return master

def safe_create_excel_writer(path):
    try:
        return pd.ExcelWriter(path, engine="openpyxl"), path
    except PermissionError:
        if PSUTIL_AVAILABLE:
            for p in psutil.process_iter(['name']):
                if p.info['name'] and "EXCEL" in p.info['name'].upper():
                    p.terminate()
            time.sleep(1)
        base, ext = os.path.splitext(path)
        new = f"{base}_{int(time.time())}{ext}"
        return pd.ExcelWriter(new, engine="openpyxl"), new

def main():
    folder = select_folder()
    if not folder:
        print("No folder selected.")
        return

    pdfs = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]
    if not pdfs:
        print("No PDFs found.")
        return

    out_xlsx = os.path.join(folder, "Merged_Extracted_Data.xlsx")
    out_csv = os.path.join(folder, "Merged_Master_Data.csv")

    writer, real_path = safe_create_excel_writer(out_xlsx)
    master = []
    summary = []

    for file in tqdm(pdfs, desc="Processing PDFs"):
        path = os.path.join(folder, file)
        password = get_pdf_password(file) if "bank" in file.lower() else None

        lines = extract_text_rows_from_pdf(path, password)
        print(f"{file}: {len(lines)} unique lines extracted ✅")

        sections = detect_sections(lines)
        rows = build_master_rows(sections, file)

        df = pd.DataFrame(rows)
        if not df.empty:
            df["Date"] = df["Date"].astype(str)
            df.to_excel(writer, sheet_name=file[:31], index=False)
            master.append(df)

        summary.append({"File": file, "Rows": len(df), "Status": "OK"})

    pd.DataFrame(summary).to_excel(writer, sheet_name="Summary", index=False)

    if master:
        combined = pd.concat(master, ignore_index=True)
        combined.to_excel(writer, sheet_name="Master", index=False)
        combined.to_csv(out_csv, index=False)

    writer.close()
    print(f"✅ Excel Saved: {real_path}")
    print(f"✅ CSV Saved: {out_csv}")
    print("🎉 Done!")

if __name__ == "__main__":
    main()
