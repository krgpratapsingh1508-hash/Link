"""
notice_to_excel.py
------------------
KRG College notice page (jaise https://krg.ac.in/notice/355/) se:
  1. Notice ka title aur date nikalta hai
  2. Notice ka PDF download karta hai
  3. PDF ke andar ki table (roll list) ko Excel (.xlsx) me convert karta hai

Install:
    pip install requests beautifulsoup4 pdfplumber pandas openpyxl

Run:
    python notice_to_excel.py
    python notice_to_excel.py https://krg.ac.in/notice/355/
    python notice_to_excel.py https://krg.ac.in/notice/355/ -o roll_list.xlsx
"""

import argparse
import io
import re
import sys
from urllib.parse import urljoin

import pandas as pd
import pdfplumber
import requests
from bs4 import BeautifulSoup

DEFAULT_URL = "https://krg.ac.in/notice/355/"
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; KRG-Notice-Scraper/1.0)"}


# ---------------------------------------------------------------- step 1
def get_notice_info(url):
    """Notice page se title, date aur PDF link nikalo."""
    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    # Title: "Notice Viewer" ke baad wala h1
    title = ""
    for h1 in soup.find_all("h1"):
        text = h1.get_text(strip=True)
        if text and text.lower() != "notice viewer" and "College" not in text:
            title = text
            break

    # Date: "Published on 05 Oct, 2026"
    date = ""
    m = re.search(r"Published on\s+([^\n<]+)", soup.get_text("\n"))
    if m:
        date = m.group(1).strip()

    # PDF link
    pdf_url = ""
    for a in soup.find_all("a", href=True):
        if a["href"].lower().endswith(".pdf") and "/media/notices/" in a["href"]:
            pdf_url = urljoin(url, a["href"])
            break
    if not pdf_url:
        raise RuntimeError("Is notice page par PDF link nahi mila.")

    return {"title": title, "date": date, "page_url": url, "pdf_url": pdf_url}


# ---------------------------------------------------------------- step 2
def download_pdf(pdf_url):
    resp = requests.get(pdf_url, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    return resp.content


# ---------------------------------------------------------------- step 3
def extract_table(pdf_bytes):
    """PDF ki saari pages se table rows nikalo. Table na mile to text lines se try karo."""
    rows = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table:
                        rows.append([(c or "").replace("\n", " ").strip() for c in row])
            else:
                # Fallback: text line ko 2+ spaces par todo
                text = page.extract_text() or ""
                for line in text.splitlines():
                    parts = re.split(r"\s{2,}", line.strip())
                    if parts and any(parts):
                        rows.append(parts)

    # Khali rows hatao
    rows = [r for r in rows if any(cell for cell in r)]
    if not rows:
        raise RuntimeError(
            "PDF me text/table nahi mili. Shayad scanned image PDF hai (OCR chahiye)."
        )

    # Sab rows ko same width do
    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]

    # Pehli row ko header maano; baar-baar aane wali header rows (har page par) hata do
    header = rows[0]
    body = [r for r in rows[1:] if r != header]
    header = [h if h else f"Column {i + 1}" for i, h in enumerate(header)]
    return pd.DataFrame(body, columns=header)


# ---------------------------------------------------------------- step 4
def save_excel(df, info, out_path):
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Roll List", index=False)
        pd.DataFrame(
            {
                "Field": ["Title", "Published on", "Notice page", "PDF link"],
                "Value": [info["title"], info["date"], info["page_url"], info["pdf_url"]],
            }
        ).to_excel(writer, sheet_name="Notice Info", index=False)

        # Column width auto-fit
        for ws in writer.book.worksheets:
            for col in ws.columns:
                longest = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(longest + 2, 60)


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description="KRG notice PDF -> Excel")
    ap.add_argument("url", nargs="?", default=DEFAULT_URL)
    ap.add_argument("-o", "--output", default="")
    args = ap.parse_args()

    print("Notice page fetch ho raha hai...")
    info = get_notice_info(args.url)
    print(f"  Title : {info['title']}\n  Date  : {info['date']}\n  PDF   : {info['pdf_url']}")

    print("PDF download ho raha hai...")
    pdf_bytes = download_pdf(info["pdf_url"])

    print("Table nikali ja rahi hai...")
    df = extract_table(pdf_bytes)

    out = args.output or re.sub(r"[^A-Za-z0-9._-]+", "_", info["title"] or "notice") + ".xlsx"
    save_excel(df, info, out)
    print(f"Done! {len(df)} rows ke saath Excel ban gaya: {out}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
