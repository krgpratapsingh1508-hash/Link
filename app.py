import io
import re
import traceback
from urllib.parse import urljoin

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

try:
    import pdfplumber
except Exception as e:
    pdfplumber = None
    PDF_IMPORT_ERROR = str(e)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                  "(KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/pdf,*/*",
}
DEFAULT_URL = "https://krg.ac.in/notice/355/"


def fetch(url, timeout=60):
    # verify=False: kuch college sites ka SSL certificate cloud par fail ho jata hai
    r = requests.get(url, headers=HEADERS, timeout=timeout, verify=False)
    r.raise_for_status()
    return r


def get_notice_info(url):
    soup = BeautifulSoup(fetch(url).text, "html.parser")

    title = ""
    for h1 in soup.find_all("h1"):
        t = h1.get_text(strip=True)
        if t and t.lower() != "notice viewer" and "college" not in t.lower():
            title = t
            break

    m = re.search(r"Published on\s+([^\n<]+)", soup.get_text("\n"))
    date = m.group(1).strip() if m else ""

    pdf_url = ""
    for a in soup.find_all("a", href=True):
        if ".pdf" in a["href"].lower() and "/media/" in a["href"].lower():
            pdf_url = urljoin(url, a["href"])
            break
    if not pdf_url:
        for a in soup.find_all("a", href=True):
            if ".pdf" in a["href"].lower():
                pdf_url = urljoin(url, a["href"])
                break
    return {"title": title, "date": date, "page_url": url, "pdf_url": pdf_url}


def extract_table(pdf_bytes):
    if pdfplumber is None:
        raise RuntimeError(f"pdfplumber import nahi hua: {PDF_IMPORT_ERROR}")
    rows = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            if tables:
                for table in tables:
                    for row in table:
                        rows.append([(c or "").replace("\n", " ").strip() for c in row])
            else:
                for line in (page.extract_text() or "").splitlines():
                    parts = re.split(r"\s{2,}", line.strip())
                    if parts and any(parts):
                        rows.append(parts)

    rows = [r for r in rows if any(c for c in r)]
    if not rows:
        raise RuntimeError("PDF me text/table nahi mili (scanned image PDF ho sakti hai).")

    width = max(len(r) for r in rows)
    rows = [r + [""] * (width - len(r)) for r in rows]
    header = rows[0]
    body = [r for r in rows[1:] if r != header]
    header = [h if h else f"Column {i + 1}" for i, h in enumerate(header)]
    return pd.DataFrame(body, columns=header)


def to_excel_bytes(df, info):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Roll List", index=False)
        pd.DataFrame(
            {"Field": ["Title", "Published on", "Notice page", "PDF link"],
             "Value": [info.get("title", ""), info.get("date", ""),
                       info.get("page_url", ""), info.get("pdf_url", "")]}
        ).to_excel(writer, sheet_name="Notice Info", index=False)
        for ws in writer.book.worksheets:
            for col in ws.columns:
                longest = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(longest + 2, 60)
    return buf.getvalue()


def show_result(df, info):
    st.success(f"{len(df)} rows mil gayi")
    st.dataframe(df, use_container_width=True)
    fname = re.sub(r"[^A-Za-z0-9._-]+", "_", info.get("title") or "notice") + ".xlsx"
    st.download_button(
        "⬇️ Excel download karo",
        data=to_excel_bytes(df, info),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


st.set_page_config(page_title="Notice to Excel", page_icon="📄")
st.title("📄 KRG Notice → Excel")

tab1, tab2 = st.tabs(["🔗 Link se", "📤 PDF upload karke"])

with tab1:
    url = st.text_input("Notice link", value=DEFAULT_URL)
    if st.button("Excel banao", type="primary"):
        step = "start"
        try:
            step = "notice page fetch"
            info = get_notice_info(url.strip())
            st.write("**Title:**", info["title"] or "(nahi mila)")
            st.write("**Date:**", info["date"] or "(nahi mili)")
            st.write("**PDF link:**", info["pdf_url"] or "(nahi mila)")
            if not info["pdf_url"]:
                st.error("PDF link nahi mila. Dusre tab me PDF upload karke try kijiye.")
            else:
                step = "PDF download"
                pdf_bytes = fetch(info["pdf_url"]).content
                st.write(f"PDF size: {len(pdf_bytes) // 1024} KB")
                step = "table extract"
                df = extract_table(pdf_bytes)
                show_result(df, info)
        except Exception as e:
            st.error(f"Step '{step}' me error: {e}")
            st.code(traceback.format_exc())
            st.info("Agar site cloud se block ho rahi hai, to dusre tab me PDF khud upload kar dijiye.")

with tab2:
    up = st.file_uploader("Roll list PDF upload kijiye", type=["pdf"])
    if up is not None:
        try:
            df = extract_table(up.read())
            show_result(df, {"title": up.name.rsplit(".", 1)[0]})
        except Exception as e:
            st.error(f"Error: {e}")
            st.code(traceback.format_exc())
