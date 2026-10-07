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
    r = requests.get(url, headers=HEADERS, timeout=timeout, verify=False)
    r.raise_for_status()
    return r


def get_notice_info(url):
    html = fetch(url).text
    soup = BeautifulSoup(html, "html.parser")

    title = ""
    for h1 in soup.find_all("h1"):
        t = h1.get_text(strip=True)
        if t and t.lower() != "notice viewer" and "college" not in t.lower():
            title = t
            break

    m = re.search(r"Published on\s+([^\n<]+)", soup.get_text("\n"))
    date = m.group(1).strip() if m else ""

    pdf_url = ""
    found = re.findall(r"""["'(=\s]([^"'\s()<>]*?/media/notices/[^"'\s()<>]*?\.pdf)""", html, flags=re.I)
    if found:
        pdf_url = urljoin(url, found[0])
    else:
        for tag, attr in (("a", "href"), ("iframe", "src"), ("embed", "src"), ("object", "data")):
            for el in soup.find_all(tag):
                link = el.get(attr, "")
                if ".pdf" in link.lower() and "/static/images/" not in link.lower():
                    pdf_url = urljoin(url, link)
                    break
            if pdf_url:
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
    # duplicate column names ko alag karo
    seen = {}
    uniq = []
    for h in header:
        seen[h] = seen.get(h, 0) + 1
        uniq.append(h if seen[h] == 1 else f"{h}_{seen[h]}")
    return pd.DataFrame(body, columns=uniq)


def add_title(df, title):
    df = df.copy()
    df["Title"] = title
    return df


def to_excel_bytes(df, status_df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="All Data", index=False)
        status_df.to_excel(writer, sheet_name="Notice Info", index=False)
        for ws in writer.book.worksheets:
            for col in ws.columns:
                longest = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(longest + 2, 60)
    return buf.getvalue()


def show_result(frames, status_rows):
    status_df = pd.DataFrame(status_rows)
    st.subheader("Status")
    st.dataframe(status_df, use_container_width=True)
    if not frames:
        st.error("Kisi bhi link/PDF se data nahi mila.")
        return
    big = pd.concat(frames, ignore_index=True).fillna("")
    st.success(f"Total {len(big)} rows, {len(frames)} notice se")
    st.dataframe(big, use_container_width=True)
    st.download_button(
        "⬇️ Excel download karo",
        data=to_excel_bytes(big, status_df),
        file_name="notices_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


st.set_page_config(page_title="Notice to Excel", page_icon="📄")
st.title("📄 KRG Notice → Excel")

tab1, tab2 = st.tabs(["🔗 Links se", "📤 PDF upload karke"])

with tab1:
    links_text = st.text_area(
        "Notice links (har line me ek link, 10 ya zyada bhi chalenge)",
        value=DEFAULT_URL,
        height=200,
    )
    if st.button("Excel banao", type="primary"):
        urls = [u.strip() for u in links_text.splitlines() if u.strip()]
        frames, status_rows = [], []
        bar = st.progress(0.0)
        for i, u in enumerate(urls, 1):
            row = {"Link": u, "Title": "", "Rows": 0, "Result": ""}
            try:
                info = get_notice_info(u)
                row["Title"] = info["title"]
                if not info["pdf_url"]:
                    raise RuntimeError("PDF link nahi mila")
                df = extract_table(fetch(info["pdf_url"]).content)
                frames.append(add_title(df, info["title"]))
                row["Rows"] = len(df)
                row["Result"] = "OK"
            except Exception as e:
                row["Result"] = f"Error: {e}"
            status_rows.append(row)
            bar.progress(i / len(urls))
        show_result(frames, status_rows)

with tab2:
    ups = st.file_uploader("Roll list PDF upload kijiye (ek ya zyada)", type=["pdf"], accept_multiple_files=True)
    if ups and st.button("PDF se Excel banao", type="primary"):
        frames, status_rows = [], []
        for up in ups:
            title = up.name.rsplit(".", 1)[0]
            row = {"Link": up.name, "Title": title, "Rows": 0, "Result": ""}
            try:
                df = extract_table(up.read())
                frames.append(add_title(df, title))
                row["Rows"] = len(df)
                row["Result"] = "OK"
            except Exception as e:
                row["Result"] = f"Error: {e}"
            status_rows.append(row)
        show_result(frames, status_rows)
