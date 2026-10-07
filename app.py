import io
import os
import re
import tempfile
import traceback
import zipfile
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

try:
    from pdf2docx import Converter
except Exception as e:
    Converter = None
    PDF2DOCX_ERROR = str(e)

try:
    from docx import Document
except Exception as e:
    Document = None
    DOCX_ERROR = str(e)

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


def _norm(x):
    return re.sub(r"[^a-z0-9\u0900-\u097f]", "", str(x).lower())


def std_header(h):
    """Alag-alag PDF ke header ko ek jaise naam do, taki combine karne par data ek hi column me aaye."""
    n = _norm(h)
    if not n:
        return ""
    if n in ("sno", "srno", "slno", "serialno", "serialnumber", "sr", "sl", "no", "क्र", "क्रमांक", "क्रसं", "क्रम") \
            or n.startswith(("sno", "srno", "slno", "serialno", "क्र")):
        return "S. No."
    if "father" in n or "mother" in n or "guardian" in n or "पिता" in n or "माता" in n:
        return "Father/Mother Name"
    if "name" in n or "नाम" in n:
        return "Student Name"
    if "enrol" in n or "enrl" in n:
        return "Enrollment No."
    if "roll" in n or "रोल" in n:
        return "Roll No."
    return str(h).strip()


def _is_header_row(row):
    hit = {std_header(c) for c in row}
    return "Student Name" in hit or ("S. No." in hit and len(hit & {"Roll No.", "Enrollment No.", "Father/Mother Name"}) > 0)


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

    # Asli header row dhundo (title/college ke naam wali upar ki rows ko chhodo)
    h_idx = next((i for i, r in enumerate(rows[:25]) if _is_header_row(r)), None)
    if h_idx is not None:
        header = [std_header(c) for c in rows[h_idx]]
        body = rows[h_idx + 1:]
    elif rows[0][0].strip().isdigit():
        # Header hai hi nahi, seedha data: pehla column S. No., doosra Student Name maano
        header = ["S. No.", "Student Name"] + [f"Column {i + 1}" for i in range(2, width)]
        body = rows
    else:
        header = [std_header(c) for c in rows[0]]
        body = rows[1:]

    header = [h if h else f"Column {i + 1}" for i, h in enumerate(header)]
    seen, uniq = {}, []
    for h in header:
        seen[h] = seen.get(h, 0) + 1
        uniq.append(h if seen[h] == 1 else f"{h}_{seen[h]}")

    # Har page par dobara aane wali header rows hatao
    hn = [_norm(std_header(c)) for c in header]
    body = [r for r in body if [_norm(std_header(c)) for c in r] != hn]
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


def safe_name(text, default="notice"):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", text or default).strip("_") or default


def pdf_to_docx_bytes(pdf_bytes):
    """Original PDF jaisa layout rakhte hue PDF ko Word (.docx) me badlo."""
    if Converter is None:
        raise RuntimeError(f"pdf2docx install nahi hai: {PDF2DOCX_ERROR}")
    with tempfile.TemporaryDirectory() as tmp:
        pdf_path = os.path.join(tmp, "in.pdf")
        docx_path = os.path.join(tmp, "out.docx")
        with open(pdf_path, "wb") as f:
            f.write(pdf_bytes)
        cv = Converter(pdf_path)
        try:
            cv.convert(docx_path)
        finally:
            cv.close()
        with open(docx_path, "rb") as f:
            return f.read()


def tables_to_docx_bytes(frames_with_titles):
    """Har notice ki table ko Word table me daalo (title heading ke saath), sab ek hi file me."""
    if Document is None:
        raise RuntimeError(f"python-docx install nahi hai: {DOCX_ERROR}")
    doc = Document()
    for n, (title, df) in enumerate(frames_with_titles):
        if n:
            doc.add_page_break()
        doc.add_heading(title or "Notice", level=1)
        cols = [c for c in df.columns if c != "Title"]
        table = doc.add_table(rows=1, cols=len(cols))
        table.style = "Table Grid"
        for i, c in enumerate(cols):
            table.rows[0].cells[i].text = str(c)
            for p in table.rows[0].cells[i].paragraphs:
                for r in p.runs:
                    r.bold = True
        for _, row in df.iterrows():
            cells = table.add_row().cells
            for i, c in enumerate(cols):
                cells[i].text = str(row[c])
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def make_zip(files):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files:
            z.writestr(name, data)
    return buf.getvalue()


def download_pdfs_from_links(text):
    """Links se PDF download karo. Return: [(title, pdf_bytes)], status_rows"""
    urls = [u.strip() for u in text.splitlines() if u.strip()]
    pdfs, status = [], []
    bar = st.progress(0.0)
    for i, u in enumerate(urls, 1):
        row = {"Link": u, "Title": "", "Result": ""}
        try:
            info = get_notice_info(u)
            row["Title"] = info["title"]
            if not info["pdf_url"]:
                raise RuntimeError("PDF link nahi mila")
            pdfs.append((info["title"] or f"notice_{i}", fetch(info["pdf_url"]).content))
            row["Result"] = "PDF mil gayi"
        except Exception as e:
            row["Result"] = f"Error: {e}"
        status.append(row)
        bar.progress(i / len(urls))
    return pdfs, status


def convert_and_offer(pdfs, status, zip_name):
    """PDFs ko Word me badlo aur download button do (ek ho to .docx, zyada ho to .zip)."""
    outputs = []
    for title, data in pdfs:
        row = next((r for r in status if r.get("Title") == title or r.get("Link") == title), None)
        try:
            outputs.append((safe_name(title) + ".docx", pdf_to_docx_bytes(data)))
            if row is not None:
                row["Result"] = "Word ban gaya"
        except Exception as e:
            if row is not None:
                row["Result"] = f"Error: {e}"
            else:
                status.append({"Link": title, "Title": title, "Result": f"Error: {e}"})
    st.subheader("Status")
    st.dataframe(pd.DataFrame(status), use_container_width=True)
    DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    if not outputs:
        st.error("Koi Word file nahi bani.")
    elif len(outputs) == 1:
        st.download_button("⬇️ Word download karo", data=outputs[0][1], file_name=outputs[0][0], mime=DOCX_MIME)
    else:
        st.download_button("⬇️ Sab Word files (ZIP) download karo", data=make_zip(outputs),
                           file_name=zip_name, mime="application/zip")


st.set_page_config(page_title="Notice to Excel", page_icon="📄")
st.title("📄 KRG Notice → Excel")

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🔗 Links → Excel", "📤 PDF → Excel",
    "📝 Links → Word", "📝 PDF → Word", "📋 Table → Word",
])

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

with tab3:
    st.write("Notice links se PDF lekar **Word (.docx)** banao (PDF jaisa layout).")
    links3 = st.text_area("Notice links (har line me ek)", value=DEFAULT_URL, height=200, key="links3")
    if st.button("Word banao", type="primary", key="btn3"):
        pdfs, status = download_pdfs_from_links(links3)
        with st.spinner("Word me convert ho raha hai..."):
            convert_and_offer(pdfs, status, "notices_word.zip")

with tab4:
    st.write("PDF upload karke **Word (.docx)** banao (PDF jaisa layout).")
    ups4 = st.file_uploader("PDF upload kijiye (ek ya zyada)", type=["pdf"], accept_multiple_files=True, key="up4")
    if ups4 and st.button("PDF se Word banao", type="primary", key="btn4"):
        pdfs = [(u.name.rsplit(".", 1)[0], u.read()) for u in ups4]
        status = [{"Link": t, "Title": t, "Result": ""} for t, _ in pdfs]
        with st.spinner("Word me convert ho raha hai..."):
            convert_and_offer(pdfs, status, "pdf_to_word.zip")

with tab5:
    st.write("Roll list ki **table** ko saaf Word table me badlo. Saare notice **ek hi Word file** me, har notice alag page par, title heading ke saath.")
    links5 = st.text_area("Notice links (har line me ek, optional)", value="", height=150, key="links5")
    ups5 = st.file_uploader("Ya PDF upload kijiye", type=["pdf"], accept_multiple_files=True, key="up5")
    if st.button("Table Word me banao", type="primary", key="btn5"):
        items, status = [], []
        if links5.strip():
            pdfs, status = download_pdfs_from_links(links5)
            items += pdfs
        items += [(u.name.rsplit(".", 1)[0], u.read()) for u in (ups5 or [])]
        results = []
        for title, data in items:
            try:
                results.append((title, extract_table(data)))
            except Exception as e:
                status.append({"Link": title, "Title": title, "Result": f"Error: {e}"})
        if status:
            st.dataframe(pd.DataFrame(status), use_container_width=True)
        if results:
            st.success(f"{len(results)} notice ki table taiyar")
            st.download_button(
                "⬇️ Word download karo",
                data=tables_to_docx_bytes(results),
                file_name="notices_tables.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        else:
            st.error("Koi table nahi mili.")
