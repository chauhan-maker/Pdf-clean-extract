"""
Generic PDF text cleaner -> DOCX
For user-owned / licensed documents only.
Does not implement exam-bank structuring (Q/A/justification filters).
"""

from __future__ import annotations

import io
import re
from typing import List

import streamlit as st
from docx import Document
from docx.shared import Pt
from pypdf import PdfReader

# Lines that look like publisher boilerplate / running headers-footers.
# Conservative: whole-line matches only so body sentences are kept.
BOILERPLATE_LINE = re.compile(
    r"""
    ^\s*(
        © |
        \(c\) |
        copyright |
        all\s+rights\s+reserved |
        page\s+intentionally\s+left\s+blank |
        this\s+page\s+(is\s+)?intentionally\s+(left\s+)?blank
    )
    """,
    re.IGNORECASE | re.VERBOSE,
)

PAGE_NUMBERISH = re.compile(r"^\s*\d{1,4}\s*$")


def extract_pages(file_bytes: bytes) -> List[str]:
    reader = PdfReader(io.BytesIO(file_bytes))
    pages = []
    for page in reader.pages:
        text = page.extract_text() or ""
        pages.append(text)
    return pages


def clean_text(raw: str) -> str:
    kept: List[str] = []
    for line in raw.splitlines():
        s = line.strip()
        if not s:
            if kept and kept[-1] != "":
                kept.append("")
            continue
        if BOILERPLATE_LINE.search(s):
            continue
        if PAGE_NUMBERISH.match(s):
            continue
        kept.append(line.rstrip())
    # collapse 3+ blank lines
    out: List[str] = []
    blanks = 0
    for line in kept:
        if line == "":
            blanks += 1
            if blanks <= 2:
                out.append("")
        else:
            blanks = 0
            out.append(line)
    return "\n".join(out).strip() + "\n"


def to_docx(text: str) -> bytes:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    for para in text.split("\n"):
        doc.add_paragraph(para)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


st.set_page_config(page_title="PDF text cleaner", layout="wide")
st.title("PDF → cleaned DOCX")
st.caption(
    "Extracts text from a PDF you own or are licensed to use, "
    "drops whole lines that look like copyright / reserved / blank-page notices, "
    "and builds a simple DOCX. Not for reproducing copyrighted exam banks."
)

uploaded = st.file_uploader("Upload a PDF", type=["pdf"])

if uploaded:
    raw_pages = extract_pages(uploaded.getvalue())
    joined = "\n\n".join(raw_pages)
    cleaned = clean_text(joined)

    st.metric("Pages read", len(raw_pages))
    st.metric("Characters after clean", len(cleaned))

    with st.expander("Preview (first 4,000 characters)"):
        st.text(cleaned[:4000])

    docx_bytes = to_docx(cleaned)
    st.download_button(
        "Download DOCX",
        data=docx_bytes,
        file_name="extracted_clean.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
else:
    st.info("Upload a PDF to extract and clean.")
  
