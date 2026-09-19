"""
Generic PDF text cleaner -> DOCX
Uses embedded text when present; optional OCR for image-only pages.
For documents you own or are licensed to process.
"""

from __future__ import annotations

import io
import re
from typing import List

import streamlit as st
from docx import Document
from docx.shared import Pt
from pypdf import PdfReader

BOILERPLATE_LINE = re.compile(
    r"""
    ^\s*(
        © |
        \\(c\\) |
        copyright |
        all\\s+rights\\s+reserved |
        page\\s+intentionally\\s+left\\s+blank |
        this\\s+page\\s+(is\\s+)?intentionally\\s+(left\\s+)?blank
    )
    """,
    re.IGNORECASE | re.VERBOSE,
)
PAGE_NUMBERISH = re.compile(r"^\\s*\\d{1,4}\\s*$")


def extract_text_layer(file_bytes: bytes) -> List[str]:
    reader = PdfReader(io.BytesIO(file_bytes))
    return [(page.extract_text() or "") for page in reader.pages]


def ocr_pages(file_bytes: bytes, max_pages: int, dpi: int) -> List[str]:
    from pdf2image import convert_from_bytes
    import pytesseract

    images = convert_from_bytes(
        file_bytes,
        dpi=dpi,
        first_page=1,
        last_page=max_pages,
        fmt="jpeg",
    )
    out: List[str] = []
    progress = st.progress(0.0, text="OCR running...")
    for i, img in enumerate(images, start=1):
        out.append(pytesseract.image_to_string(img) or "")
        progress.progress(i / len(images), text=f"OCR page {i}/{len(images)}")
    progress.empty()
    return out


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
    return ("\n".join(out).strip() + "\n") if out else ""


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
st.title("PDF -> cleaned DOCX")
st.caption(
    "Extracts selectable text, or OCRs image pages if you turn that on. "
    "Use only on files you own or are licensed to copy."
)

uploaded = st.file_uploader("Upload a PDF", type=["pdf"])
use_ocr = st.checkbox("Use OCR if the PDF has little or no text", value=False)
max_ocr_pages = st.slider("Max pages to OCR", 1, 50, 15)
ocr_dpi = st.select_slider("OCR DPI", options=[150, 200, 300], value=200)

if uploaded:
    data = uploaded.getvalue()
    text_pages = extract_text_layer(data)
    layer_chars = sum(len(p.strip()) for p in text_pages)
    st.metric("Pages in PDF", len(text_pages))
    st.metric("Text-layer characters", layer_chars)

    pages = text_pages
    if use_ocr and layer_chars < 200:
        try:
            pages = ocr_pages(data, max_pages=min(max_ocr_pages, len(text_pages)), dpi=ocr_dpi)
            st.success(f"OCR finished on {len(pages)} page(s).")
        except Exception as exc:
            st.error(f"OCR failed: {exc}")

    cleaned = clean_text("\n\n".join(pages))
    st.metric("Characters after clean", len(cleaned.strip()))

    with st.expander("Preview (first 4,000 characters)"):
        st.text(cleaned[:4000] if cleaned else "(empty)")

    if cleaned.strip():
        st.download_button(
            "Download DOCX",
            data=to_docx(cleaned),
            file_name="extracted_clean.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    else:
        st.warning("No text extracted. Enable OCR for image-only PDFs, or use a text PDF.")
else:
    st.info("Upload a PDF to extract and clean.")
