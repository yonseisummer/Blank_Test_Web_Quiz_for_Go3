import io
import re

import streamlit as st
from PyPDF2 import PdfReader
from docx import Document


st.set_page_config(
    page_title="PDF to DOCX Converter",
    page_icon="📄",
    layout="centered",
)


def clean_xml_text(text: str | None) -> str:
    """
    Remove characters that are illegal in XML/Word documents.
    Keeps normal text, tabs, newlines, and carriage returns.
    """
    if text is None:
        return ""

    return re.sub(
        r"[\x00-\x08\x0B\x0C\x0E-\x1F]",
        "",
        text,
    )


def pdf_to_docx_simple(pdf_file) -> io.BytesIO:
    """
    Convert uploaded PDF file to DOCX.
    This keeps the original notebook's logic as much as possible:
    - extract text page by page
    - split extracted text into lines
    - add each line as a paragraph
    - add page breaks between PDF pages
    """
    reader = PdfReader(pdf_file, strict=False)
    doc = Document()

    for i, page in enumerate(reader.pages):
        text = page.extract_text()

        if text:
            text = clean_xml_text(text)

            for line in text.splitlines():
                line = clean_xml_text(line)

                if line.strip():
                    doc.add_paragraph(line)
                else:
                    doc.add_paragraph()

        if i < len(reader.pages) - 1:
            doc.add_page_break()

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output


st.title("PDF to DOCX Converter")
st.write("PDF 파일을 업로드하면 텍스트를 추출해 DOCX 파일로 변환합니다.")

uploaded_file = st.file_uploader("PDF 파일을 선택하세요", type=["pdf"])

if uploaded_file is not None:
    output_filename = uploaded_file.name.rsplit(".", 1)[0] + ".docx"

    if st.button("DOCX로 변환하기"):
        try:
            with st.spinner("변환 중입니다..."):
                docx_file = pdf_to_docx_simple(uploaded_file)

            st.success("변환 완료!")
            st.download_button(
                label="DOCX 파일 다운로드",
                data=docx_file,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")
else:
    st.info("먼저 PDF 파일을 업로드해주세요.")
