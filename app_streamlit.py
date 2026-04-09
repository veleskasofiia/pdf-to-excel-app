import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO
from PyPDF2 import PdfMerger
from docx import Document
from fpdf import FPDF
import tempfile
import os

# ✅ Config
st.set_page_config(page_title="PDF Tools", page_icon="📊", layout="wide")

st.markdown("<h1 style='text-align:center;'>PDF Tools</h1>", unsafe_allow_html=True)

# 🔥 Tabs
tab1, tab2, tab3 = st.tabs(["📊 PDF to Excel", "📄 Word to PDF", "🔗 Merge PDFs"])

# =========================================================
# 📊 PDF → EXCEL (YOUR WORKING LOGIC)
# =========================================================
with tab1:

    uploaded_file = st.file_uploader("Upload PDF", type="pdf")
    convert = st.button("Convert to Excel")

    if uploaded_file and convert:
        data = []

        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                lines = text.split("\n")
                current_item = None
                current_descriptions = []

                for line in lines:
                    line = line.strip()
                    if not line:
                        continue

                    parts = line.split()

                    if len(parts) > 5 and parts[0].isdigit():

                        if current_item:
                            data.append(current_item + current_descriptions)

                        try:
                            line_no, date, item_id = parts[0], parts[1], parts[2]
                            qty = parts[-4]
                            unit = parts[-3]
                            price = parts[-2].replace("$","").replace(",","")
                            amount = parts[-1].replace("$","").replace(",","")

                            description = " ".join(parts[3:-4])

                            current_item = [line_no, date, item_id, qty, unit, price, amount]
                            current_descriptions = [description] if description else []

                        except:
                            current_item = None
                            current_descriptions = []

                    else:
                        if current_item:
                            current_descriptions.append(line)

                if current_item:
                    data.append(current_item + current_descriptions)

        if data:
            max_desc = max(len(r) - 7 for r in data)
            columns = ["Line","Date","ItemID","Qty","Unit","Price","Amount"]
            columns += [f"Desc{i+1}" for i in range(max_desc)]

            for r in data:
                while len(r) < len(columns):
                    r.append("")

            df = pd.DataFrame(data, columns=columns)

            st.data_editor(df, use_container_width=True, height=600)

            output = BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)

            st.download_button("Download Excel", output, "output.xlsx")

# =========================================================
# 📄 WORD → PDF
# =========================================================
with tab2:

    from docx import Document
    from fpdf import FPDF

    uploaded_docx = st.file_uploader("Upload Word file", type="docx")

    if uploaded_docx:
        if st.button("Convert to PDF"):

            doc = Document(uploaded_docx)

            pdf = FPDF()
            pdf.add_page()

            # ✅ LOCAL FONT (IMPORTANT)
            pdf.add_font("DejaVu", "", "DejaVuSans.ttf", uni=True)
            pdf.set_font("DejaVu", size=12)

            for para in doc.paragraphs:
                text = para.text.strip()
                if text:
                    pdf.multi_cell(0, 8, text)

            pdf_output = BytesIO()
            pdf.output(pdf_output)
            pdf_output.seek(0)

            st.download_button(
                "⬇️ Download PDF",
                pdf_output,
                "converted.pdf",
                mime="application/pdf"
            )

            st.success("✅ Conversion successful!")

# =========================================================
# 🔗 MERGE PDF
# =========================================================
with tab3:

    uploaded_pdfs = st.file_uploader("Upload multiple PDFs", type="pdf", accept_multiple_files=True)

    if uploaded_pdfs:
        if st.button("Merge PDFs"):

            merger = PdfMerger()

            for pdf_file in uploaded_pdfs:
                merger.append(pdf_file)

            output = BytesIO()
            merger.write(output)
            merger.close()
            output.seek(0)

            st.download_button(
                "Download merged PDF",
                output,
                "merged.pdf",
                mime="application/pdf"
            )