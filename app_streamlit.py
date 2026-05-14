import streamlit as st
import pdfplumber
import pandas as pd
import zipfile
import io
import time

# Page config
st.set_page_config(page_title="PO to Excel", layout="centered")

# Custom styling
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.block-container {
    padding-top: 2.5rem;
    padding-bottom: 2.5rem;
    max-width: 720px;
}

/* Header */
.app-header {
    text-align: center;
    margin-bottom: 2rem;
}
.app-title {
    font-size: 2.2rem;
    font-weight: 700;
    color: #1a1a2e;
    letter-spacing: -0.5px;
    margin-bottom: 0.3rem;
}
.app-title span {
    color: #e63946;
}
.app-subtitle {
    color: #6b7280;
    font-size: 1rem;
    margin-top: 0;
}

/* Card wrapper */
.card {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 14px;
    padding: 1.8rem 2rem;
    margin-bottom: 1.2rem;
    box-shadow: 0 1px 4px rgba(0,0,0,0.05);
}

/* Section label */
.section-label {
    font-size: 0.78rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: #9ca3af;
    margin-bottom: 0.6rem;
}

/* Convert button */
.stButton > button {
    background: linear-gradient(135deg, #e63946 0%, #c1121f 100%);
    color: white;
    border: none;
    border-radius: 10px;
    height: 52px;
    width: 100%;
    font-size: 1rem;
    font-weight: 600;
    letter-spacing: 0.02em;
    transition: opacity 0.15s;
    cursor: pointer;
}
.stButton > button:hover {
    opacity: 0.88;
}

/* Processing banner */
.processing-bar {
    background: #fff7ed;
    border: 1px solid #fed7aa;
    border-radius: 10px;
    padding: 0.9rem 1.2rem;
    color: #c2410c;
    font-size: 0.95rem;
    font-weight: 500;
    text-align: center;
    margin: 0.8rem 0;
}

/* Metrics row */
.metrics-row {
    display: flex;
    gap: 1rem;
    margin-bottom: 1rem;
}
.metric-box {
    flex: 1;
    background: #f9fafb;
    border: 1px solid #e5e7eb;
    border-radius: 10px;
    padding: 0.8rem 1rem;
    text-align: center;
}
.metric-value {
    font-size: 1.5rem;
    font-weight: 700;
    color: #1a1a2e;
}
.metric-label {
    font-size: 0.75rem;
    color: #9ca3af;
    margin-top: 0.1rem;
}

/* Download buttons */
.stDownloadButton > button {
    border-radius: 10px;
    height: 46px;
    font-size: 0.9rem;
    font-weight: 500;
}

/* Success / error overrides */
.stAlert {
    border-radius: 10px;
}

/* Divider */
.divider {
    border: none;
    border-top: 1px solid #e5e7eb;
    margin: 1rem 0;
}

/* Checkbox */
.stCheckbox label {
    font-size: 0.9rem;
    color: #374151;
}

/* Bigger file uploader drop zone */
[data-testid="stFileUploader"] {
    width: 100%;
}
[data-testid="stFileUploader"] section {
    padding: 3rem 2rem !important;
    border: 2.5px dashed #d1d5db !important;
    border-radius: 14px !important;
    background: #fafafa !important;
    transition: border-color 0.2s;
}
[data-testid="stFileUploader"] section:hover {
    border-color: #e63946 !important;
    background: #fff5f5 !important;
}
[data-testid="stFileUploader"] section > div {
    font-size: 1rem !important;
    color: #6b7280 !important;
}

/* Epic notice banner */
.epic-notice {
    background: #b91c1c;
    color: #ffffff;
    border-radius: 10px;
    padding: 0.85rem 1.2rem;
    font-size: 0.95rem;
    font-weight: 600;
    text-align: center;
    margin-bottom: 1.4rem;
    letter-spacing: 0.01em;
}
.epic-notice span {
    opacity: 0.85;
    font-weight: 400;
}
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
<div class="app-header">
    <div class="app-title">PO <span>→</span> Excel</div>
    <div class="app-subtitle">Convert purchase order PDFs to structured Excel files</div>
</div>
<div class="epic-notice">
    This tool is designed exclusively for <strong>Epic PO</strong> orders &nbsp;·&nbsp; <span>Other PDF formats will not parse correctly</span>
</div>
""", unsafe_allow_html=True)

# Upload card
st.markdown('<div class="section-label">Step 1 — Upload PDF</div>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("Drag & drop your PDF here, or click to browse", type="pdf", label_visibility="collapsed")

st.markdown('<div class="section-label" style="margin-top:1.2rem;">Step 2 — Options</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2)
with col1:
    compress = st.checkbox("Compress output as .zip", value=False)
with col2:
    pass

st.markdown('<div style="margin-top:1rem;"></div>', unsafe_allow_html=True)

# Convert button
convert = st.button("Convert to Excel")

# Debug switch
DEBUG = False

# Processing spinner
if convert and uploaded_file:
    st.markdown('<div class="processing-bar">Processing your file, please wait...</div>', unsafe_allow_html=True)

# Processing logic
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
                if line.split()[0].isdigit():
                    if current_item:
                        row = current_item + current_descriptions
                        data.append(row)
                    parts = line.split()
                    try:
                        line_no = parts[0]
                        date = parts[1]
                        item_id = parts[2]
                        qty = parts[-4]
                        unit = parts[-3]
                        price = parts[-2].replace("$", "").replace(",", "")
                        amount = parts[-1].replace("$", "").replace(",", "")
                        description = " ".join(parts[3:-4]).strip()
                        current_item = [line_no, date, item_id, qty, unit, price, amount]
                        current_descriptions = [description] if description else []
                    except Exception as e:
                        if DEBUG:
                            st.write("Skipped line:", line)
                        current_item = None
                        current_descriptions = []
                else:
                    if current_item:
                        current_descriptions.append(line)
            if current_item:
                row = current_item + current_descriptions
                data.append(row)

    # Output
    if data:
        max_desc = max(len(r) - 7 for r in data)
        columns = ["Line", "Date", "ItemID", "Qty", "Unit", "Price", "Amount"]
        columns += [f"Desc{i+1}" for i in range(max_desc)]
        for r in data:
            while len(r) < len(columns):
                r.append("")
        df = pd.DataFrame(data, columns=columns)

        st.success("Conversion successful!")

        # Metrics row
        total_amount = 0
        for v in df["Amount"]:
            try:
                total_amount += float(str(v).replace(",", ""))
            except Exception:
                pass
        st.markdown(f"""
        <div class="metrics-row">
            <div class="metric-box">
                <div class="metric-value">{len(df)}</div>
                <div class="metric-label">Line Items</div>
            </div>
            <div class="metric-box">
                <div class="metric-value">{len(df.columns)}</div>
                <div class="metric-label">Columns</div>
            </div>
            <div class="metric-box">
                <div class="metric-value">${total_amount:,.2f}</div>
                <div class="metric-label">Total Amount</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.dataframe(df, use_container_width=True)

        # Build Excel in memory
        excel_buffer = io.BytesIO()
        df.to_excel(excel_buffer, index=False)
        excel_bytes = excel_buffer.getvalue()

        st.markdown('<div class="section-label" style="margin-top:1rem;">Step 3 — Download</div>', unsafe_allow_html=True)

        if compress:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("output.xlsx", excel_bytes)
            st.download_button(
                "Download output.zip",
                data=zip_buffer.getvalue(),
                file_name="output.zip",
                mime="application/zip"
            )
        else:
            st.download_button(
                "Download output.xlsx",
                data=excel_bytes,
                file_name="output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("No data found in the PDF. Check that the file format matches the expected PO layout.")
