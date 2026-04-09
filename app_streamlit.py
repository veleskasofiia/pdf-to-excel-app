import streamlit as st
import pdfplumber
import pandas as pd

# Page config
st.set_page_config(page_title="PDF to Excel", layout="centered")

# 🔹 Custom styling (modern look)
st.markdown("""
<style>
body {
    background-color: #f5f5f5;
}
.main-title {
    text-align: center;
    font-size: 40px;
    font-weight: bold;
    color: #ff4b4b;
}
.sub-text {
    text-align: center;
    color: #666;
    margin-bottom: 30px;
}
.upload-box {
    border: 2px dashed #ccc;
    padding: 30px;
    text-align: center;
    border-radius: 10px;
    background-color: white;
}
.stButton button {
    background-color: #ff4b4b;
    color: white;
    border-radius: 8px;
    height: 50px;
    width: 100%;
    font-size: 16px;
}
</style>
""", unsafe_allow_html=True)

# 🔹 Header
st.markdown('<div class="main-title">PDF to Excel</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-text">Upload your PDF and convert it instantly</div>', unsafe_allow_html=True)

# 🔹 Upload section
uploaded_file = st.file_uploader("📄 Drag & drop your PDF here", type="pdf")

# 🔹 Convert button
convert = st.button("Convert to Excel")

# 🔹 Debug switch
DEBUG = False

# 🔹 Processing
if uploaded_file and convert:
    st.info("Processing your file... ⏳")

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

                        current_item = [
                            line_no, date, item_id,
                            qty, unit, price, amount
                        ]

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

    # 🔹 Output
    if data:
        max_desc = max(len(r) - 7 for r in data)

        columns = ["Line", "Date", "ItemID", "Qty", "Unit", "Price", "Amount"]
        columns += [f"Desc{i+1}" for i in range(max_desc)]

        for r in data:
            while len(r) < len(columns):
                r.append("")

        df = pd.DataFrame(data, columns=columns)

        st.success("✅ Conversion successful!")

        st.dataframe(df, use_container_width=True)

        df.to_excel("output.xlsx", index=False)

        with open("output.xlsx", "rb") as f:
            st.download_button("⬇️ Download Excel", f, "output.xlsx")

    else:
        st.error("❌ No data found in PDF")