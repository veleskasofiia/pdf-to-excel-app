import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

# ✅ Page config
st.set_page_config(
    page_title="PDF to Excel",
    page_icon="📊",
    layout="wide"
)

# 🔹 Styling
st.markdown("""
<style>
.main-title { text-align: center; font-size: 42px; font-weight: bold; color: #ff4b4b; }
.sub-text { text-align: center; color: #666; margin-bottom: 30px; }
.stButton button { background-color: #ff4b4b; color: white; border-radius: 8px; height: 50px; width: 100%; font-size: 16px; }
</style>
""", unsafe_allow_html=True)

# 🔹 Header
st.markdown('<div class="main-title">PDF to Excel</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-text">Upload your PDF and convert it instantly</div>', unsafe_allow_html=True)

# 🔹 Upload
uploaded_file = st.file_uploader("📄 Drag & drop your PDF here", type="pdf")
convert = st.button("Convert to Excel")
DEBUG = False

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

                parts = line.split()

                # ✅ ONLY your working logic
                if len(parts) > 5 and parts[0].isdigit():

                    if current_item:
                        row = current_item + current_descriptions
                        data.append(row)

                    try:
                        line_no = parts[0]
                        date = parts[1]
                        item_id = parts[2]
                        qty = parts[-4]
                        unit = parts[-3]
                        price = parts[-2].replace("$","").replace(",","")
                        amount = parts[-1].replace("$","").replace(",","")

                        description = " ".join(parts[3:-4]).strip()

                        current_item = [line_no, date, item_id, qty, unit, price, amount]
                        current_descriptions = [description] if description else []

                    except:
                        if DEBUG:
                            st.write("Skipped line:", line)
                        current_item = None
                        current_descriptions = []

                else:
                    # ✅ multi-line description support
                    if current_item:
                        current_descriptions.append(line)

            # Save last row
            if current_item:
                row = current_item + current_descriptions
                data.append(row)

    # ✅ Create DataFrame
    if data:
        max_desc = max(len(r) - 7 for r in data)

        columns = ["Line","Date","ItemID","Qty","Unit","Price","Amount"]
        columns += [f"Desc{i+1}" for i in range(max_desc)]

        for r in data:
            while len(r) < len(columns):
                r.append("")

        df = pd.DataFrame(data, columns=columns)

        st.success("✅ Conversion successful!")

        # 🔹 Column rename
        st.subheader("🛠 Customize Columns")
        new_columns = {}
        for col in df.columns:
            new_name = st.text_input(f"Rename '{col}'", value=col)
            new_columns[col] = new_name

        df.rename(columns=new_columns, inplace=True)

        # 🔥 BIG table
        st.subheader("📊 Preview")
        st.data_editor(df, use_container_width=True, height=600)

        # 📥 Download
        output = BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            "⬇️ Download Excel",
            output,
            "output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.error("❌ No valid rows found in PDF")