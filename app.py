import streamlit as st
import pdfplumber
import re
import os
from openpyxl import load_workbook
from io import BytesIO

# -------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------

st.set_page_config(
    page_title="DISCOM Bill Analysis",
    layout="wide"
)

st.title("⚡ DISCOM Bill Analysis App")
st.subheader("Before Solar vs After Solar Analysis")

# -------------------------------------------------
# FILE UPLOAD
# -------------------------------------------------

uploaded_file = st.file_uploader(
    "Upload Electricity Bill PDF",
    type=["pdf"]
)

# -------------------------------------------------
# HELPER FUNCTION
# -------------------------------------------------

def extract_value(pattern, text):

    match = re.search(pattern, text)

    if match:
        return match.group(1)

    return ""

# -------------------------------------------------
# MAIN LOGIC
# -------------------------------------------------

if uploaded_file:

    text = ""

    try:

        # -----------------------------------------
        # READ PDF
        # -----------------------------------------

        with pdfplumber.open(uploaded_file) as pdf:

            for page in pdf.pages:

                extracted = page.extract_text()

                if extracted:
                    text += extracted

        st.success("✅ PDF Processed Successfully")

        # -----------------------------------------
        # EXTRACT DATA
        # -----------------------------------------

        consumer_number = extract_value(
            r'Consumer Number\s+(\d+)',
            text
        )

        bill_month = extract_value(
            r'Bill Month\s+([A-Z0-9\-]+)',
            text
        )

        payable_amount = extract_value(
            r'Total Bill Amount \(Rounded\) Rs\.\s+([\d,]+\.\d+)',
            text
        )

        billed_demand = extract_value(
            r'Billed Demand\s+([\d\.]+)',
            text
        )

        total_drawal = extract_value(
            r'01\-APR\-2026 TO 30\-APR\-2026\s+([\d,]+)',
            text
        )

        # -----------------------------------------
        # DISPLAY EXTRACTED DATA
        # -----------------------------------------

        st.markdown("## 📋 Extracted Bill Details")

        col1, col2 = st.columns(2)

        with col1:
            st.write("### Consumer Details")
            st.write("Consumer Number:", consumer_number)
            st.write("Bill Month:", bill_month)

        with col2:
            st.write("### Billing Details")
            st.write("Payable Amount:", payable_amount)
            st.write("Billed Demand:", billed_demand)
            st.write("Total Drawal Units:", total_drawal)

        st.markdown("---")

        # -----------------------------------------
        # GENERATE EXCEL
        # -----------------------------------------

        if st.button("Generate Excel Report"):

            try:

                # TEMPLATE PATH
                template_path = os.path.join(
                    "templates",
                    "bill_template.xlsx"
                )

                # LOAD EXCEL
                wb = load_workbook(template_path)

                # SHOW SHEETS
                st.write("Available Sheets:", wb.sheetnames)

                # SELECT FIRST SHEET
                ws = wb[wb.sheetnames[0]]

                # -------------------------------------------------
                # IMPORTANT:
                # USE ONLY NON-MERGED SAFE CELLS
                # -------------------------------------------------

                ws["B2"] = "Consumer Number"
                ws["C2"] = consumer_number

                ws["B3"] = "Bill Month"
                ws["C3"] = bill_month

                ws["B4"] = "Payable Amount"
                ws["C4"] = payable_amount

                ws["B5"] = "Billed Demand"
                ws["C5"] = billed_demand

                ws["B6"] = "Total Drawal"
                ws["C6"] = total_drawal

                # -----------------------------------------
                # SAVE OUTPUT
                # -----------------------------------------

                output = BytesIO()

                wb.save(output)

                output.seek(0)

                st.success("✅ Excel Report Generated Successfully")

                # -----------------------------------------
                # DOWNLOAD BUTTON
                # -----------------------------------------

                st.download_button(
                    label="⬇ Download Excel Report",
                    data=output,
                    file_name="Before_After_Solar_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as excel_error:

                st.error(f"Excel Generation Error: {excel_error}")

    except Exception as e:

        st.error(f"PDF Processing Error: {e}")
