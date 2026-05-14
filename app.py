import streamlit as st
import pdfplumber
import re
from openpyxl import load_workbook
from io import BytesIO

st.set_page_config(page_title="DISCOM Bill Analysis", layout="wide")

st.title("⚡ DISCOM Bill Analysis App")
st.subheader("Before Solar vs After Solar Analysis")

uploaded_file = st.file_uploader(
    "Upload Electricity Bill PDF",
    type=["pdf"]
)

def extract_value(pattern, text):
    match = re.search(pattern, text)

    if match:
        return match.group(1)

    return ""

if uploaded_file:

    text = ""

    with pdfplumber.open(uploaded_file) as pdf:

        for page in pdf.pages:

            extracted = page.extract_text()

            if extracted:
                text += extracted

    st.success("PDF Processed Successfully")

    consumer_number = extract_value(
        r'Consumer Number\\s+(\\d+)',
        text
    )

    bill_month = extract_value(
        r'Bill Month\\s+([A-Z\\-0-9]+)',
        text
    )

    payable_amount = extract_value(
        r'Total Bill Amount \\(Rounded\\) Rs\\.\\s+([\\d,]+\\.\\d+)',
        text
    )

    billed_demand = extract_value(
        r'Billed Demand\\s+([\\d\\.]+)',
        text
    )

    total_drawal = extract_value(
        r'01\\-APR\\-2026 TO 30\\-APR\\-2026\\s+([\\d,]+)',
        text
    )

    st.markdown("## Extracted Bill Details")

    col1, col2 = st.columns(2)

    with col1:
        st.write("Consumer Number:", consumer_number)
        st.write("Bill Month:", bill_month)

    with col2:
        st.write("Payable Amount:", payable_amount)
        st.write("Billed Demand:", billed_demand)

    st.markdown("---")

    if st.button("Generate Excel Report"):

        wb = load_workbook("templates/bill_template.xlsx")

        ws = wb.active

        # SAMPLE CELL MAPPING
        # CHANGE THESE CELLS AS PER YOUR TEMPLATE

        ws["C5"] = consumer_number
        ws["C6"] = bill_month
        ws["C7"] = payable_amount
        ws["C8"] = billed_demand
        ws["C9"] = total_drawal

        output = BytesIO()

        wb.save(output)

        output.seek(0)

        st.success("Excel Report Generated")

        st.download_button(
            label="Download Excel Report",
            data=output,
            file_name="Before_After_Solar_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
