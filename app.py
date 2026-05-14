import streamlit as st
import pdfplumber
import re
import os
from openpyxl import load_workbook
from io import BytesIO

# ---------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------

st.set_page_config(
    page_title="DISCOM Bill Analysis",
    layout="wide"
)

st.title("⚡ DISCOM Bill Analysis App")
st.subheader("Before Solar vs After Solar Analysis")

# ---------------------------------------------------
# USER INPUTS
# ---------------------------------------------------

solar_capacity = st.number_input(
    "Enter Solar Capacity (kW)",
    min_value=0.0,
    value=1000.0
)

plant_load = st.number_input(
    "Enter Plant Load / Contract Demand",
    min_value=0.0,
    value=1800.0
)

transmission_charge_input = st.number_input(
    "Enter Transmission Charges",
    min_value=0.0,
    value=0.0
)

# ---------------------------------------------------
# FILE UPLOAD
# ---------------------------------------------------

uploaded_file = st.file_uploader(
    "Upload Electricity Bill PDF",
    type=["pdf"]
)

# ---------------------------------------------------
# HELPER FUNCTION
# ---------------------------------------------------

def extract_value(pattern, text):

    match = re.search(pattern, text)

    if match:
        return match.group(1)

    return ""

# ---------------------------------------------------
# MAIN PROCESS
# ---------------------------------------------------

if uploaded_file:

    text = ""

    try:

        # ---------------------------------------------------
        # READ PDF
        # ---------------------------------------------------

        with pdfplumber.open(uploaded_file) as pdf:

            for page in pdf.pages:

                extracted = page.extract_text()

                if extracted:
                    text += extracted

        st.success("✅ PDF Processed Successfully")

        # ---------------------------------------------------
        # EXTRACT BILL DATA
        # ---------------------------------------------------

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

        energy_charges = extract_value(
            r'Energy Charges\s+([\d,]+\.\d+)',
            text
        )

        fac_charges = extract_value(
            r'FAC Charges\s+([\d,]+\.\d+)',
            text
        )

        wheeling_charges = extract_value(
            r'Wheeling Charges\s+([\d,]+\.\d+)',
            text
        )

        # ---------------------------------------------------
        # DISPLAY EXTRACTED DATA
        # ---------------------------------------------------

        st.markdown("## 📋 Extracted Bill Details")

        col1, col2 = st.columns(2)

        with col1:
            st.write("### Consumer Details")
            st.write("Consumer Number:", consumer_number)
            st.write("Bill Month:", bill_month)
            st.write("Payable Amount:", payable_amount)

        with col2:
            st.write("### Billing Details")
            st.write("Billed Demand:", billed_demand)
            st.write("Total Drawal Units:", total_drawal)
            st.write("Energy Charges:", energy_charges)

        st.markdown("---")

        # ---------------------------------------------------
        # GENERATE EXCEL REPORT
        # ---------------------------------------------------

        if st.button("Generate Excel Report"):

            try:

                # ---------------------------------------------------
                # LOAD TEMPLATE
                # ---------------------------------------------------

                template_path = os.path.join(
                    "templates",
                    "bill_template.xlsx"
                )

                wb = load_workbook(template_path)

                # ---------------------------------------------------
                # AUTO SELECT FIRST SHEET
                # ---------------------------------------------------

                sheet_names = wb.sheetnames

                st.write("Detected Sheets:", sheet_names)

                ws = wb[sheet_names[0]]

                # ---------------------------------------------------
                # USER INPUTS
                # ---------------------------------------------------

                ws["C2"] = solar_capacity
                ws["C3"] = plant_load
                ws["C9"] = transmission_charge_input

                # ---------------------------------------------------
                # BILL DATA INPUTS
                # ---------------------------------------------------

                ws["C13"] = consumer_number
                ws["C14"] = bill_month
                ws["C15"] = payable_amount
                ws["C16"] = billed_demand
                ws["C17"] = total_drawal
                ws["C18"] = energy_charges
                ws["C19"] = fac_charges
                ws["C20"] = wheeling_charges

                # Optional extra inputs
                ws["C21"] = payable_amount
                ws["C22"] = total_drawal

                # ---------------------------------------------------
                # SAVE OUTPUT
                # ---------------------------------------------------

                output = BytesIO()

                wb.save(output)

                output.seek(0)

                st.success("✅ Before vs After Solar Report Generated Successfully")

                # ---------------------------------------------------
                # DOWNLOAD BUTTON
                # ---------------------------------------------------

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
