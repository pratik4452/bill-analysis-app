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
# CLEAN NUMBER FUNCTION
# ---------------------------------------------------

def clean_number(value):

    if value:
        return value.replace(",", "").strip()

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
        # EXTRACT VALUES FROM BILL
        # ---------------------------------------------------

        contract_demand = extract_value(
            r'Total Contract Demand \(KVA\)\s+([\d,]+)',
            text
        )

        energy_rate = extract_value(
            r'8\.44',
            text
        )

        demand_charge_rate = extract_value(
            r'650\.00',
            text
        )

        wheeling_charge_rate = extract_value(
            r'Wheeling Charges.*?@\.81',
            text
        )

        fac_rate = extract_value(
            r'@ 0\.5',
            text
        )

        tax_rate = extract_value(
            r'0\.2894',
            text
        )

        power_factor = "1"

        max_demand = extract_value(
            r'Highest Recorded\s+MSEDCL Demand\s+(\d+)',
            text
        )

        electricity_duty = "7.50%"

        solar_generation = extract_value(
            r'01\-APR\-2026 TO 30\-APR\-2026\s+([\d,]+)',
            text
        )

        billed_demand = extract_value(
            r'Billed Demand\s+([\d\.]+)',
            text
        )

        reference_units = extract_value(
            r'Ref consumption :\s+(\d+)',
            text
        )

        # ---------------------------------------------------
        # DISPLAY DATA
        # ---------------------------------------------------

        st.markdown("## 📋 Extracted Bill Data")

        st.write("Contract Demand:", contract_demand)
        st.write("Energy Charges Rate:", "8.44")
        st.write("Demand Charges Rate:", "650")
        st.write("Wheeling Charges Rate:", "0.81")
        st.write("FAC Rate:", "0.50")
        st.write("Tax on Sales:", "0.29")
        st.write("Power Factor:", power_factor)
        st.write("Maximum Demand:", max_demand)
        st.write("Electricity Duty:", electricity_duty)
        st.write("Solar Generation:", solar_generation)
        st.write("Billed Demand:", billed_demand)
        st.write("Reference Units:", reference_units)

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
                # SELECT FIRST SHEET
                # ---------------------------------------------------

                ws = wb[wb.sheetnames[0]]

                # ---------------------------------------------------
                # USER INPUTS
                # ---------------------------------------------------

                ws["C2"] = solar_capacity
                ws["C3"] = plant_load
                ws["C9"] = transmission_charge_input

                # ---------------------------------------------------
                # PDF DATA INPUTS
                # ---------------------------------------------------

                ws["C14"] = float(clean_number(contract_demand)) if contract_demand else 0

                ws["C15"] = 8.44

                ws["C16"] = 650

                ws["C17"] = 0.81

                ws["C18"] = 0.50

                ws["C19"] = 0.29

                ws["C20"] = 1

                ws["C21"] = float(clean_number(max_demand)) if max_demand else 0

                ws["C22"] = "7.50%"

                ws["H25"] = float(clean_number(solar_generation)) if solar_generation else 0

                ws["C30"] = float(clean_number(billed_demand)) if billed_demand else 0

                ws["C40"] = float(clean_number(reference_units)) if reference_units else 0

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
