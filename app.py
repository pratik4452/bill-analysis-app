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
# FIXED VALUES
# ---------------------------------------------------

solar_capacity = 1000

plant_load = 1800

# ---------------------------------------------------
# MANUAL CURRENT MONTH GENERATION
# ---------------------------------------------------

current_month_generation = st.number_input(
    "Enter Current Month Generation (kWh)",
    min_value=0.0,
    value=237415.0
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

    match = re.search(
        pattern,
        text,
        re.IGNORECASE | re.DOTALL
    )

    if match:
        return match.group(1)

    return ""

# ---------------------------------------------------
# CLEAN NUMBER FUNCTION
# ---------------------------------------------------

def clean_number(value):

    if value:

        return (
            value
            .replace(",", "")
            .replace("%", "")
            .replace("₹", "")
            .strip()
        )

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
                    text += extracted + "\n"

        st.success("✅ PDF Processed Successfully")

        # ---------------------------------------------------
        # DEBUG PDF TEXT
        # Uncomment if needed
        # ---------------------------------------------------

        # st.text(text)

        # ---------------------------------------------------
        # EXTRACT VALUES FROM BILL
        # ---------------------------------------------------

        contract_demand = extract_value(
            r'Total Contract Demand \(KVA\)\s+([\d,]+)',
            text
        )

        max_demand = extract_value(
            r'Highest Recorded\s+MSEDCL Demand\s+([\d,]+)',
            text
        )

        billed_demand = extract_value(
            r'Billed Demand\s+([\d\.]+)',
            text
        )

        reference_units = extract_value(
            r'Ref consumption\s*:?\s*([\d,]+)',
            text
        )

        # ---------------------------------------------------
        # TRANSMISSION CHARGES FROM PDF
        # ---------------------------------------------------

        transmission_charges = extract_value(
            r'Transmission Charges\s*:?[\s\n₹]*([\d,\.]+)',
            text
        )

        # ---------------------------------------------------
        # DISPLAY DATA
        # ---------------------------------------------------

        st.markdown("## 📋 Extracted Bill Data")

        col1, col2 = st.columns(2)

        with col1:

            st.write("Contract Demand:", contract_demand)
            st.write("Transmission Charges:", transmission_charges)
            st.write("Maximum Demand:", max_demand)

        with col2:

            st.write("Current Month Generation:", current_month_generation)
            st.write("Billed Demand:", billed_demand)
            st.write("Reference Units:", reference_units)

        st.markdown("---")

        # ---------------------------------------------------
        # STATIC VALUES
        # ---------------------------------------------------

        energy_rate = 8.44
        demand_charge_rate = 650
        wheeling_charge_rate = 0.81
        fac_rate = 0.50
        tax_rate = 0.29
        power_factor = 1
        electricity_duty = "7.50%"

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

                ws = wb[wb.sheetnames[0]]

                # ---------------------------------------------------
                # FIXED VALUES
                # ---------------------------------------------------

                ws["C2"] = solar_capacity

                ws["C3"] = plant_load

                # ---------------------------------------------------
                # TRANSMISSION CHARGES
                # ---------------------------------------------------

                ws["C9"] = (
                    float(clean_number(transmission_charges))
                    if transmission_charges else 0
                )

                # ---------------------------------------------------
                # BILL VALUES
                # ---------------------------------------------------

                ws["C14"] = (
                    float(clean_number(contract_demand))
                    if contract_demand else 0
                )

                ws["C15"] = energy_rate

                ws["C16"] = demand_charge_rate

                ws["C17"] = wheeling_charge_rate

                ws["C18"] = fac_rate

                ws["C19"] = tax_rate

                ws["C20"] = power_factor

                ws["C21"] = (
                    float(clean_number(max_demand))
                    if max_demand else 0
                )

                ws["C22"] = electricity_duty

                # ---------------------------------------------------
                # CURRENT MONTH GENERATION
                # ---------------------------------------------------

                ws["H25"] = current_month_generation

                # ---------------------------------------------------
                # OTHER BILL VALUES
                # ---------------------------------------------------

                ws["C30"] = (
                    float(clean_number(billed_demand))
                    if billed_demand else 0
                )

                ws["C40"] = (
                    float(clean_number(reference_units))
                    if reference_units else 0
                )

                # ---------------------------------------------------
                # SAVE OUTPUT
                # ---------------------------------------------------

                output = BytesIO()

                wb.save(output)

                output.seek(0)

                st.success(
                    "✅ Before vs After Solar Report Generated Successfully"
                )

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

                st.error(
                    f"Excel Generation Error: {excel_error}"
                )

    except Exception as e:

        st.error(
            f"PDF Processing Error: {e}"
        )
