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
# MANUAL INPUTS
# ---------------------------------------------------

st.markdown("## ⚡ Manual Inputs")

col1, col2 = st.columns(2)

with col1:

    current_month_generation = st.number_input(
        "Solar Generation (kWh)",
        min_value=0.0,
        value=0.0
    )

    a_zone = st.number_input(
        "A Zone Units",
        min_value=0.0,
        value=0.0
    )

    b_zone = st.number_input(
        "B Zone Units",
        min_value=0.0,
        value=0.0
    )

    debit_bill_adjustment = st.number_input(
        "Debit Bill Adjustment (₹)",
        value=0.0
    )

with col2:

    c_zone = st.number_input(
        "C Zone Units",
        min_value=0.0,
        value=0.0
    )

    d_zone = st.number_input(
        "D Zone Units",
        min_value=0.0,
        value=0.0
    )

    grid_support_charges = st.number_input(
        "Grid Support Charges (₹)",
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
        # EXTRACT VALUES FROM BILL
        # ---------------------------------------------------

        contract_demand = extract_value(
            r'Total\s*Contract\s*Demand\s*\(KVA\)\s*([\d,\.]+)',
            text
        )

        highest_recorded_msedcl_demand = extract_value(
            r'Highest\s*Recorded\s*MSEDCL\s*Demand\s*([\d,\.]+)',
            text
        )

        if not highest_recorded_msedcl_demand:

            highest_recorded_msedcl_demand = extract_value(
                r'MSEDCL\s*Demand\s*([\d,\.]+)',
                text
            )

        billed_demand = extract_value(
            r'Billed\s*Demand\s*([\d,\.]+)',
            text
        )

        reference_units = extract_value(
            r'Ref\s*consumption\s*:?\s*([\d,\.]+)',
            text
        )

        transmission_charges = extract_value(
            r'Transmission\s*Charges\s*:?\s*₹?\s*([\d,\.]+)',
            text
        )

        # ---------------------------------------------------
        # DISPLAY DATA
        # ---------------------------------------------------

        st.markdown("## 📋 Extracted Bill Data")

        col3, col4 = st.columns(2)

        with col3:

            st.write("Contract Demand:", contract_demand)

            st.write(
                "Highest Recorded MSEDCL Demand:",
                highest_recorded_msedcl_demand
            )

            st.write(
                "Transmission Charges:",
                transmission_charges
            )

        with col4:

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
                # SELECT SHEETS
                # ---------------------------------------------------

                input_sheet = wb[wb.sheetnames[0]]

                output_sheet_name = "Bill After Solar_Apr 26"

                if output_sheet_name in wb.sheetnames:

                    output_sheet = wb[output_sheet_name]

                else:

                    output_sheet = wb[wb.sheetnames[1]]

                # ---------------------------------------------------
                # FIXED VALUES
                # ---------------------------------------------------

                input_sheet["C2"] = solar_capacity
                input_sheet["C3"] = plant_load

                # ---------------------------------------------------
                # TRANSMISSION CHARGES
                # ---------------------------------------------------

                input_sheet["C9"] = (
                    float(clean_number(transmission_charges))
                    if transmission_charges else 0
                )

                # ---------------------------------------------------
                # BILL VALUES
                # ---------------------------------------------------

                input_sheet["C14"] = (
                    float(clean_number(contract_demand))
                    if contract_demand else 0
                )

                input_sheet["C15"] = energy_rate
                input_sheet["C16"] = demand_charge_rate
                input_sheet["C17"] = wheeling_charge_rate
                input_sheet["C18"] = fac_rate
                input_sheet["C19"] = tax_rate
                input_sheet["C20"] = power_factor

                input_sheet["C21"] = (
                    float(clean_number(
                        highest_recorded_msedcl_demand
                    ))
                    if highest_recorded_msedcl_demand else 0
                )

                input_sheet["C22"] = electricity_duty

                # ---------------------------------------------------
                # MANUAL SOLAR GENERATION
                # ---------------------------------------------------

                input_sheet["H25"] = float(
                    current_month_generation
                )

                # ---------------------------------------------------
                # MANUAL TOD ZONES
                # ---------------------------------------------------

                input_sheet["K26"] = float(a_zone)
                input_sheet["L26"] = float(b_zone)
                input_sheet["M26"] = float(c_zone)
                input_sheet["N26"] = float(d_zone)

                # ---------------------------------------------------
                # OTHER BILL VALUES
                # ---------------------------------------------------

                input_sheet["C30"] = (
                    float(clean_number(billed_demand))
                    if billed_demand else 0
                )

                input_sheet["C40"] = (
                    float(clean_number(reference_units))
                    if reference_units else 0
                )

                # ---------------------------------------------------
                # OUTPUT SHEET VALUES
                # ---------------------------------------------------

                # C22 = Only Debit Bill Adjustment

                output_sheet["C22"] = float(
                    debit_bill_adjustment
                )

                # D22 = Debit + Grid Support

                output_sheet["D22"] = (
                    float(debit_bill_adjustment)
                    +
                    float(grid_support_charges)
                )

                # ---------------------------------------------------
                # FORCE FORMULA RECALCULATION
                # ---------------------------------------------------

                wb.calculation.fullCalcOnLoad = True
                wb.calculation.forceFullCalc = True

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
