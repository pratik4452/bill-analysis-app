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

energy_rate = 8.44
demand_charge_rate = 650
wheeling_charge_rate = 0.81
fac_rate = 0.50
tax_rate = 0.29

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
# CLEAN NUMBER
# ---------------------------------------------------

def clean_number(value):

    if value:

        return (
            str(value)
            .replace(",", "")
            .replace("₹", "")
            .replace("%", "")
            .strip()
        )

    return "0"

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
        # REMOVE LINE BREAKS
        # ---------------------------------------------------

        single_text = text.replace("\n", " ")

        # ---------------------------------------------------
        # MONTH
        # ---------------------------------------------------

        month_match = re.search(
            r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[\-\s]?\d{2}',
            single_text,
            re.IGNORECASE
        )

        bill_month = ""

        if month_match:
            bill_month = month_match.group(0)

        # ---------------------------------------------------
        # CONTRACT DEMAND
        # ---------------------------------------------------

        contract_demand = extract_value(
            r'Total\s*Contract\s*Demand\s*\(KVA\)\s*([\d,\.]+)',
            single_text
        )

        # ---------------------------------------------------
        # MAXIMUM DEMAND
        # ---------------------------------------------------

        highest_recorded_msedcl_demand = extract_value(
            r'Highest\s*Recorded\s*MSEDCL\s*Demand\s*([\d,\.]+)',
            single_text
        )

        if not highest_recorded_msedcl_demand:

            highest_recorded_msedcl_demand = extract_value(
                r'MSEDCL\s*Demand\s*([\d,\.]+)',
                single_text
            )

        # ---------------------------------------------------
        # TRANSMISSION CHARGES
        # ---------------------------------------------------

        transmission_charges = extract_value(
            r'Transmission\s*Charges\s*:?\s*₹?\s*([\d,\.]+)',
            single_text
        )

        # ---------------------------------------------------
        # BILLED DEMAND
        # ---------------------------------------------------

        billed_demand = extract_value(
            r'Billed\s*Demand\s*([\d,\.]+)',
            single_text
        )

        # ---------------------------------------------------
        # REFERENCE UNITS
        # ---------------------------------------------------

        reference_units = extract_value(
            r'Ref\s*consumption\s*:?\s*([\d,\.]+)',
            single_text
        )

        # ---------------------------------------------------
        # ELECTRICITY DUTY
        # ---------------------------------------------------

        electricity_duty = extract_value(
            r'Electricity\s*Duty\s*[:\-]?\s*([\d\.]+%)',
            single_text
        )

        if not electricity_duty:
            electricity_duty = "7.50%"

        # ---------------------------------------------------
        # POWER FACTOR EXTRACTION
        # ---------------------------------------------------

        power_factor = ""

        # METHOD 1
        # 0.996 26 P.F.

        pf_match = re.search(

            r'(\d+\.\d+)\s+26\s+P\.F',

            single_text,

            re.IGNORECASE

        )

        if pf_match:

            power_factor = pf_match.group(1)

        # METHOD 2
        # P.F. 0.996

        if not power_factor:

            pf_match = re.search(

                r'P\.F\.?\s*[:\-]?\s*(\d+\.\d+)',

                single_text,

                re.IGNORECASE

            )

            if pf_match:

                power_factor = pf_match.group(1)

        # METHOD 3
        # Power Factor

        if not power_factor:

            pf_match = re.search(

                r'Power\s*Factor\s*[:\-]?\s*(\d+\.\d+)',

                single_text,

                re.IGNORECASE

            )

            if pf_match:

                power_factor = pf_match.group(1)

        # DEFAULT

        if not power_factor:

            power_factor = "1"

        # ---------------------------------------------------
        # DISPLAY DATA
        # ---------------------------------------------------

        st.markdown("## 📋 Extracted Bill Data")

        col3, col4 = st.columns(2)

        with col3:

            st.write("Month:", bill_month)
            st.write("Contract Demand:", contract_demand)
            st.write("Maximum Demand:", highest_recorded_msedcl_demand)
            st.write("Transmission Charges:", transmission_charges)

        with col4:

            st.write("P.F.:", power_factor)
            st.write("Electricity Duty:", electricity_duty)
            st.write("Billed Demand:", billed_demand)
            st.write("Reference Units:", reference_units)

        st.markdown("---")

        # ---------------------------------------------------
        # GENERATE EXCEL
        # ---------------------------------------------------

        if st.button("Generate Excel Report"):

            try:

                # ---------------------------------------------------
                # TEMPLATE PATH
                # ---------------------------------------------------

                template_path = os.path.join(
                    "templates",
                    "bill_template.xlsx"
                )

                wb = load_workbook(template_path)

                # ---------------------------------------------------
                # SHEETS
                # ---------------------------------------------------

                input_sheet = wb[wb.sheetnames[0]]

                if len(wb.sheetnames) > 1:

                    output_sheet = wb[wb.sheetnames[1]]

                else:

                    output_sheet = input_sheet

                # ---------------------------------------------------
                # FIXED VALUES
                # ---------------------------------------------------

                input_sheet["C2"] = solar_capacity
                input_sheet["C3"] = plant_load

                # ---------------------------------------------------
                # MONTH
                # ---------------------------------------------------

                input_sheet["C13"] = bill_month

                # ---------------------------------------------------
                # TRANSMISSION CHARGES
                # ---------------------------------------------------

                input_sheet["C9"] = float(
                    clean_number(transmission_charges)
                )

                # ---------------------------------------------------
                # BILL VALUES
                # ---------------------------------------------------

                input_sheet["C14"] = float(
                    clean_number(contract_demand)
                )

                input_sheet["C15"] = energy_rate

                input_sheet["C16"] = demand_charge_rate

                input_sheet["C17"] = wheeling_charge_rate

                input_sheet["C18"] = fac_rate

                input_sheet["C19"] = tax_rate

                # ---------------------------------------------------
                # POWER FACTOR FROM PDF
                # ---------------------------------------------------

                input_sheet["C20"] = float(
                    clean_number(power_factor)
                )

                # ---------------------------------------------------
                # MAXIMUM DEMAND
                # ---------------------------------------------------

                input_sheet["C21"] = float(
                    clean_number(
                        highest_recorded_msedcl_demand
                    )
                )

                # ---------------------------------------------------
                # ELECTRICITY DUTY
                # ---------------------------------------------------

                input_sheet["C22"] = electricity_duty

                # ---------------------------------------------------
                # SOLAR GENERATION
                # ---------------------------------------------------

                input_sheet["H25"] = float(
                    current_month_generation
                )

                # ---------------------------------------------------
                # ZONE VALUES
                # ---------------------------------------------------

                input_sheet["K26"] = float(a_zone)
                input_sheet["L26"] = float(b_zone)
                input_sheet["M26"] = float(c_zone)
                input_sheet["N26"] = float(d_zone)

                # ---------------------------------------------------
                # OTHER VALUES
                # ---------------------------------------------------

                input_sheet["C30"] = float(
                    clean_number(billed_demand)
                )

                input_sheet["C40"] = float(
                    clean_number(reference_units)
                )

                # ---------------------------------------------------
                # OUTPUT SHEET VALUES
                # ---------------------------------------------------

                output_sheet["C22"] = float(
                    debit_bill_adjustment
                )

                output_sheet["D22"] = (
                    float(debit_bill_adjustment)
                    +
                    float(grid_support_charges)
                )

                # ---------------------------------------------------
                # RECALCULATE
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
                    "✅ Excel Report Generated Successfully"
                )

                # ---------------------------------------------------
                # DOWNLOAD
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
