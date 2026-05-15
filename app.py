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

        if match.groups():
            return match.group(1)

        return match.group(0)

    return ""

# ---------------------------------------------------
# CLEAN NUMBER FUNCTION
# ---------------------------------------------------

def clean_number(value):

    if value:

        return (
            str(value)
            .replace(",", "")
            .replace("%", "")
            .replace("₹", "")
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
        # EXTRACT VALUES FROM BILL
        # ---------------------------------------------------

        contract_demand = extract_value(
            r'Contract\s*Demand.*?([\d,]+)',
            text
        )

        highest_recorded_msedcl_demand = extract_value(
            r'Highest\s*Recorded.*?Demand.*?([\d,]+)',
            text
        )

        if not highest_recorded_msedcl_demand:

            highest_recorded_msedcl_demand = extract_value(
                r'Maximum\s*Demand.*?([\d,]+)',
                text
            )

        billed_demand = extract_value(
            r'Billed\s*Demand.*?([\d,]+)',
            text
        )

        reference_units = extract_value(
            r'Ref.*?Consumption.*?([\d,]+)',
            text
        )

        transmission_charges = extract_value(
            r'Transmission\s*Charges.*?([\d,]+\.\d+|[\d,]+)',
            text
        )

        bill_month = extract_value(
            r'Apr-\d+|May-\d+|Jun-\d+|Jul-\d+|Aug-\d+|Sep-\d+|Oct-\d+|Nov-\d+|Dec-\d+|Jan-\d+|Feb-\d+|Mar-\d+',
            text
        )

        # ---------------------------------------------------
        # AUTO FETCH VALUES FROM PDF
        # ---------------------------------------------------

        energy_rate = extract_value(
            r'Energy\s*Charges.*?([\d]+\.[\d]+)',
            text
        )

        demand_charge_rate = extract_value(
            r'Demand\s*Charges.*?([\d,]+\.[\d]+|[\d,]+)',
            text
        )

        wheeling_charge_rate = extract_value(
            r'Wheeling\s*Charges.*?([\d]+\.[\d]+)',
            text
        )

        fac_rate = extract_value(
            r'FAC.*?([\d]+\.[\d]+)',
            text
        )

        tax_rate = extract_value(
            r'Tax\s*on\s*Sales.*?([\d]+\.[\d]+)',
            text
        )

        if not tax_rate:

            tax_rate = extract_value(
                r'Tax.*?([\d]+\.[\d]+)',
                text
            )

        power_factor = extract_value(
            r'Power\s*Factor.*?([\d]+\.[\d]+|[\d]+)',
            text
        )

        if not power_factor:

            power_factor = "1"

        electricity_duty = extract_value(
            r'Electricity\s*Duty.*?([\d]+\.[\d]+%)',
            text
        )

        if not electricity_duty:

            electricity_duty = "7.50%"

        # ---------------------------------------------------
        # DISPLAY DATA
        # ---------------------------------------------------

        st.markdown("## 📋 Extracted Bill Data")

        col3, col4 = st.columns(2)

        with col3:

            st.write("Month:", bill_month)

            st.write(
                "Contract Demand:",
                contract_demand
            )

            st.write(
                "Highest Recorded MSEDCL Demand:",
                highest_recorded_msedcl_demand
            )

            st.write(
                "Energy Charges:",
                energy_rate
            )

            st.write(
                "Demand Charges:",
                demand_charge_rate
            )

            st.write(
                "Wheeling Charges:",
                wheeling_charge_rate
            )

        with col4:

            st.write(
                "FAC:",
                fac_rate
            )

            st.write(
                "Tax on Sales:",
                tax_rate
            )

            st.write(
                "Power Factor:",
                power_factor
            )

            st.write(
                "Electricity Duty:",
                electricity_duty
            )

            st.write(
                "Transmission Charges:",
                transmission_charges
            )

            st.write(
                "Reference Units:",
                reference_units
            )

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
                # SELECT SHEETS
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

                input_sheet["C9"] = (
                    float(clean_number(
                        transmission_charges
                    ))
                    if transmission_charges else 0
                )

                # ---------------------------------------------------
                # BILL VALUES
                # ---------------------------------------------------

                input_sheet["C14"] = (
                    float(clean_number(
                        contract_demand
                    ))
                    if contract_demand else 0
                )

                input_sheet["C15"] = float(
                    clean_number(energy_rate)
                )

                input_sheet["C16"] = float(
                    clean_number(demand_charge_rate)
                )

                input_sheet["C17"] = float(
                    clean_number(wheeling_charge_rate)
                )

                input_sheet["C18"] = float(
                    clean_number(fac_rate)
                )

                input_sheet["C19"] = float(
                    clean_number(tax_rate)
                )

                input_sheet["C20"] = float(
                    clean_number(power_factor)
                )

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
                    float(clean_number(
                        billed_demand
                    ))
                    if billed_demand else 0
                )

                input_sheet["C40"] = (
                    float(clean_number(
                        reference_units
                    ))
                    if reference_units else 0
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
