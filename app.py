import streamlit as st
import pdfplumber
import re
import os
import pandas as pd

from openpyxl import load_workbook
from io import BytesIO

# ---------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------

st.set_page_config(
    page_title="DISCOM Bill Analysis Dashboard",
    layout="wide"
)

st.title("⚡ DISCOM Bill Analysis Dashboard")
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
power_factor = 1
electricity_duty = "7.50%"

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
# HELPER FUNCTIONS
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
        # EXTRACT VALUES
        # ---------------------------------------------------

        contract_demand = extract_value(
            r'Total\s*Contract\s*Demand\s*\(KVA\)\s*([\d,\.]+)',
            text
        )

        highest_recorded_msedcl_demand = extract_value(
            r'Highest\s*Recorded\s*MSEDCL\s*Demand\s*([\d,\.]+)',
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
        # DISPLAY EXTRACTED DATA
        # ---------------------------------------------------

        st.markdown("## 📋 Extracted Bill Data")

        c1, c2 = st.columns(2)

        with c1:

            st.info(
                f"Contract Demand: {contract_demand}"
            )

            st.info(
                f"Highest Recorded MSEDCL Demand: "
                f"{highest_recorded_msedcl_demand}"
            )

        with c2:

            st.info(
                f"Transmission Charges: ₹ "
                f"{transmission_charges}"
            )

            st.info(
                f"Reference Units: "
                f"{reference_units}"
            )

        # ---------------------------------------------------
        # GENERATE REPORT
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
                # SHEETS
                # ---------------------------------------------------

                sheet_names = wb.sheetnames

                input_sheet = wb[sheet_names[0]]

                output_sheet = wb[sheet_names[1]]

                # ---------------------------------------------------
                # INPUT VALUES
                # ---------------------------------------------------

                input_sheet["C2"] = solar_capacity
                input_sheet["C3"] = plant_load

                input_sheet["C9"] = float(
                    clean_number(transmission_charges)
                )

                input_sheet["C14"] = float(
                    clean_number(contract_demand)
                )

                input_sheet["C15"] = energy_rate
                input_sheet["C16"] = demand_charge_rate
                input_sheet["C17"] = wheeling_charge_rate
                input_sheet["C18"] = fac_rate
                input_sheet["C19"] = tax_rate
                input_sheet["C20"] = power_factor

                input_sheet["C21"] = float(
                    clean_number(
                        highest_recorded_msedcl_demand
                    )
                )

                input_sheet["C22"] = electricity_duty

                input_sheet["H25"] = float(
                    current_month_generation
                )

                input_sheet["K26"] = float(a_zone)
                input_sheet["L26"] = float(b_zone)
                input_sheet["M26"] = float(c_zone)
                input_sheet["N26"] = float(d_zone)

                input_sheet["C30"] = float(
                    clean_number(billed_demand)
                )

                input_sheet["C40"] = float(
                    clean_number(reference_units)
                )

                # ---------------------------------------------------
                # OUTPUT VALUES
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
                # SAVE OUTPUT
                # ---------------------------------------------------

                output = BytesIO()

                wb.save(output)

                output.seek(0)

                st.success(
                    "✅ Report Generated Successfully"
                )

                # ---------------------------------------------------
                # DOWNLOAD BUTTON
                # ---------------------------------------------------

                st.download_button(
                    label="⬇ Download Excel Report",
                    data=output,
                    file_name="Before_After_Solar_Report.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-"
                        "officedocument.spreadsheetml.sheet"
                    )
                )

                # ---------------------------------------------------
                # SHOW SECOND SHEET DATA DIRECTLY
                # ---------------------------------------------------

                st.markdown("---")

                st.header(
                    "📊 Output Sheet Preview"
                )

                # ---------------------------------------------------
                # CONVERT SECOND SHEET TO DATAFRAME
                # ---------------------------------------------------

                data = output_sheet.values

                data_list = list(data)

                df = pd.DataFrame(data_list)

                # ---------------------------------------------------
                # DISPLAY SECOND SHEET
                # ---------------------------------------------------

                st.dataframe(
                    df,
                    use_container_width=True,
                    height=800
                )

            except Exception as excel_error:

                st.error(
                    f"Excel Generation Error: "
                    f"{excel_error}"
                )

    except Exception as e:

        st.error(
            f"PDF Processing Error: {e}"
        )
