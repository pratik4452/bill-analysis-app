import streamlit as st
import pdfplumber
import re
import os
import pandas as pd
import plotly.express as px
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

def safe_float(sheet, cell):

    try:

        value = sheet[cell].value

        if value is None:
            return 0

        return float(
            str(value)
            .replace(",", "")
        )

    except:

        return 0

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
                # AUTO SELECT SHEETS
                # ---------------------------------------------------

                sheet_names = wb.sheetnames

                input_sheet = wb[sheet_names[0]]

                if len(sheet_names) > 1:

                    output_sheet = wb[sheet_names[1]]

                else:

                    output_sheet = wb[sheet_names[0]]

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
                # DASHBOARD
                # ---------------------------------------------------

                st.markdown("---")

                st.header(
                    "📊 Before Solar vs After Solar Dashboard"
                )

                # ---------------------------------------------------
                # FETCH VALUES FROM OUTPUT SHEET
                # ---------------------------------------------------

                before_total_bill = safe_float(
                    output_sheet,
                    "C32"
                )

                after_total_bill = safe_float(
                    output_sheet,
                    "D32"
                )

                monthly_savings = safe_float(
                    output_sheet,
                    "D35"
                )

                saving_percentage = safe_float(
                    output_sheet,
                    "D36"
                )

                before_demand = safe_float(
                    output_sheet,
                    "C15"
                )

                after_demand = safe_float(
                    output_sheet,
                    "D15"
                )

                before_wheeling = safe_float(
                    output_sheet,
                    "C16"
                )

                after_wheeling = safe_float(
                    output_sheet,
                    "D16"
                )

                before_energy = safe_float(
                    output_sheet,
                    "C18"
                )

                after_energy = safe_float(
                    output_sheet,
                    "D18"
                )

                before_fac = safe_float(
                    output_sheet,
                    "C19"
                )

                after_fac = safe_float(
                    output_sheet,
                    "D19"
                )

                before_tax = safe_float(
                    output_sheet,
                    "C22"
                )

                after_tax = safe_float(
                    output_sheet,
                    "D22"
                )

                # ---------------------------------------------------
                # KPI CARDS
                # ---------------------------------------------------

                k1, k2, k3, k4 = st.columns(4)

                with k1:

                    st.metric(
                        "⚡ Before Solar Bill",
                        f"₹ {before_total_bill:,.0f}"
                    )

                with k2:

                    st.metric(
                        "☀ After Solar Bill",
                        f"₹ {after_total_bill:,.0f}"
                    )

                with k3:

                    st.metric(
                        "💰 Total Savings",
                        f"₹ {monthly_savings:,.0f}"
                    )

                with k4:

                    st.metric(
                        "📉 Savings %",
                        f"{saving_percentage:.1f}%"
                    )

                st.markdown("---")

                # ---------------------------------------------------
                # BILL COMPARISON CHART
                # ---------------------------------------------------

                comparison_df = pd.DataFrame({

                    "Bill": [
                        "Before Solar",
                        "After Solar"
                    ],

                    "Amount": [
                        before_total_bill,
                        after_total_bill
                    ]

                })

                comparison_fig = px.bar(
                    comparison_df,
                    x="Bill",
                    y="Amount",
                    text="Amount",
                    title="Before vs After Solar Bill"
                )

                comparison_fig.update_traces(
                    texttemplate='₹ %{text:,.0f}',
                    textposition='outside'
                )

                comparison_fig.update_layout(
                    height=500
                )

                st.plotly_chart(
                    comparison_fig,
                    use_container_width=True
                )

                # ---------------------------------------------------
                # CHARGES COMPARISON
                # ---------------------------------------------------

                st.markdown("---")

                st.subheader(
                    "💵 Charges Comparison"
                )

                charges_df = pd.DataFrame({

                    "Particulars": [

                        "Demand Charges",
                        "Wheeling Charges",
                        "Energy Charges",
                        "FAC",
                        "Tax"

                    ],

                    "Before Solar": [

                        before_demand,
                        before_wheeling,
                        before_energy,
                        before_fac,
                        before_tax

                    ],

                    "After Solar": [

                        after_demand,
                        after_wheeling,
                        after_energy,
                        after_fac,
                        after_tax

                    ]

                })

                charges_fig = px.bar(

                    charges_df,

                    x="Particulars",

                    y=[
                        "Before Solar",
                        "After Solar"
                    ],

                    barmode="group",

                    title="Charges Comparison"
                )

                st.plotly_chart(
                    charges_fig,
                    use_container_width=True
                )

                # ---------------------------------------------------
                # SAVINGS DONUT CHART
                # ---------------------------------------------------

                st.markdown("---")

                donut_df = pd.DataFrame({

                    "Category": [
                        "Savings",
                        "Remaining Bill"
                    ],

                    "Amount": [
                        monthly_savings,
                        after_total_bill
                    ]

                })

                donut_fig = px.pie(

                    donut_df,

                    names="Category",

                    values="Amount",

                    hole=0.5,

                    title="Savings Distribution"
                )

                st.plotly_chart(
                    donut_fig,
                    use_container_width=True
                )

                # ---------------------------------------------------
                # DETAILED TABLE
                # ---------------------------------------------------

                st.markdown("---")

                st.subheader(
                    "📋 Detailed Comparison Table"
                )

                detailed_df = pd.DataFrame({

                    "Particulars": [

                        "Demand Charges",
                        "Wheeling Charges",
                        "Energy Charges",
                        "FAC",
                        "Tax",
                        "Total Bill"

                    ],

                    "Before Solar": [

                        before_demand,
                        before_wheeling,
                        before_energy,
                        before_fac,
                        before_tax,
                        before_total_bill

                    ],

                    "After Solar": [

                        after_demand,
                        after_wheeling,
                        after_energy,
                        after_fac,
                        after_tax,
                        after_total_bill

                    ]

                })

                st.dataframe(
                    detailed_df,
                    use_container_width=True
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

            except Exception as excel_error:

                st.error(
                    f"Excel Generation Error: "
                    f"{excel_error}"
                )

    except Exception as e:

        st.error(
            f"PDF Processing Error: {e}"
        )
