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

                input_sheet = wb[wb.sheetnames[0]]

                output_sheet = wb["Bill After Solar_Apr 26"]

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

                input_sheet["C15"] = 8.44
                input_sheet["C16"] = 650
                input_sheet["C17"] = 0.81
                input_sheet["C18"] = 0.50
                input_sheet["C19"] = 0.29
                input_sheet["C20"] = 1

                input_sheet["C21"] = float(
                    clean_number(
                        highest_recorded_msedcl_demand
                    )
                )

                input_sheet["C22"] = "7.50%"

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
                # FORCE RECALCULATION
                # ---------------------------------------------------

                wb.calculation.fullCalcOnLoad = True
                wb.calculation.forceFullCalc = True

                # ---------------------------------------------------
                # SAVE FILE
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
                    "📊 Solar Savings Dashboard"
                )

                before_solar_bill = (
                    output_sheet["C32"].value
                )

                after_solar_bill = (
                    output_sheet["D32"].value
                )

                try:

                    before_solar_bill = float(
                        before_solar_bill
                    )

                except:
                    before_solar_bill = 0

                try:

                    after_solar_bill = float(
                        after_solar_bill
                    )

                except:
                    after_solar_bill = 0

                monthly_savings = (
                    before_solar_bill
                    -
                    after_solar_bill
                )

                savings_percentage = 0

                if before_solar_bill > 0:

                    savings_percentage = (
                        monthly_savings
                        /
                        before_solar_bill
                    ) * 100

                # ---------------------------------------------------
                # KPI CARDS
                # ---------------------------------------------------

                k1, k2, k3 = st.columns(3)

                with k1:

                    st.metric(
                        "💡 Before Solar Bill",
                        f"₹ {before_solar_bill:,.0f}"
                    )

                with k2:

                    st.metric(
                        "⚡ After Solar Bill",
                        f"₹ {after_solar_bill:,.0f}"
                    )

                with k3:

                    st.metric(
                        "💰 Monthly Savings",
                        f"₹ {monthly_savings:,.0f}",
                        f"{savings_percentage:.1f}%"
                    )

                st.markdown("---")

                # ---------------------------------------------------
                # BAR CHART
                # ---------------------------------------------------

                chart_df = pd.DataFrame({

                    "Bill Type": [
                        "Before Solar",
                        "After Solar"
                    ],

                    "Amount": [
                        before_solar_bill,
                        after_solar_bill
                    ]

                })

                fig = px.bar(
                    chart_df,
                    x="Bill Type",
                    y="Amount",
                    text="Amount",
                    title="Before vs After Solar Bill"
                )

                fig.update_traces(
                    texttemplate='₹ %{text:,.0f}',
                    textposition='outside'
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True
                )

                # ---------------------------------------------------
                # PIE CHART
                # ---------------------------------------------------

                pie_df = pd.DataFrame({

                    "Category": [
                        "Savings",
                        "After Solar Bill"
                    ],

                    "Value": [
                        monthly_savings,
                        after_solar_bill
                    ]

                })

                pie_fig = px.pie(
                    pie_df,
                    names="Category",
                    values="Value",
                    title="Solar Savings Distribution"
                )

                st.plotly_chart(
                    pie_fig,
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
