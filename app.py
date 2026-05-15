import streamlit as st
import pdfplumber
import re
import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

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
                # ADVANCED PROFESSIONAL DASHBOARD
                # ---------------------------------------------------

                st.markdown("---")

                st.header("📊 Executive Solar Savings Dashboard")

                # ---------------------------------------------------
                # CALCULATIONS
                # ---------------------------------------------------

                reference_units_value = float(
                    clean_number(reference_units)
                )

                solar_generation_value = float(
                    current_month_generation
                )

                before_energy_charges = (
                    reference_units_value
                    * energy_rate
                )

                after_units = (
                    reference_units_value
                    - solar_generation_value
                )

                if after_units < 0:

                    after_units = 0

                after_energy_charges = (
                    after_units
                    * energy_rate
                )

                before_demand_charges = (
                    float(clean_number(contract_demand))
                    * demand_charge_rate
                )

                after_demand_charges = (
                    float(clean_number(
                        highest_recorded_msedcl_demand
                    ))
                    * demand_charge_rate
                )

                before_total_bill = (
                    before_energy_charges
                    + before_demand_charges
                )

                after_total_bill = (
                    after_energy_charges
                    + after_demand_charges
                    + float(debit_bill_adjustment)
                    + float(grid_support_charges)
                )

                monthly_savings = (
                    before_total_bill
                    - after_total_bill
                )

                saving_percentage = 0

                if before_total_bill > 0:

                    saving_percentage = (
                        monthly_savings
                        / before_total_bill
                    ) * 100

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
                        "💰 Monthly Savings",
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

                    "Category": [
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
                    x="Category",
                    y="Amount",
                    text="Amount",
                    title="Before vs After Solar Bill Comparison"
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
                # COST BREAKUP
                # ---------------------------------------------------

                st.markdown("---")

                st.subheader("💵 Cost Breakup Analysis")

                cost_df = pd.DataFrame({

                    "Charges": [
                        "Energy Charges",
                        "Demand Charges",
                        "Debit Adjustment",
                        "Grid Support"
                    ],

                    "Amount": [
                        after_energy_charges,
                        after_demand_charges,
                        float(debit_bill_adjustment),
                        float(grid_support_charges)
                    ]

                })

                cost_fig = px.pie(
                    cost_df,
                    names="Charges",
                    values="Amount",
                    title="After Solar Cost Distribution"
                )

                st.plotly_chart(
                    cost_fig,
                    use_container_width=True
                )

                # ---------------------------------------------------
                # WATERFALL CHART
                # ---------------------------------------------------

                st.markdown("---")

                st.subheader("📉 Savings Waterfall")

                waterfall = go.Figure(go.Waterfall(

                    name="Savings",

                    orientation="v",

                    measure=[
                        "absolute",
                        "relative",
                        "relative",
                        "total"
                    ],

                    x=[
                        "Before Solar",
                        "Solar Savings",
                        "Additional Charges",
                        "After Solar"
                    ],

                    textposition="outside",

                    y=[
                        before_total_bill,
                        -monthly_savings,
                        (
                            float(debit_bill_adjustment)
                            +
                            float(grid_support_charges)
                        ),
                        after_total_bill
                    ]

                ))

                waterfall.update_layout(
                    title="Solar Savings Waterfall Analysis",
                    height=500
                )

                st.plotly_chart(
                    waterfall,
                    use_container_width=True
                )

                # ---------------------------------------------------
                # TOD ZONE ANALYSIS
                # ---------------------------------------------------

                st.markdown("---")

                st.subheader("⚡ TOD Zone Analysis")

                zone_df = pd.DataFrame({

                    "Zone": [
                        "A Zone",
                        "B Zone",
                        "C Zone",
                        "D Zone"
                    ],

                    "Units": [
                        float(a_zone),
                        float(b_zone),
                        float(c_zone),
                        float(d_zone)
                    ]

                })

                zone_fig = px.bar(
                    zone_df,
                    x="Zone",
                    y="Units",
                    text="Units",
                    title="TOD Zone Consumption"
                )

                zone_fig.update_traces(
                    texttemplate='%{text:,.0f}',
                    textposition='outside'
                )

                zone_fig.update_layout(
                    height=500
                )

                st.plotly_chart(
                    zone_fig,
                    use_container_width=True
                )

                # ---------------------------------------------------
                # ENERGY SUMMARY
                # ---------------------------------------------------

                st.markdown("---")

                st.subheader("⚡ Energy Summary")

                e1, e2, e3 = st.columns(3)

                with e1:

                    st.success(
                        f"""
                        ### Reference Consumption

                        {reference_units_value:,.0f} kWh
                        """
                    )

                with e2:

                    st.success(
                        f"""
                        ### Solar Generation

                        {solar_generation_value:,.0f} kWh
                        """
                    )

                with e3:

                    net_grid_units = (
                        reference_units_value
                        - solar_generation_value
                    )

                    if net_grid_units < 0:

                        net_grid_units = 0

                    st.success(
                        f"""
                        ### Net Grid Consumption

                        {net_grid_units:,.0f} kWh
                        """
                    )

                # ---------------------------------------------------
                # DETAILED TABLE
                # ---------------------------------------------------

                st.markdown("---")

                st.subheader(
                    "📋 Detailed Charges Comparison"
                )

                charges_df = pd.DataFrame({

                    "Particulars": [

                        "Energy Charges",
                        "Demand Charges",
                        "Transmission Charges",
                        "Debit Adjustment",
                        "Grid Support Charges",
                        "Total Bill"

                    ],

                    "Before Solar": [

                        round(before_energy_charges, 2),
                        round(before_demand_charges, 2),
                        round(float(clean_number(
                            transmission_charges
                        )), 2),
                        0,
                        0,
                        round(before_total_bill, 2)

                    ],

                    "After Solar": [

                        round(after_energy_charges, 2),
                        round(after_demand_charges, 2),
                        round(float(clean_number(
                            transmission_charges
                        )), 2),
                        round(float(debit_bill_adjustment), 2),
                        round(float(grid_support_charges), 2),
                        round(after_total_bill, 2)

                    ]

                })

                st.dataframe(
                    charges_df,
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
