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
        # CURRENT MONTH GENERATION
        # ---------------------------------------------------

        current_month_generation = ""

        generation_patterns = [

            r'Units\s*Offset\s*Against\s*Drawal.*?Current\s*Month\s*Generation\s*([\d,]+)',

            r'Current\s*Month\s*Generation\s*([\d,]+)',

            r'Current\s*Month\s*\n\s*Generation\s*\n\s*([\d,]+)',

            r'Generation\s*\n\s*([\d,]+)',

        ]

        for pattern in generation_patterns:

            match = re.search(
                pattern,
                text,
                re.IGNORECASE | re.DOTALL
            )

            if match:

                current_month_generation = match.group(1)

                break

        # ---------------------------------------------------
        # TOD ZONE VALUES
        # ---------------------------------------------------

        a_zone = extract_value(
            r'A\s*Zone\s*([\d,\.]+)',
            text
        )

        b_zone = extract_value(
            r'B\s*Zone\s*([\d,\.]+)',
            text
        )

        c_zone = extract_value(
            r'C\s*Zone\s*([\d,\.]+)',
            text
        )

        d_zone = extract_value(
            r'D\s*Zone\s*([\d,\.]+)',
            text
        )

        # ---------------------------------------------------
        # DISPLAY DATA
        # ---------------------------------------------------

        st.markdown("## 📋 Extracted Bill Data")

        col1, col2 = st.columns(2)

        with col1:

            st.write("Contract Demand:", contract_demand)

            st.write(
                "Highest Recorded MSEDCL Demand:",
                highest_recorded_msedcl_demand
            )

            st.write(
                "Transmission Charges:",
                transmission_charges
            )

            st.write(
                "Solar Units at Consumption End:",
                current_month_generation
            )

        with col2:

            st.write("A Zone:", a_zone)
            st.write("B Zone:", b_zone)
            st.write("C Zone:", c_zone)
            st.write("D Zone:", d_zone)

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
                    float(clean_number(
                        highest_recorded_msedcl_demand
                    ))
                    if highest_recorded_msedcl_demand else 0
                )

                ws["C22"] = electricity_duty

                # ---------------------------------------------------
                # SOLAR UNITS AT CONSUMPTION END
                # ---------------------------------------------------

                ws["H25"] = (
                    float(clean_number(
                        current_month_generation
                    ))
                    if current_month_generation else 0
                )

                # ---------------------------------------------------
                # TOD ZONES
                # ---------------------------------------------------

                ws["K26"] = (
                    float(clean_number(a_zone))
                    if a_zone else 0
                )

                ws["L26"] = (
                    float(clean_number(b_zone))
                    if b_zone else 0
                )

                ws["M26"] = (
                    float(clean_number(c_zone))
                    if c_zone else 0
                )

                ws["N26"] = (
                    float(clean_number(d_zone))
                    if d_zone else 0
                )

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
