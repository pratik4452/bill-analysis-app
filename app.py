# ---------------------------------------------------
# POWER FACTOR EXTRACTION
# ---------------------------------------------------

power_factor = ""

# Multiple regex patterns for different bill formats

pf_patterns = [

    # P.F. 0.99
    r'P\.?\s*F\.?\s*[:\-]?\s*([\d\.]+)',

    # PF 0.99
    r'PF\s*[:\-]?\s*([\d\.]+)',

    # Power Factor 0.99
    r'Power\s*Factor\s*[:\-]?\s*([\d\.]+)',

    # Avg. Power Factor 0.99
    r'Avg\.?\s*Power\s*Factor\s*[:\-]?\s*([\d\.]+)',

    # Average Power Factor 0.99
    r'Average\s*Power\s*Factor\s*[:\-]?\s*([\d\.]+)',

    # Power Factor : 1
    r'Power\s*Factor\s*[:\-]?\s*([\d]+)',

    # PF : 1
    r'PF\s*[:\-]?\s*([\d]+)',

]

# Loop through patterns

for pattern in pf_patterns:

    match = re.search(
        pattern,
        text,
        re.IGNORECASE
    )

    if match:

        power_factor = match.group(1)

        break

# ---------------------------------------------------
# CLEAN VALUE
# ---------------------------------------------------

if power_factor:

    power_factor = (
        str(power_factor)
        .replace("%", "")
        .replace(",", "")
        .strip()
    )

# ---------------------------------------------------
# DEFAULT VALUE
# ---------------------------------------------------

if not power_factor:

    power_factor = "1"

# ---------------------------------------------------
# DEBUG DISPLAY
# ---------------------------------------------------

st.write(
    "Detected Power Factor:",
    power_factor
)

# ---------------------------------------------------
# WRITE TO EXCEL
# ---------------------------------------------------

input_sheet["C20"] = float(power_factor)
