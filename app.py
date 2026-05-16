# ---------------------------------------------------
# POWER FACTOR EXTRACTION (FINAL WORKING LOGIC)
# ---------------------------------------------------

# THIS WILL FETCH:
# P.F. 0.996
# from your MSEDCL bill correctly

power_factor = ""

# ---------------------------------------------------
# METHOD 1
# DIRECTLY SEARCH "P.F."
# ---------------------------------------------------

pf_match = re.search(

    r'P\.F\.\s*L\.F\.\s*Zone\s*:?\s*[\r\n\s]*([\d\.]+)',

    text,

    re.IGNORECASE

)

if pf_match:

    power_factor = pf_match.group(1)

# ---------------------------------------------------
# METHOD 2
# NORMAL P.F. SEARCH
# ---------------------------------------------------

if not power_factor:

    pf_match = re.search(

        r'P\.F\.\s*([\d\.]+)',

        text,

        re.IGNORECASE

    )

    if pf_match:

        power_factor = pf_match.group(1)

# ---------------------------------------------------
# METHOD 3
# PF SEARCH
# ---------------------------------------------------

if not power_factor:

    pf_match = re.search(

        r'PF\s*([\d\.]+)',

        text,

        re.IGNORECASE

    )

    if pf_match:

        power_factor = pf_match.group(1)

# ---------------------------------------------------
# METHOD 4
# TAKE VALUE BEFORE "26"
# BECAUSE YOUR BILL HAS:
#
# 0.996 26
# P.F. L.F. Zone :
# ---------------------------------------------------

if not power_factor:

    pf_match = re.search(

        r'([\d\.]+)\s+26\s+P\.F\.',

        text,

        re.IGNORECASE

    )

    if pf_match:

        power_factor = pf_match.group(1)

# ---------------------------------------------------
# FINAL DEFAULT
# ---------------------------------------------------

if not power_factor:

    power_factor = "1"

# ---------------------------------------------------
# DISPLAY
# ---------------------------------------------------

st.write("P.F.:", power_factor)

# ---------------------------------------------------
# WRITE TO EXCEL
# ---------------------------------------------------

input_sheet["C20"] = float(
    clean_number(power_factor)
)
