import streamlit as st

st.title("DISCOM Bill Analysis App")

uploaded_file = st.file_uploader(
    "Upload Electricity Bill PDF",
    type=["pdf"]
)

if uploaded_file:
    st.success("PDF Uploaded Successfully")
