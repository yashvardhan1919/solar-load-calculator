import streamlit as st

from bill_backend import process_bill


st.set_page_config(page_title="Solar Load Calculator")

st.markdown(
    """
    <style>
    .stApp { background: white; color: black; }
    header, footer { visibility: hidden; }
    .block-container { max-width: 600px; padding-top: 20px; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Solar Load Calculator")

uploaded_file = st.file_uploader("Upload image", type=["jpg", "jpeg", "png"])

if uploaded_file and st.button("Generate Excel"):
    data, output_bytes = process_bill(uploaded_file)
    st.write("Consumer Number:", data["Consumer Number"])
    st.write("Name:", data["Name"])
    st.write("Units:", data["Units"])
    st.write("Amount:", data["Amount"])
    st.download_button(
        "Download Excel",
        output_bytes,
        "output.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
