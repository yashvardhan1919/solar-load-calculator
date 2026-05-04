import streamlit as st

try:
    from bill_backend import process_bill
except Exception as e:
    st.set_page_config(page_title="Solar Load Calculator")
    st.error(f"Startup error: {e}")
    st.stop()


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

if "data" not in st.session_state:
    st.session_state.data = None
if "output_bytes" not in st.session_state:
    st.session_state.output_bytes = None

uploaded_file = st.file_uploader("Upload image", type=["jpg", "jpeg", "png"])

if uploaded_file and st.button("Generate Excel"):
    try:
        data, output_bytes = process_bill(uploaded_file)
        st.session_state.data = data
        st.session_state.output_bytes = output_bytes
    except Exception as e:
        st.error(f"Processing error: {e}")

if st.session_state.data:
    st.write("Consumer Number:", st.session_state.data["Consumer Number"])
    st.write("Name:", st.session_state.data["Name"])
    st.write("Units:", st.session_state.data["Units"])
    st.write("Amount:", st.session_state.data["Amount"])
    st.write("Load:", st.session_state.data["Load"])
    st.write("Connection Type:", st.session_state.data["Connection Type"])
    st.download_button(
        "Download Excel",
        st.session_state.output_bytes,
        "output.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
