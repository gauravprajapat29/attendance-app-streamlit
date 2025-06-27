import streamlit as st
from logic import process_excel

st.set_page_config(page_title="Attendance Analyzer", layout="centered")

st.title("📊 Attendance Data Analyzer")

# File uploader
uploaded_file = st.file_uploader("📁 Upload Excel File", type=["xls", "xlsx"])

# Year input
year = st.number_input("📅 Enter Year", min_value=2000, max_value=2100, value=2025, step=1)

# Month input
month = st.number_input("🗓️ Enter Month (1-12)", min_value=1, max_value=12, value=5, step=1)

# Submit Button
if st.button("Submit"):
    if uploaded_file is not None:
        try:
            result_df = process_excel(uploaded_file, year, month)
            st.success("✅ Attendance Data Processed Successfully!")
            st.dataframe(result_df, use_container_width=True)
        except Exception as e:
            st.error(f"❌ Error processing the file: {e}")
    else:
        st.warning("⚠️ Please upload a valid Excel file.")
