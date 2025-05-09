
import streamlit as st
import pandas as pd

st.set_page_config(page_title="DSL to Fiber Dashboard", layout="wide")
st.title("📊 DSL to Fiber Conversion Dashboard")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    selected_sheets = st.multiselect("Select sheet(s) to include in analysis:", sheet_names, default=sheet_names)

    all_data = []
    for sheet in selected_sheets:
        df = xls.parse(sheet, skiprows=1)
        df.dropna(how="all", inplace=True)
        df = df.dropna(axis=1, how="all")
        df.columns = df.columns.map(str)  # Fix: ensure all column names are strings
        df["__sheet__"] = sheet
        all_data.append(df)

    if all_data:
        df = pd.concat(all_data, ignore_index=True)
        st.subheader("🗂️ Combined Data from Selected Sheets")
        with st.expander("Preview Combined Data"):
            st.dataframe(df)

        address_cols = [col for col in df.columns if "address" in col.lower()]
        status_cols = [col for col in df.columns if "status" in col.lower()]

        if address_cols:
            address_col = address_cols[0]
            address_vals = df[address_col].dropna().unique().tolist()
            selected_addresses = st.multiselect(f"Filter by {address_col}", options=address_vals, default=address_vals)
            df = df[df[address_col].isin(selected_addresses)]

        if status_cols:
            status_col = status_cols[0]
            status_vals = df[status_col].dropna().unique().tolist()
            selected_status = st.multiselect(f"Filter by {status_col}", options=status_vals, default=status_vals)
            df = df[df[status_col].isin(selected_status)]

        st.markdown("### 📈 Filtered Results")
        st.dataframe(df)

        if status_cols:
            chart_data = df[status_col].value_counts().reset_index()
            chart_data.columns = [status_col, "Count"]
            st.bar_chart(chart_data.set_index(status_col))

        csv = df.to_csv(index=False).encode("utf-8")
        st.download_button("Download filtered data as CSV", csv, "filtered_data.csv", "text/csv")
    else:
        st.warning("No valid data found in selected sheets.")

else:
    st.info("👈 Upload an Excel file to get started.")
