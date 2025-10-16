import streamlit as st
import pandas as pd
import io
from pathlib import Path
import chardet

st.set_page_config(page_title="Excel Repair Tool", layout="centered")
st.title("📘🔧 Excel Repair & Export Assistant")

# Allowed file types
allowed_types = ["csv", "xlsx", "xls", "xlsm", "xlsb", "odf", "ods", "odt"]

# Initialize session state
if "uploaded_file" not in st.session_state:
    st.session_state.uploaded_file = None

# Step 1: If no file uploaded yet, show uploader
if st.session_state.uploaded_file is None:
    uploaded_file = st.file_uploader("📂 Upload your Excel or CSV file", type=allowed_types)
    if uploaded_file is not None:
        st.session_state.uploaded_file = uploaded_file
        st.rerun()
else:
    uploaded_file = st.session_state.uploaded_file

    # Capture original file name (without extension)
    original_name = Path(uploaded_file.name).stem
    file_extension = Path(uploaded_file.name).suffix.lower().strip(".")

    try:
        # Step 2: Handle CSV files with robust encoding detection
        if file_extension == "csv":
            # Read the entire file bytes for encoding detection
            file_bytes = uploaded_file.read()
            detection = chardet.detect(file_bytes)
            detected_encoding = detection.get("encoding") or "utf-8"

            st.info(f"Detected file encoding: {detected_encoding}")

            # Create BytesIO for pandas to read again
            file_buffer = io.BytesIO(file_bytes)

            # Load raw data without headers
            temp_df = pd.read_csv(
                file_buffer,
                header=None,
                on_bad_lines="skip",
                encoding=detected_encoding,
                low_memory=False
            )

        elif file_extension == "xls":
            temp_df = pd.read_excel(uploaded_file, header=None, engine="xlrd")
        elif file_extension == "xlsb":
            temp_df = pd.read_excel(uploaded_file, header=None, engine="pyxlsb")
        else:
            temp_df = pd.read_excel(uploaded_file, header=None, engine="openpyxl")

        if temp_df.empty:
            st.warning("⚠️ The file has no valid data or columns. Showing raw file content instead.")
            st.stop()

        # Step 2b: Choose header row
        st.subheader("⚙️ Choose Header Row")
        header_row = st.number_input(
            "Select the row number to use as header (First row as Default)",
            min_value=1,
            max_value=len(temp_df),
            value=1
        ) - 1

        # Step 2c: Reload file with selected header
        if file_extension == "csv":
            file_buffer.seek(0)
            df = pd.read_csv(
                file_buffer,
                header=header_row,
                on_bad_lines="skip",
                encoding=detected_encoding,
                low_memory=False
            )
        elif file_extension == "xls":
            df = pd.read_excel(uploaded_file, header=header_row, engine="xlrd")
        elif file_extension == "xlsb":
            df = pd.read_excel(uploaded_file, header=header_row, engine="pyxlsb")
        else:
            df = pd.read_excel(uploaded_file, header=header_row, engine="openpyxl")

        # Step 3: Preview
        num_rows = st.number_input(
            "How many rows do you want to view?",
            min_value=1,
            max_value=len(df),
            value=min(5, len(df))
        )
        st.subheader(f"Total rows: {df.shape[0]} | Total columns: {df.shape[1]}")
        st.subheader(f"👀 Preview of first {num_rows} rows")
        st.dataframe(df.head(num_rows))

        # Step 4: Prepare file for download
        repaired_filename = f"{original_name}_repaired.xlsx"
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            label="✨ Download repaired file",
            data=output,
            file_name=repaired_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Step 5: Upload another file
        if st.button("🔄 Upload Another File"):
            st.session_state.uploaded_file = None
            st.rerun()

    except Exception as e:
        st.error(f"⚠️ Error reading file: {e}")
