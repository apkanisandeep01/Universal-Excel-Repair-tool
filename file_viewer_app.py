import streamlit as st
import pandas as pd
import io
from pathlib import Path

st.set_page_config(page_title="Excel Repair Tool", layout="centered")
st.title("📘🔧 Excel Repair & Export Assistant")

# Allowed file types
allowed_types = ["csv", "xlsx", "xls", "xlsm", "xlsb", "odf", "ods", "odt"]

# Path to Downloads folder
downloads_path = Path.home() / "Downloads"

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
        # Step 2: First read file WITHOUT headers (raw load)
        if file_extension == "csv":
            temp_df = pd.read_csv(uploaded_file, header=None, on_bad_lines="skip")
        elif file_extension == "xls":
            temp_df = pd.read_excel(uploaded_file, header=None, engine="xlrd")
        elif file_extension == "xlsb":
            temp_df = pd.read_excel(uploaded_file, header=None, engine="pyxlsb")
        else:
            temp_df = pd.read_excel(uploaded_file, header=None, engine="openpyxl")

        # Step 2a: If file is empty / no columns
        if temp_df.empty:
            st.warning("⚠️ The file has no valid data or columns. Showing raw file content instead.")
            uploaded_file.seek(0)
            raw_content = uploaded_file.read(500)  # first 500 bytes
            try:
                st.text(raw_content.decode("utf-8", errors="ignore"))
            except Exception:
                st.text(str(raw_content))
            st.stop()

        # Step 2b: Ask which row to use as header (1-based index for user)
        st.subheader("⚙️ Choose Header Row")
        header_row = st.number_input(
            "Select the row number to use as header (First row as Default)",
            min_value=1,
            max_value=len(temp_df),
            value=1
        ) - 1  # Convert to 0-based index for pandas

        # Step 2c: Reload file with selected header row
        uploaded_file.seek(0)  # reset file pointer before re-reading
        if file_extension == "csv":
            df = pd.read_csv(uploaded_file, header=header_row, on_bad_lines="skip")
        elif file_extension == "xls":
            df = pd.read_excel(uploaded_file, header=header_row, engine="xlrd")
        elif file_extension == "xlsb":
            df = pd.read_excel(uploaded_file, header=header_row, engine="pyxlsb")
        else:
            df = pd.read_excel(uploaded_file, header=header_row, engine="openpyxl")
        
        # Step 3: Ask number of rows to preview
        
        num_rows = st.number_input(
            "How many rows do you want to view?",
            min_value=1,
            max_value=len(df),
            value=min(5, len(df))
        )
        st.subheader(f"Total No of rows are: {df.shape[0]} and Total no columns are: {df.shape[1]}")
        st.subheader(f"👀 Preview of first {num_rows} rows")
        st.dataframe(df.head(num_rows))
        

        # Step 4: Create new filename with "_repaired"
        repaired_filename = f"{original_name}_repaired.xlsx"
        repaired_path = downloads_path / repaired_filename

        # Step 5: Save to Downloads folder
        df.to_excel(repaired_path, index=False)

        # Step 6: Prepare in-memory file for browser download
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)

        st.download_button(
            label="✨ Download repaired file",
            data=output,
            file_name=repaired_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Step 7: Button to upload another file
        if st.button("🔄 Upload Another File"):
            st.session_state.uploaded_file = None
            st.rerun()

    except Exception as e:
        st.error(f"⚠️ Error reading file: {e}")
