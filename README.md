# 📘🔧 Universal Excel Repair Assistant

Ever been blocked by the dreaded **"Excel cannot open the file..."** error? Or opened a CSV only to find a chaotic mess of data with no clear headers?

This simple yet powerful tool is your first-aid kit for stubborn, misformatted, or lightly corrupted spreadsheet files. It helps you bypass common errors, fix structural problems, and export your data into a clean, universally compatible `.xlsx` format.


## ✨ Try the Live App Now\!

### [🚀 Launch the Excel Repair Tool](https://excelrepairtool.streamlit.app/)


-----

## Why You'll Love This Tool

Many spreadsheet errors aren't due to deep file corruption but simple structural issues. This tool is expertly designed to handle the most common frustrations:

  * **✔️ Bypasses Common Errors:** Uses robust data-loading libraries to read files that standard Excel or Google Sheets might reject.
  * **✔️ Fixes Broken Headers:** The \#1 issue with messy data is a misplaced header. Our tool lets you visually identify and assign any row as the official column header, instantly restructuring your entire file.
  * **✔️ Unifies File Formats:** Have a `.xlsb`, `.ods`, or an old `.xls` file? The tool effortlessly reads a wide variety of formats and converts them into the modern `.xlsx` standard.
  * **✔️ Safe & Secure:** Your file is processed in memory and is never stored on a server. It's your data, your control.

-----

## 📖 A Simple 4-Step Rescue Mission

Repairing your file is incredibly straightforward.

1.  **📂 Upload Your File**
    Drag and drop or browse to select your problematic spreadsheet. It supports a wide range of formats.

2.  **⚙️ Set the Header Row**
    The app will load a raw version of your data. Simply enter the row number that *should* be your header. The default is `1`.

3.  **👀 Preview the Cleaned Data**
    Instantly see a preview of your table with the correct headers applied. You can verify that your data is now clean, structured, and ready for use.

4.  **✨ Download the Repaired File**
    Click the "Download" button to save your repaired data as a new, clean `.xlsx` file. The new file will be named `[your-original-name]_repaired.xlsx`.

-----

## 🗂️ Supported File Types

The tool is built to be flexible and accepts most common spreadsheet and data file formats:

  * `xlsx` (Modern Excel)
  * `xls` (Legacy Excel)
  * `xlsm` (Macro-Enabled Excel)
  * `xlsb` (Excel Binary)
  * `csv` (Comma-Separated Values)
  * `ods`, `odf`, `odt` (OpenDocument Formats)

-----

## 🛠️ Built With

This application is powered by industry-standard open-source technologies:

  * **Language:** Python
  * **Web Framework:** Streamlit
  * **Data Processing:** Pandas

-----

## 💻 Running Locally

For developers who wish to run this tool on their own machine:

1.  **Clone the Repository:**

    ```bash
    git clone https://github.com/apkanisandeep01/Universal-Excel-Repair-tool.git
    cd Universal-Excel-Repair-tool
    ```

2.  **Install Dependencies:**

    ```bash
    pip install streamlit pandas openpyxl xlrd pyxlsb
    ```

3.  **Run the App:**

    ```bash
    streamlit run app.py
    ```

    Your browser will open a new tab with the application running locally.

-----

## 🤝 Feedback & Contributions

Find this tool useful? Have an idea for a new feature or found a bug? Please feel free to **open an issue** on GitHub. Contributions are always welcome\!
