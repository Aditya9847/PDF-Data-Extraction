# Installation

1. Create and activate a virtual environment
    - Install `requirements.txt`

2. Configure Azure Form Recognizer
    - Generate Access key from Azure Document Intelligence Studio.
    
# Ready to run!

# Steps: -
1. Run PDF_Data_Extraction_Main.py
    - Uses Azure Form Recognizer to extract structured data from PDF or document files.
    - Credentials are automatically loaded from the .env file.
    - Output: Structured Excel file with extracted data.
    - Adjust the PDF folder path inside the script if necessary.
    - Run excel_data_quality_tools.py

2. Run excel_data_quality_tools.py
    - Perform QA/QC Checks
    - This consolidated script contains three independent tools you can run as needed:

     (a) highlight_excel_errors(file_path)
        - Detects formatting and typographical issues, such as:
        - Leading/trailing spaces (" 123", "123 ")
        - Numbers with embedded spaces ("0 008", "7 368")
        - Numbers using commas instead of decimals
        - Characters like "I", "l", or "o" used instead of digits
        - Patterns like "<.008" (invalid numeric notation)
        - Highlights suspicious cells in red — no automatic replacement.

    b) highlight_outliers(file_path)
        🔸 Highlights inconsistent or outlier values in yellow.

        - Performs outlier detection and pattern consistency checks:
        - Begins checking from row 3 onwards
        - Column B should contain the variable name (identifier)
        - Each row represents one variable; data values are from column C to the last used column
        - Ignores unit columns (containing text like mg/L, ug/L, uS, pH)
        - Detects inconsistent patterns — e.g., some values use <0.002 while others don’t
        - Uses Median Absolute Deviation (MAD) for robust outlier detection

    c) fix_dates_in_excel(file_path, rows_with_dates=None, cols_with_dates=None)
        - Standardizes and validates date columns:
        - Converts detected dates to MM/DD/YYYY format (text)
        - Keeps non-date values unchanged
        - Requires specifying row and column ranges containing date fields
        - Saves output as a new Excel file named *_corrected.xlsx*

# Recommended Workflow
    - Extract data → PDF_Data_Extraction_Main.py
    - Check for format/entry issues → excel_data_quality_tools.highlight_excel_errors()
    - Detect outliers → excel_data_quality_tools.highlight_outliers()
    - Fix and standardize date columns → excel_data_quality_tools.fix_dates_in_excel()
