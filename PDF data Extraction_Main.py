# %%
import os
import dotenv
import pandas as pd
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential

# 🔑 Load Azure Form Recognizer credentials
dotenv_path = r"advanalyticsformbuilderinstance.env"
dotenv.load_dotenv(dotenv_path, override=True)

fr_endpoint = os.environ["AZURE_FORMRECOGNIZER_URL"]
fr_key = os.environ["AZURE_FORMRECOGNIZER_KEY"]

document_analysis_client = DocumentAnalysisClient(
    endpoint=fr_endpoint, credential=AzureKeyCredential(fr_key)
)

# 📂 Folder containing PDFs
pdf_folder = r"D:\Projects\Data Extraction 8th August\OneDrive_1_8-13-2025\GW"

# %%
# Process each PDF in the folder
for file in os.listdir(pdf_folder):
    if file.lower().endswith(".pdf"):
        pdf_path = os.path.join(pdf_folder, file)
        excel_filename = os.path.splitext(file)[0] + ".xlsx"
        excel_path = os.path.join(pdf_folder, excel_filename)

        print(f"📄 Processing: {file}")

        # Analyze PDF with Azure Form Recognizer
        with open(pdf_path, "rb") as f:
            poller = document_analysis_client.begin_analyze_document(
                "prebuilt-layout", f.read()
            )
            result = poller.result()

        if not result.tables:
            print(f"⚠️ No tables found in {file}")
            continue

        # Write each table to a separate sheet
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            for idx, table in enumerate(result.tables, start=1):
                # Get page number
                page_number = (
                    table.bounding_regions[0].page_number
                    if table.bounding_regions
                    else 1
                )

                # Create a matrix for table data
                max_columns = table.column_count
                table_matrix = [
                    ["" for _ in range(max_columns)] for _ in range(table.row_count)
                ]
                for cell in table.cells:
                    table_matrix[cell.row_index][cell.column_index] = cell.content

                # Convert to DataFrame
                df = pd.DataFrame(table_matrix)

                # Sheet name: limit to 31 chars (Excel limit)
                sheet_name = f"Pg{page_number}_Tbl{idx}"
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]

                # Save table
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"✅ Extracted {len(result.tables)} tables → {excel_filename}")

# %%
