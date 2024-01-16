import pdfplumber
import pandas as pd

def extract_and_identify_tables(pdf_path, start_page=42, end_page=102):
    # Open the PDF file using pdfplumber
    with pdfplumber.open(pdf_path) as pdf:
        # Initialize an empty list to store tables
        tables = []

        # Iterate through each page in the specified range
        for page_num in range(start_page - 1, min(end_page, len(pdf.pages))):
            # Extract text from the page
            text = pdf.pages[page_num].extract_text()

            # Split text into lines
            lines = text.split('\n')

            # Identify potential table-like structures (very basic example)
            potential_tables = [line.split() for line in lines if len(line.split()) > 1]

            # Convert potential tables to DataFrames
            for table_data in potential_tables:
                table_df = pd.DataFrame([table_data], columns=[f'Column_{i+1}' for i in range(len(table_data))])
                tables.append(table_df)

    return tables

def save_tables_to_excel(tables, excel_file_path='output_tables.xlsx'):
    # Concatenate all tables into a single DataFrame
    combined_table = pd.concat(tables, ignore_index=True)

    # Save the combined table to an Excel file
    combined_table.to_excel(excel_file_path, index=False)
    print(f"Tables saved to {excel_file_path}")

if __name__ == "__main__":
    # Replace 'your_pdf_file.pdf' with the path to your PDF file
    pdf_file_path = "C:\\Users\\savag\\Downloads\\tech_profile_report.pdf"

    # Extract and identify potential tables from pages 42 to 108
    extracted_tables = extract_and_identify_tables(pdf_file_path, start_page=42, end_page=102)

    # Save identified tables to Excel
    save_tables_to_excel(extracted_tables)
