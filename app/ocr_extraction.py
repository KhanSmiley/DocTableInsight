import pandas as pd
import tabula
 
# Path to your PDF file
pdf_path ='Data/Simple Tables/sample.pdf'
 
# Extract tables from the PDF
tables = tabula.read_pdf(pdf_path, pages='all', multiple_tables=True)
 
# Define the coordinates for the text area (example values; adjust as needed)
top, left, bottom, right = 0, 0, 50, 500  # These are example coordinates
 
# Extract the text from the PDF that includes the table names
text_areas = tabula.read_pdf(pdf_path, pages='all', area=[top, left, bottom, right], output_format="json", silent=True)
 
# Create an Excel writer
excel_writer = pd.ExcelWriter("output.xlsx", engine='xlsxwriter')
 
# Iterate over the extracted tables and their corresponding text areas
for i, (table, text_area) in enumerate(zip(tables, text_areas)):
    # Extract the table name from the text_area (assuming it's the first item)
    table_name = text_area[0]['data'][0][0]['text'] if text_area else f"Table_{i+1}"
    # Create a DataFrame from the table data
    df = pd.DataFrame(table)
    # Add the table name as the first row (or you could use it as the sheet name)
    df.loc[-1] = [table_name] + [''] * (df.shape[1] - 1)  # Adding table name as the first row
    df.index = df.index + 1  # Shifting index
    df = df.sort_index()  # Sorting index
    # Write the DataFrame to the Excel sheet
    df.to_excel(excel_writer, sheet_name=f"Sheet_{i+1}", index=False)
 
# Save the Excel file
excel_writer.save()