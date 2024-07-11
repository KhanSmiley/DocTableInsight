import os
import PyPDF2
import tabula
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd

class DocTableExtractor:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        self.num_pages = self._get_num_pages()
        self.tables = []
        self.tables_with_pages = []
        self.text_by_page = self._extract_text()

    def _get_num_pages(self):
        with open(self.pdf_path, 'rb') as f:
            pdf_reader = PyPDF2.PdfReader(f)
            return len(pdf_reader.pages)

    def _extract_text(self):
        with open(self.pdf_path, 'rb') as f:
            pdf_reader = PyPDF2.PdfReader(f)
            text_by_page = []
            for page in pdf_reader.pages:
                text = page.extract_text()
                # Remove footer from each page
                text = self._remove_footer(text)
                text_by_page.append(text)
        return text_by_page

    def _remove_footer(self, text):
        # Split text into lines
        lines = text.split('\n')
        
        # Remove the last line if it's identified as a footer
        if lines and lines[-1].strip():
            lines = lines[:-1]
        
        # Join the remaining lines to form the updated text
        updated_text = '\n'.join(lines)
        return updated_text

    def extract_tables(self):
        for page_num in range(1, self.num_pages + 1):
            tables = tabula.read_pdf(self.pdf_path, pages=page_num, multiple_tables=True, stream=True)
            for table in tables:
                #print(f"Page {page_num}, Table:\n{table}\n")  # Debug: Print table content
                self.tables.append(table)
                self.tables_with_pages.append((table, page_num))
        for i, (table, page_num) in enumerate(self.tables_with_pages):
            table.name = self._extract_table_name(page_num, table)
        return len(self.tables)

    def _extract_table_name(self, page_num, table):
        # Extract text above the table on the same page
        text_above_table = self._get_text_above_table(page_num, table)

        # Extract the first 10 words from the text
        table_name = self._generate_table_name(text_above_table)

        return table_name

    def _get_text_above_table(self, page_num, table):
        page_text = self.text_by_page[page_num - 1]  # Page number in PdfReader starts from 0
        table_start_position = page_text.find(str(table.iloc[0, 0]))  # Find start position of the table
        text_before_table = page_text[:table_start_position]
        paragraphs = text_before_table.split('\n\n')
        relevant_paragraph = paragraphs[-1] if paragraphs else text_before_table
        return relevant_paragraph

    def _generate_table_name(self, text):
        tokens = text.split()
        table_name = ' '.join(tokens[:10])
        return table_name

class DocTableExporter:
    def __init__(self, output_folder='DocTableInsight_extracted_tables'):
        self.output_folder = output_folder
        os.makedirs(self.output_folder, exist_ok=True)

    def sanitize_sheet_title(self, title):
        # Remove or replace invalid characters
        invalid_chars = ['\\', '/', '?', '*', '[', ']', ':']
        for char in invalid_chars:
            title = title.replace(char, '')
        return title[:10]  # Limit to 10 characters

    def export_to_excel(self, tables, base_name, processor):
        wb = Workbook()
        
        # Create Information sheet
        info_ws = wb.active
        info_ws.title = "Information"
        info_ws['A1'] = f"Number of pages: {processor.get_num_pages()}"
        info_ws['A2'] = f"Number of tables detected and extracted: {len(tables)}"
        info_ws['A3'] = "List of tables:"
        
        # Populate table information
        for i, table in enumerate(tables, 1):
            info_ws[f'A{i+3}'] = f"Table {i}:"
            info_ws[f'B{i+3}'] = f"  Name: {table.name}"
            info_ws[f'C{i+3}'] = f"  Shape: {table.shape}"
        
        # Create Table sheets
        for i, table in enumerate(tables, 1):
            ws = wb.create_sheet(title=f"Table {i}")
            for r in dataframe_to_rows(table, index=False, header=True):  # Ensure header is included
                cleaned_row = [self.clean_cell(cell) for cell in r]
                #print(f"Appending row to sheet: {cleaned_row}")  # Debug: Print each row before appending
                ws.append(cleaned_row)
        
        excel_filename = os.path.join(self.output_folder, f'{base_name}.xlsx')
        wb.save(excel_filename)
        return excel_filename

    def clean_cell(self, cell):
        if isinstance(cell, str):
            if 'unnamed' in cell.lower():  # Check if cell content contains 'unnamed'
                return ''  # Replace with empty string
            else:
                return cell  # Return original cell value if not 'unnamed'
        elif pd.isna(cell):
            return ''  # Replace NaN with empty string
        else:
            return cell  # Return original cell value

class DocTableInsightProcessor:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.extractor = DocTableExtractor(pdf_path)
        self.exporter = DocTableExporter()

    def process(self):
        num_tables = self.extractor.extract_tables()
        excel_filename = self.exporter.export_to_excel(self.extractor.tables, self.extractor.base_name, self)
        return num_tables, excel_filename

    def get_table_names(self):
        return [table.name for table in self.extractor.tables]

    def get_num_pages(self):
        return self.extractor.num_pages

def main():
    pdf_path = 'Data/Simple Tables/sample.pdf'
    processor = DocTableInsightProcessor(pdf_path)
    num_tables, excel_filename = processor.process()

    print(f"Number of pages: {processor.get_num_pages()}")
    print(f"Number of tables detected and extracted: {num_tables}")
    print("List of tables:")
    for i, table in enumerate(processor.extractor.tables, 1):
        print(f"Table {i}:")
        print(f"  Name: {table.name}")
        print(f"  Shape: {table.shape}")
        print()
    print(f"Extracted tables saved to: {excel_filename}")

if __name__ == "__main__":
    main()
