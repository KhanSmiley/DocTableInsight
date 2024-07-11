import os
from PyPDF2 import PdfReader
from tabula import read_pdf
import pandas as pd
from difflib import SequenceMatcher

class DocInsight:
    def __init__(self, pdf_path, ground_truth_path):
        self.pdf_path = pdf_path
        self.ground_truth_path = ground_truth_path
        self.base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        self.extracted_tables = None
        self.ground_truth = None

    def count_pages(self):
        reader = PdfReader(self.pdf_path)
        return len(reader.pages)

    def count_tables(self):
        try:
            tables = read_pdf(self.pdf_path, pages='all', multiple_tables=True)
            return len(tables)
        except Exception as e:
            print(f"Error reading tables with tabula: {e}")
            return 0

    def extract_tables(self):
        self.extracted_tables = read_pdf(self.pdf_path, pages='all', multiple_tables=True)

    def load_ground_truth(self):
        self.ground_truth = pd.read_excel(self.ground_truth_path, sheet_name=None)

    @staticmethod
    def compare_tables(extracted_table, ground_truth_table):
        matched_cells = 0
        total_extracted_cells = extracted_table.size
        total_ground_truth_cells = ground_truth_table.size

        for i in range(min(len(extracted_table), len(ground_truth_table))):
            for j in range(min(len(extracted_table.columns), len(ground_truth_table.columns))):
                if SequenceMatcher(None, str(extracted_table.iat[i, j]), str(ground_truth_table.iat[i, j])).ratio() > 0.8:
                    matched_cells += 1

        precision = matched_cells / total_extracted_cells if total_extracted_cells > 0 else 0
        recall = matched_cells / total_ground_truth_cells if total_ground_truth_cells > 0 else 0

        return precision, recall

    def process_tables(self):
        if not os.path.exists(self.pdf_path) or not os.path.exists(self.ground_truth_path):
            print("PDF file or ground truth file not found.")
            return

        num_pages = self.count_pages()
        print(f"Number of pages in the PDF: {num_pages}")

        num_tables = self.count_tables()
        print(f"Number of tables detected: {num_tables}")

        if num_tables == 0:
            print("No tables found in the PDF.")
            return

        try:
            self.extract_tables()
            self.load_ground_truth()

            output_path = f"{self.base_name}_extracted_tables.xlsx"
            
            with pd.ExcelWriter(output_path) as writer:
                for idx, table in enumerate(self.extracted_tables):
                    sheet_name = f"Table {idx + 1}"
                    table.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    if sheet_name in self.ground_truth:
                        ground_truth_table = self.ground_truth[sheet_name]
                        precision, recall = self.compare_tables(table, ground_truth_table)
                        
                        print(f"{sheet_name} Metrics:")
                        print(f"Precision: {precision:.2f}")
                        print(f"Recall: {recall:.2f}")
                    else:
                        print(f"No ground truth available for {sheet_name}")
            
            print(f"All tables saved to {output_path}")
        except Exception as e:
            print(f"Error processing tables: {e}")

def main():
    pdf_path = 'Simple Table pdf/sample.pdf'
    ground_truth_path = 'Simple Table pdf/sample_GT.xlsx'
    extractor = DocInsight(pdf_path, ground_truth_path)
    extractor.process_tables()

if __name__ == "__main__":
    main()