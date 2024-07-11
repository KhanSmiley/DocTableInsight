import streamlit as st
import os
from pdf_extraction import DocTableInsightProcessor

def main():
    st.title("DocTableInsight")
    
    # Information about the app
    st.markdown("""
    **DocTableInsight** is a powerful tool designed to extract tables from PDF documents. 
    Simply upload your PDF file, and our tool will automatically detect and extract tables,
    saving the results in an Excel file that you can download. This application helps you 
    streamline the process of extracting structured data from complex documents.
    """)

    # Display an image
    st.image('projectimg.png', use_column_width=True)
    
    # File uploader for PDF
    uploaded_file = st.file_uploader("Upload a PDF file", type="pdf")
    
    if uploaded_file is not None:
        # Ensure the temp directory exists
        temp_dir = "temp"
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
        
        # Save the uploaded PDF to the temp directory
        pdf_path = os.path.join(temp_dir, uploaded_file.name)
        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Process the PDF file
        processor = DocTableInsightProcessor(pdf_path)
        num_tables, excel_filename = processor.process()
        
        # Display the results
        st.write(f"Number of pages: {processor.get_num_pages()}")
        st.write(f"Number of tables detected and extracted: {num_tables}")
        
        table_names = processor.get_table_names()
        for i, name in enumerate(table_names, 1):
            st.write(f"Table {i} Name: {name}")
        
        # Provide download link for the Excel file
        with open(excel_filename, "rb") as f:
            st.download_button(
                label="Download Extracted Tables",
                data=f,
                file_name=os.path.basename(excel_filename),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
