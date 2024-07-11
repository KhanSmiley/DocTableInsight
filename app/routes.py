from fastapi import APIRouter, UploadFile, File
from app.ocr_extraction import OCRExtraction
from app.pdf_extraction import PDFExtraction

router = APIRouter()

@router.post("/upload_pdf/")
async def upload_pdf(file: UploadFile = File(...)):
    pdf_extractor = PDFExtraction()
    tables = pdf_extractor.extract_tables(file.file)
    # Store tables in database
    return {"tables": tables}

@router.post("/extract_tables/")
async def extract_tables(file: UploadFile = File(...)):
    ocr_extractor = OCRExtraction()
    tables = ocr_extractor.extract_tables(file.file)
    # Store tables in database
    return {"tables": tables}
