import os
from typing import List, Optional
from fastapi import FastAPI, UploadFile, File, HTTPException, Form
from fastapi.responses import FileResponse
from pathlib import Path
from fastapi.middleware.cors import CORSMiddleware
from enum import Enum
from utils.pdf_handler import (
compress_pdf, excel_to_pdf, pdf_to_word, pdf_to_excel,merge_two_pdfs
)
from utils.word_converter import (convert_docx_to_pdf)
app = FastAPI(title="PDF Tools API", version="1.0.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])



import logging

# Setup logging (optional but recommended)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Define folders (create if they don't exist)
UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


@app.delete("/clear-outputs")
async def clear_outputs(delete_all: bool = False):
    """
    DELETE /clear-outputs
    - delete_all=False (default): Delete only .pdf files
    - delete_all=True: Delete ALL files in outputs/
    
    Query params:
    - ?delete_all=true  (optional)
    
    Returns: JSON with deletion stats
    """
    if not os.path.exists(OUTPUT_DIR):
        raise HTTPException(status_code=404, detail=f"Directory '{OUTPUT_DIR}' not found. Create it first.")

    deleted_count = 0
    errors = []

    try:
        for filename in os.listdir(OUTPUT_DIR):
            file_path = os.path.join(OUTPUT_DIR, filename)
            if os.path.isfile(file_path):
                if delete_all or filename.lower().endswith('.pdf'):
                    try:
                        os.remove(file_path)
                        deleted_count += 1
                        logger.info(f"Deleted: {filename}")
                    except OSError as e:
                        errors.append(f"Failed to delete {filename}: {str(e)}")
                else:
                    logger.debug(f"Skipped non-PDF: {filename}")
            else:
                logger.debug(f"Skipped non-file: {file_path}")

        # Optional: If no files left, remove empty folder (uncomment if desired)
        # if not os.listdir(OUTPUT_DIR):
        #     os.rmdir(OUTPUT_DIR)
        #     logger.info("Removed empty outputs/ directory")

        return {
            "status": "success",
            "deleted_count": deleted_count,
            "remaining_files": len(os.listdir(OUTPUT_DIR)) if os.path.exists(OUTPUT_DIR) else 0,
            "errors": errors,
            "message": f"Deleted {deleted_count} file(s) from {OUTPUT_DIR}. Only PDFs targeted." if not delete_all else f"Deleted {deleted_count} file(s) from {OUTPUT_DIR}."
        }

    except Exception as e:
        logger.error(f"Clear outputs failed: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Operation failed: {str(e)}")




@app.on_event("startup")
async def check_libreoffice():
    try:
        result = subprocess.run(
            ["libreoffice", "--version"],
            capture_output=True,
            text=True,
            timeout=10
        )
        if result.returncode == 0:
            print(f"LibreOffice ready: {result.stdout.splitlines()[0]}")
        else:
            print("LibreOffice not working.")
    except FileNotFoundError:
        print("ERROR: libreoffice command not found. Run: sudo apt install libreoffice")
    except Exception as e:
        print(f"LibreOffice check failed: {e}")

# merging 
class PDFLibrary(str, Enum):
    # --- Merging Libraries ---
    pypdf       = "pypdf"        # Modern, recommended
    pypdf2      = "PyPDF2"
    pdfrw       = "pdfrw"
    fitz        = "fitz"         # PyMuPDF
    pypdf4      = "PyPDF4"
    pypdf3      = "PyPDF3"
    pdfplumber  = "pdfplumber"

@app.post("/merge")
async def merge_pdfs(
    file_1: UploadFile = File(..., description="First PDF to merge"),
    file_2: UploadFile = File(..., description="Second PDF to merge"),
    library: str = Form("pypdf", description="PDF merging library: pypdf, PyPDF2, pdfrw, fitz, PyPDF4, PyPDF3, pdfplumber")
):
    """
    Merge two PDFs using the specified library and return the result for download.
    
    Supported libraries:
        pypdf (default), PyPDF2, pdfrw, fitz, PyPDF4, PyPDF3, pdfplumber
    """
    try:
        await clear_outputs()  # your cleanup function
        result = merge_two_pdfs(
            file_1=file_1,
            file_2=file_2,
            upload_dir=UPLOAD_DIR,
            output_dir=OUTPUT_DIR,
            library=library  # passed here
        )
        filename = result["filename"]
        file_path = os.path.join(OUTPUT_DIR, filename)

        return FileResponse(
            path=file_path,
            media_type="application/pdf",
            filename=filename,
            headers={"message": "PDFs merged successfully!"}
        )

    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except NotImplementedError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error merging PDFs: {str(e)}")
# compression woring

class CompressLibrary(str, Enum):
# --- Compression Libraries ---
    fitz        = "fitz"
    pdfrw       = "pdfrw"
    pikepdf     = "pikepdf"      # Lossless, fast
    ghostscript = "ghostscript"  # Max compression (requires system install)
    qpdf        = "qpdf"         # Linearize + compress (requires system install)
    pdfminer_six = "pdfminer.six"      # Text-only rebuild (lossy)


@app.post("/compress")
async def compress_pdf_endpoint(
    file_1: UploadFile = File(..., description="PDF file to compress"),
    target_size: Optional[str] = Form(
        None,
        description="Target size like '500KB' or '2MB'"
    ),
    compression_percent: Optional[float] = Form(
        None,
        description="Reduce to X% of original (1–99)"
    ),
    library: CompressLibrary = Form(
        CompressLibrary.pikepdf,
        description="Choose compression backend"
    )
):

    """
    Compress a PDF using the selected library.
    """
    try:
        await clear_outputs()
        result = compress_pdf(
            file_1=file_1,
            upload_dir=UPLOAD_DIR,
            output_dir=OUTPUT_DIR,
            target_size=target_size,
            compression_percent=compression_percent,
            library=library.value  # .value gives the string
        )

        file_path = os.path.join(OUTPUT_DIR, result["filename"])
        return FileResponse(
            path=file_path,
            media_type="application/pdf",
            filename=result["filename"],
            headers={
                "X-Original-Size-KB": str(result["original_size_kb"]),
                "X-Compressed-Size-KB": str(result["compressed_size_kb"]),
                "X-Target-Accuracy": result["target_accuracy"]
            }
        )

    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except ImportError as e:
        raise HTTPException(status_code=400, detail=f"Library not available: {str(e)}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Compression failed: {str(e)}")


@app.post("/excel-to-pdf")
async def excel_to_pdf_endpoint(
    file_excel: UploadFile = File(..., description="Excel file"),
    sheet_name: Optional[str] = Form(None, description="Sheet name (e.g. 'Sales')"),
    sheet_index: Optional[int] = Form(None, description="Sheet index (0-based, e.g. 1)")
):
    await clear_outputs()
    result = excel_to_pdf(file_excel, UPLOAD_DIR, OUTPUT_DIR, sheet_name, sheet_index)
    path = os.path.join(OUTPUT_DIR, result["filename"])

    return FileResponse(
        path=path,
        media_type="application/pdf",
        filename=result["filename"],
        headers={
            "sheets-converted": str(result["sheets_converted"]),
            "sheet-names": result["sheet_names"],
            "selection-mode": result["mode"]
        }
    )

# need to above are working

@app.post("/docx-to-pdf")
async def docx_to_pdf_download(
    file_docx: UploadFile = File(...),
    page_numbers: Optional[str] = Form(None),
):
    """
    Upload .docx → Get converted PDF for download
    """
    await clear_outputs()
    pdf_path = convert_docx_to_pdf(
        file_docx=file_docx,
        upload_dir=UPLOAD_DIR,
        output_dir=OUTPUT_DIR,
        page_numbers=page_numbers,
    )

    # Return PDF for download
    return FileResponse(
        path=pdf_path,
        media_type="application/pdf",
        filename=f"converted_{file_docx.filename.replace('.docx', '.pdf')}"
    )


@app.post("/pdf-to-word")
async def pdf_to_word_endpoint(
    file_pdf: UploadFile = File(..., description="PDF file to convert")
):
    try:
        await clear_outputs()
        result = pdf_to_word(file_pdf, UPLOAD_DIR, OUTPUT_DIR)
        filename = result["filename"]
        return FileResponse(
            os.path.join(OUTPUT_DIR, filename),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename=filename,
            headers={"message": "PDF converted to Word successfully."}
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error converting PDF to Word: {str(e)}")

@app.post("/pdf-to-excel")
async def pdf_to_excel_endpoint(
    file_pdf: UploadFile = File(..., description="PDF file to extract tables from"),
    pages: Optional[str] = Form("all", description="Pages to extract from, e.g., 'all' or '1-3'")
):
    try:
        result = pdf_to_excel(file_pdf, UPLOAD_DIR, OUTPUT_DIR, pages)
        filename = result["filename"]
        return FileResponse(
            os.path.join(OUTPUT_DIR, filename),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=filename,
            headers={"message": "PDF tables extracted to Excel successfully."}
        )
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error extracting PDF to Excel: {str(e)}")

@app.get("/download/{filename:path}")
async def download_file(filename: str):
    file_path = os.path.join(OUTPUT_DIR, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found.")
    ext = Path(filename).suffix.lower()
    media_types = {
        '.pdf': 'application/pdf',
        '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    }
    media_type = media_types.get(ext, 'application/octet-stream')
    return FileResponse(file_path, media_type=media_type, filename=filename)


@app.get("/")
async def root():
    return {"message": "PDF Merge API is running. Use POST /merge with file_1 and file_2."}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)