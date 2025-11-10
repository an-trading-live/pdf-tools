import os
import shutil
from datetime import datetime
from typing import Dict, Optional
from fastapi import UploadFile, Form
from pypdf import PdfWriter
import pikepdf
from openpyxl import load_workbook
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from docx2pdf import convert
from pdf2docx import Converter
from typing import List, Dict, Optional
import tabula
import pandas as pd
import pymupdf as fitz
import uuid
import shutil
import subprocess
import shutil
import subprocess
from datetime import datetime
from typing import Dict
from fastapi import UploadFile, HTTPException
from pathlib import Path

import asyncio
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware


UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

def get_temp_path(filename: str, prefix: str, upload_dir: str) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(upload_dir, f"{prefix}_{timestamp}_{filename}")

def get_output_path(prefix: str, output_dir: str) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(output_dir, f"{prefix}_{timestamp}.pdf")
#working  merge
def merge_two_pdfs(file_1: UploadFile, file_2: UploadFile,
                   upload_dir: str, output_dir: str) -> Dict[str, str]:
    """
    Merges exactly two PDFs.
    Returns dict with the filename of the merged PDF.
    """
    from pypdf import PdfWriter

    # Validate
    for f in (file_1, file_2):
        if not f.filename.lower().endswith('.pdf'):
            raise ValueError(f"File '{f.filename}' is not a PDF.")

    temp_paths = []
    try:
        # Save both files temporarily
        temp_path1 = get_temp_path(file_1.filename, "merge1", upload_dir)
        temp_path2 = get_temp_path(file_2.filename, "merge2", upload_dir)

        with open(temp_path1, "wb") as buf:
            shutil.copyfileobj(file_1.file, buf)
        with open(temp_path2, "wb") as buf:
            shutil.copyfileobj(file_2.file, buf)

        temp_paths = [temp_path1, temp_path2]

        # Merge
        writer = PdfWriter()
        writer.append(temp_path1)
        writer.append(temp_path2)

        output_path = get_output_path("merged", output_dir)
        with open(output_path, "wb") as f:
            writer.write(f)
        writer.close()

        return {"filename": os.path.basename(output_path)}

    finally:
        # Always delete temps
        for p in temp_paths:
            if os.path.exists(p):
                os.remove(p)

# new compress
def _calculate_target_size(
    original_bytes: int,
    target_size: Optional[str] = None,
    percent: Optional[float] = None
) -> int:
    """
    Calculate target file size in bytes.

    Priority:
    1. target_size → "500KB" or "2MB"
    2. percent → 70 (means 70% of original)
    3. Default → 50% of original

    Returns:
        int: Target size in bytes
    """
    # 1. Target Size (KB/MB)
    if target_size:
        target_size = target_size.strip().upper()
        if not target_size.endswith(("KB", "MB")):
            raise ValueError("target_size must end with 'KB' or 'MB' (e.g., '500KB', '2MB')")

        try:
            value = float(target_size[:-2])
        except ValueError:
            raise ValueError("Invalid number in target_size")

        if target_size.endswith("KB"):
            return int(value * 1024)
        elif target_size.endswith("MB"):
            return int(value * 1024 * 1024)

    # 2. Compression Percentage
    elif percent is not None:
        if not (1 <= percent <= 99):
            raise ValueError("compression_percent must be between 1 and 99")
        return int(original_bytes * (percent / 100))

    # 3. Default: 50%
    else:
        return int(original_bytes * 0.75)

def _compress_with_images(input_path: str, output_path: str, target_bytes: int):
    doc = fitz.open(input_path)
    low, high = 0.3, 1.0
    best_zoom = 1.0

    for _ in range(15):
        zoom = (low + high) / 2
        size = _render_and_measure(doc, output_path, zoom)
        if size <= target_bytes:
            best_zoom = zoom
            low = zoom
        else:
            high = zoom

    # Final render
    _render_and_measure(doc, output_path, best_zoom, final=True)
    doc.close()


def _render_and_measure(doc, temp_path, zoom, final=False):
    new_doc = fitz.open()
    for page in doc:
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_data = pix.tobytes("png")
        new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
        new_page.insert_image(page.rect, stream=img_data)
    new_doc.save(temp_path, garbage=4, deflate=True, clean=True)
    new_doc.close()
    return os.path.getsize(temp_path)

def compress_pdf(
    file_1: UploadFile,
    upload_dir: str,
    output_dir: str,
    target_size: Optional[str] = None,
    compression_percent: Optional[float] = None
) -> Dict[str, str]:
    import pikepdf
    from pikepdf import ObjectStreamMode, StreamDecodeLevel
    import pymupdf as fitz
    import os

    if not file_1.filename.lower().endswith('.pdf'):
        raise ValueError("File must be a PDF.")

    temp_input = get_temp_path(file_1.filename, "compress_input", upload_dir)
    final_output = get_output_path("compressed", output_dir)
    temp_output = final_output + ".tmp"  # intermediate

    try:
        # Save uploaded file
        with open(temp_input, "wb") as f:
            shutil.copyfileobj(file_1.file, f)

        original_size = os.path.getsize(temp_input)
        print("original_size ",original_size)
        target_bytes = _calculate_target_size(original_size, target_size, compression_percent)
        print("target_bytes ",target_bytes)
        # Try image compression
        _compress_with_images(temp_input, temp_output, target_bytes)

        # Decide final file
        if os.path.exists(temp_output) and os.path.getsize(temp_output) <= target_bytes * 1.1:
            # Use image-compressed version
            os.rename(temp_output, final_output)  # rename .tmp → final
            # print(temp_output,final_output)
            compressed_path = final_output
        else:
            # Fallback: pikepdf
            if os.path.exists(temp_output):
                os.remove(temp_output)
            with pikepdf.open(temp_input) as pdf:
                pdf.save(
                    final_output,
                    compress_streams=True,
                    stream_decode_level=StreamDecodeLevel.generalized,
                    object_stream_mode=ObjectStreamMode.preserve,
                    linearize=True
                )
            compressed_path = final_output
            # print(compressed_path)
        final_size = os.path.getsize(compressed_path)
        return {
            "filename": os.path.basename(compressed_path),
            "original_size_kb": round(original_size / 1024, 1),
            "compressed_size_kb": round(final_size / 1024, 1),
            "target_accuracy": f"{round(final_size / target_bytes * 100, 1)}%"
        }

    finally:
        # Clean only temp files
        for path in [temp_input, temp_output]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass  # ignore if already gone


#close compress



def cleanup_temp_files(temp_paths: List[str]) -> None:
    for path in temp_paths:
        if os.path.exists(path):
            os.remove(path)

def get_temp_path(filename: str, prefix: str, upload_dir: str) -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(upload_dir, f"{prefix}_{timestamp}_{filename}")

def get_output_path(prefix: str, output_dir: str, extension: str = ".pdf") -> str:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{prefix}_{timestamp}{extension}"
    return os.path.join(output_dir, filename)
# not used 
def merge_pdfs(file_1: UploadFile, file_2: UploadFile, upload_dir: str, output_dir: str) -> Dict[str, str]:
    files = [file_1, file_2]
    if len(files) < 2:
        raise ValueError("At least two PDFs required.")
    temp_paths = []
    try:
        for file in files:
            if not file.filename.lower().endswith('.pdf'):
                raise ValueError(f"File {file.filename} is not a PDF.")
            temp_path = get_temp_path(file.filename, "merge_temp", upload_dir)
            with open(temp_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            temp_paths.append(temp_path)
        
        writer = PdfWriter()
        for temp_path in temp_paths:
            writer.append(temp_path)
        
        output_path = get_output_path("merged", output_dir)
        with open(output_path, "wb") as f:
            writer.write(f)
        writer.close()
        
        return {"filename": os.path.basename(output_path)}
    
    except Exception as e:
        raise e
    
    finally:
        cleanup_temp_files(temp_paths)

def compress_pdf_old(file_1: UploadFile, upload_dir: str, output_dir: str) -> Dict[str, str]:
    if not file_1.filename.lower().endswith('.pdf'):
        raise ValueError("File is not a PDF.")
    temp_path = get_temp_path(file_1.filename, "compress_temp", upload_dir)
    try:
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file_1.file, buffer)
        
        output_path = get_output_path("compressed", output_dir)
        with pikepdf.open(temp_path) as pdf:
            pdf.save(output_path, optimize_images=True, compression_level=9)
        
        return {"filename": os.path.basename(output_path)}
    
    except Exception as e:
        raise e
    
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)
# working fing upto this functionality

def excel_to_pdf(
    file_excel: UploadFile,
    upload_dir: str,
    output_dir: str,
    sheet_name: Optional[str] = None,     # e.g. "Sales"
    sheet_index: Optional[int] = None     # e.g. 0, 1, 2
) -> Dict[str, str]:
    """
    Convert Excel to PDF with flexible sheet selection:
    - sheet_name="Sales" → by name
    - sheet_index=1 → by index (0-based)
    - both None → ALL sheets (multi-page PDF)
    """
    from openpyxl import load_workbook
    from reportlab.lib.pagesizes import letter
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet
    import os
    import shutil
    from fastapi import HTTPException

    if not file_excel.filename.lower().endswith(('.xlsx', '.xls')):
        raise ValueError("File must be Excel (.xlsx or .xls).")

    temp_path = get_temp_path(file_excel.filename, "excel_temp", upload_dir)
    output_path = get_output_path("excel_to_pdf", output_dir)

    try:
        with open(temp_path, "wb") as f:
            shutil.copyfileobj(file_excel.file, f)

        wb = load_workbook(temp_path, data_only=True)
        doc = SimpleDocTemplate(output_path, pagesize=letter, topMargin=40)
        elements = []
        styles = getSampleStyleSheet()

        # === Determine sheets to convert ===
        sheets = []

        if sheet_name and sheet_index is not None:
            raise ValueError("Use either sheet_name OR sheet_index, not both.")

        if sheet_name:
            if sheet_name not in wb.sheetnames:
                raise ValueError(f"Sheet '{sheet_name}' not found. Available: {', '.join(wb.sheetnames)}")
            sheets = [wb[sheet_name]]

        elif sheet_index is not None:
            if not (0 <= sheet_index < len(wb.worksheets)):
                raise ValueError(f"sheet_index {sheet_index} out of range. Valid: 0–{len(wb.worksheets)-1}")
            sheets = [wb.worksheets[sheet_index]]

        else:
            sheets = wb.worksheets  # ALL sheets

        # === Process Each Sheet ===
        for ws in sheets:
            title = Paragraph(f"<b>{ws.title}</b>", styles["Title"])
            elements.append(title)
            elements.append(Spacer(1, 12))

            data = []
            for row in ws.iter_rows(values_only=True):
                safe_row = ["" if c is None else str(c).strip() for c in row]
                data.append(safe_row)

            data = [row for row in data if any(cell for cell in row)]
            if not data:
                elements.append(Paragraph("No data in this sheet.", styles["Normal"]))
                elements.append(PageBreak())
                continue

            table = Table(data)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ('FONTSIZE', (0, 1), (-1, -1), 9),
            ]))
            elements.append(table)
            elements.append(PageBreak())

        if not elements:
            raise ValueError("No data found in selected sheet(s).")

        doc.build(elements)

        mode = "all_sheets" if not sheet_name and sheet_index is None else "by_name" if sheet_name else "by_index"
        return {
            "filename": os.path.basename(output_path),
            "sheets_converted": len(sheets),
            "sheet_names": ", ".join(ws.title for ws in sheets),
            "mode": mode
        }

    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error converting Excel to PDF: {str(e)}")

    finally:
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except:
                pass

# above code is working fineee
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50 MB
LIBREOFFICE_TIMEOUT = 60


def docx_to_pdfs( file_excel: UploadFile,
    upload_dir: str,
    output_dir: str, output_pdf=None):
    """
    Convert a .docx file to .pdf
    """
    if not os.path.exists(input_docx):
        print(f"Error: File {input_docx} not found!")
        return False

    if output_pdf is None:
        output_pdf = input_docx.replace(".docx", ".pdf")

    try:
        convert(input_docx, output_pdf)
        print(f"Successfully converted: {input_docx} → {output_pdf}")
        return True
    except Exception as e:
        print(f"Conversion failed: {e}")
        return False


async def word_to_pdf(file_doc: UploadFile, upload_dir: str, output_dir: str) -> Dict[str, str]:
    """
    Convert .docx → PDF using docx2pdf (pure pip package).
    Works on macOS, Windows, Linux.
    """
    original_name = Path(file_doc.filename).name

    # Only .docx supported
    if not original_name.lower().endswith('.docx'):
        raise HTTPException(400, "Only .docx files are supported with docx2pdf")

    # Paths
    input_path = os.path.join(upload_dir, f"docx_{uuid.uuid4().hex}.docx")
    final_pdf = os.path.join(output_dir, f"docx_{uuid.uuid4().hex}.pdf")
    cleanup = [input_path]

    try:
        # Save uploaded .docx
        os.makedirs(upload_dir, exist_ok=True)
        with open(input_path, "wb") as f:
            shutil.copyfileobj(file_doc.file, f)

        # Convert: docx2pdf writes directly to final_pdf
        os.makedirs(output_dir, exist_ok=True)
        docx_convert(input_path, final_pdf)

        if not os.path.exists(final_pdf):
            raise HTTPException(500, "PDF was not generated by docx2pdf")

        return {
            "filename": Path(final_pdf).name,
            "original_format": "docx",
            "method": "docx2pdf"
        }

    except Exception as e:
        raise HTTPException(500, f"Conversion failed: {str(e)}")
    finally:
        # Clean up input file
        for p in cleanup:
            if os.path.exists(p):
                try:
                    os.remove(p)
                except:
                    pass


def pdf_to_word(file_pdf: UploadFile, upload_dir: str, output_dir: str) -> Dict[str, str]:
    if not file_pdf.filename.lower().endswith('.pdf'):
        raise ValueError("File must be .pdf.")
    temp_path = get_temp_path(file_pdf.filename, "pdf_to_word_temp", upload_dir)
    try:
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file_pdf.file, buffer)
        
        output_path = get_output_path("pdf_to_word", output_dir, ".docx")
        cv = Converter(temp_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()
        
        return {"filename": os.path.basename(output_path)}
    
    except Exception as e:
        raise e
    
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

def pdf_to_excel(file_pdf: UploadFile, upload_dir: str, output_dir: str, pages: Optional[str] = "all") -> Dict[str, str]:
    if not file_pdf.filename.lower().endswith('.pdf'):
        raise ValueError("File must be .pdf.")
    temp_path = get_temp_path(file_pdf.filename, "pdf_to_excel_temp", upload_dir)
    try:
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file_pdf.file, buffer)
        
        dfs = tabula.read_pdf(temp_path, pages=pages, multiple_tables=True)
        if not dfs:
            raise ValueError("No tables found in PDF.")
        output_path = get_output_path("pdf_to_excel", output_dir, ".xlsx")
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            for i, df in enumerate(dfs):
                df.to_excel(writer, sheet_name=f"Table_{i+1}", index=False)
        
        return {"filename": os.path.basename(output_path)}
    
    except Exception as e:
        raise e
    
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)
# new pdf functions
