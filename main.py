from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pdf2docx import Converter
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image, ImageDraw, ImageFont
import io, os, zipfile, json, base64, qrcode, tempfile, shutil
from uuid import uuid4
from pdf2image import convert_from_path
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from openpyxl import load_workbook
import fitz  # PyMuPDF
import aspose.slides as slides

# ---------------- App Setup ----------------
app = FastAPI(title="ToolForge Backend API 🚀")

UPLOAD_DIR = "temp"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ---------------- CORS ----------------
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://tool-forge-frontend-bu5k.vercel.app",
        "http://localhost:3000"
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------- Utility Functions ----------------
async def save_upload(file: UploadFile) -> str:
    """Save uploaded file to temp directory with unique name"""
    ext = os.path.splitext(file.filename)[1]
    unique_name = f"{uuid4().hex}{ext}"
    file_path = os.path.join(UPLOAD_DIR, unique_name)
    content = await file.read()
    with open(file_path, "wb") as f:
        f.write(content)
    return file_path

def cleanup_file(file_path: str):
    """Safely remove file if it exists"""
    try:
        if os.path.exists(file_path):
            os.remove(file_path)
    except Exception:
        pass

# ---------------- Root ----------------
@app.get("/", include_in_schema=False)
async def root():
    return {
        "message": "Welcome to ToolForge Backend 🚀",
        "docs": "/docs",
        "redoc": "/redoc"
    }

# ---------- CONVERSION ENDPOINTS ---------- #

# ---------------- PDF → DOCX ----------------
@app.post("/convert/pdf-to-docx")
async def convert_pdf_to_docx(file: UploadFile = File(...)):
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="File must be a PDF")
    
    input_path = None
    output_path = None
    
    try:
        input_path = await save_upload(file)
        output_path = input_path.replace(".pdf", ".docx")
        
        # Convert PDF to DOCX
        cv = Converter(input_path)
        cv.convert(output_path, start=0, end=None)
        cv.close()
        
        if not os.path.exists(output_path):
            raise Exception("Conversion failed: DOCX file not created")
            
        return FileResponse(
            output_path, 
            filename=f"{os.path.splitext(file.filename)[0]}.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF to DOCX conversion failed: {str(e)}")
    finally:
        cleanup_file(input_path)
        # Don't cleanup output_path here as it's being served

# ---------------- DOCX → PDF ----------------
@app.post("/convert/docx-to-pdf")
async def convert_docx_to_pdf(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(('.docx', '.doc')):
        raise HTTPException(status_code=400, detail="File must be a DOCX/DOC")
    
    input_path = None
    output_path = None
    
    try:
        input_path = await save_upload(file)
        output_path = input_path.replace(os.path.splitext(input_path)[1], ".pdf")
        
        # Read DOCX and convert to PDF
        doc = Document(input_path)
        c = canvas.Canvas(output_path, pagesize=A4)
        width, height = A4
        y = height - 50
        
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text:
                continue
                
            if y < 50:  # Start new page if needed
                c.showPage()
                y = height - 50
                
            # Handle long text by wrapping
            max_width = width - 100
            if c.stringWidth(text) > max_width:
                words = text.split()
                line = ""
                for word in words:
                    test_line = line + " " + word if line else word
                    if c.stringWidth(test_line) > max_width:
                        if line:
                            c.drawString(50, y, line)
                            y -= 15
                            if y < 50:
                                c.showPage()
                                y = height - 50
                        line = word
                    else:
                        line = test_line
                if line:
                    c.drawString(50, y, line)
                    y -= 15
            else:
                c.drawString(50, y, text)
                y -= 15
        
        c.save()
        
        if not os.path.exists(output_path):
            raise Exception("PDF not created")
            
        return FileResponse(
            output_path,
            filename=f"{os.path.splitext(file.filename)[0]}.pdf",
            media_type="application/pdf"
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"DOCX to PDF conversion failed: {str(e)}")
    finally:
        cleanup_file(input_path)

# ---------------- PDF → Image ----------------
@app.post("/convert/pdf-to-image")
async def convert_pdf_to_image(file: UploadFile = File(...)):
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="File must be a PDF")
    
    temp_dir = None
    
    try:
        temp_dir = tempfile.mkdtemp()
        input_path = os.path.join(temp_dir, file.filename)
        
        with open(input_path, "wb") as f:
            f.write(await file.read())
        
        # Convert PDF to images
        images = convert_from_path(input_path, dpi=200)
        output_files = []
        
        for i, img in enumerate(images):
            img = img.convert("RGB")
            out_path = os.path.join(temp_dir, f"page_{i+1}.jpg")
            img.save(out_path, "JPEG", quality=95)
            output_files.append(out_path)
        
        # Return single image or ZIP for multiple
        if len(output_files) == 1:
            return FileResponse(
                output_files[0],
                filename=f"{os.path.splitext(file.filename)[0]}.jpg",
                media_type="image/jpeg"
            )
        else:
            zip_path = os.path.join(temp_dir, "pdf_images.zip")
            with zipfile.ZipFile(zip_path, "w") as zipf:
                for i, img_file in enumerate(output_files):
                    zipf.write(img_file, f"page_{i+1}.jpg")
            
            return FileResponse(
                zip_path,
                filename=f"{os.path.splitext(file.filename)[0]}_images.zip",
                media_type="application/zip"
            )
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF to Image conversion failed: {str(e)}")
    finally:
        if temp_dir and os.path.exists(temp_dir):
            shutil.rmtree(temp_dir, ignore_errors=True)

# ---------------- Image → PDF ----------------
@app.post("/convert/image-to-pdf")
async def convert_image_to_pdf(file: UploadFile = File(...)):
    if not file.content_type.startswith('image/'):
        raise HTTPException(status_code=400, detail="File must be an image")
    
    input_path = None
    output_path = None
    
    try:
        input_path = await save_upload(file)
        output_path = os.path.splitext(input_path)[0] + ".pdf"
        
        # Convert image to PDF
        img = Image.open(input_path)
        if img.mode != 'RGB':
            img = img.convert('RGB')
        
        img.save(output_path, "PDF", resolution=100.0)
        
        if not os.path.exists(output_path):
            raise Exception("PDF not created")
            
        return FileResponse(
            output_path,
            filename=f"{os.path.splitext(file.filename)[0]}.pdf",
            media_type="application/pdf"
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Image to PDF conversion failed: {str(e)}")
    finally:
        cleanup_file(input_path)

# ---------------- PPT/PPTX → PDF ----------------
@app.post("/convert/ppt-to-pdf")
async def convert_ppt_to_pdf(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(('.ppt', '.pptx')):
        raise HTTPException(status_code=400, detail="File must be a PPT/PPTX")
    
    input_path = None
    output_path = None
    
    try:
        input_path = await save_upload(file)
        output_path = os.path.splitext(input_path)[0] + ".pdf"
        
        # Load and convert presentation
        pres = slides.Presentation(input_path)
        pres.save(output_path, slides.export.SaveFormat.PDF)
        pres.dispose()
        
        if not os.path.exists(output_path):
            raise Exception("PDF not created")
            
        return FileResponse(
            output_path,
            filename=f"{os.path.splitext(file.filename)[0]}.pdf",
            media_type="application/pdf"
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PPT to PDF conversion failed: {str(e)}")
    finally:
        cleanup_file(input_path)

# ---------------- Excel → PDF ----------------
@app.post("/convert/excel-to-pdf")
async def convert_excel_to_pdf(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="File must be an Excel file")
    
    input_path = None
    output_path = None
    
    try:
        input_path = await save_upload(file)
        output_path = os.path.splitext(input_path)[0] + ".pdf"
        
        wb = load_workbook(input_path)
        c = canvas.Canvas(output_path, pagesize=A4)
        width, height = A4
        margin = 50
        
        for sheet_idx, sheet in enumerate(wb.worksheets):
            if sheet_idx > 0:  # New page for each sheet except first
                c.showPage()
            
            y = height - margin
            
            # Sheet title
            c.setFont("Helvetica-Bold", 16)
            c.drawString(margin, y, f"Sheet: {sheet.title}")
            y -= 30
            
            c.setFont("Helvetica", 10)
            
            # Process rows
            for row_idx, row in enumerate(sheet.iter_rows(values_only=True)):
                if y < margin + 20:  # Need new page
                    c.showPage()
                    y = height - margin
                
                # Convert row to string
                row_values = []
                for cell in row:
                    if cell is not None:
                        row_values.append(str(cell)[:50])  # Limit cell content
                    else:
                        row_values.append("")
                
                row_text = " | ".join(row_values)
                
                # Wrap long lines
                max_width = width - 2 * margin
                if c.stringWidth(row_text) > max_width:
                    row_text = row_text[:100] + "..."  # Truncate very long rows
                
                c.drawString(margin, y, row_text)
                y -= 12
                
                # Stop after reasonable number of rows to avoid huge PDFs
                if row_idx > 200:
                    c.drawString(margin, y, f"... (showing first 200 rows)")
                    break
        
        c.save()
        
        if not os.path.exists(output_path):
            raise Exception("PDF not created")
            
        return FileResponse(
            output_path,
            filename=f"{os.path.splitext(file.filename)[0]}.pdf",
            media_type="application/pdf"
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Excel to PDF conversion failed: {str(e)}")
    finally:
        cleanup_file(input_path)

# ---------- OTHER ENDPOINTS ---------- #

# Merge PDFs
@app.post("/merge-pdf/")
async def merge_pdf(files: list[UploadFile] = File(...)):
    if len(files) < 2:
        raise HTTPException(status_code=400, detail="Need at least 2 PDF files to merge")
    
    temp_files = []
    try:
        merger = PdfWriter()
        
        for file in files:
            if not file.filename.lower().endswith('.pdf'):
                raise HTTPException(status_code=400, detail="All files must be PDFs")
            
            file_path = await save_upload(file)
            temp_files.append(file_path)
            
            reader = PdfReader(file_path)
            for page in reader.pages:
                merger.add_page(page)
        
        output_file = os.path.join(UPLOAD_DIR, f"merged_{uuid4().hex}.pdf")
        with open(output_file, "wb") as f:
            merger.write(f)
        merger.close()
        
        return FileResponse(
            output_file,
            filename="merged.pdf",
            media_type="application/pdf"
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF merge failed: {str(e)}")
    finally:
        for temp_file in temp_files:
            cleanup_file(temp_file)

# Compress PDF
@app.post("/compress-pdf/")
async def compress_pdf(file: UploadFile = File(...), level: str = Form("medium")):
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="File must be a PDF")
    
    try:
        pdf_bytes = await file.read()
        pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        # Compression settings
        compress_options = {
            "high": {"garbage": 4, "deflate": True, "clean": True, "linear": True},
            "medium": {"garbage": 3, "deflate": True, "clean": True},
            "low": {"garbage": 2, "deflate": True}
        }
        
        options = compress_options.get(level.lower(), compress_options["medium"])
        
        compressed_pdf = io.BytesIO()
        pdf_doc.save(compressed_pdf, **options)
        pdf_doc.close()
        compressed_pdf.seek(0)
        
        return StreamingResponse(
            compressed_pdf,
            media_type="application/pdf",
            headers={"Content-Disposition": f"attachment; filename=compressed_{file.filename}"}
        )
        
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF compression failed: {str(e)}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
