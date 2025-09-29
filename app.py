#!/usr/bin/env python3
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
import uuid
import asyncio
from pathlib import Path
import aiofiles
import logging

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="DOCX to PDF Converter API",
    description="Convert DOCX files to PDF format - n8n compatible",
    version="1.0.0"
)

# Enable CORS for n8n integration
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("outputs")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


class ConversionMethods:
    @staticmethod
    async def convert_with_reportlab(input_path: str, output_path: str) -> bool:
        """Convert using ReportLab + python-docx"""
        try:
            loop = asyncio.get_event_loop()

            def _convert():
                doc = Document(input_path)
                pdf_doc = SimpleDocTemplate(output_path, pagesize=letter)
                styles = getSampleStyleSheet()
                story = []

                for paragraph in doc.paragraphs:
                    if paragraph.text.strip():
                        style = styles['Normal']
                        if paragraph.style.name.startswith('Heading'):
                            style = styles['Heading1']
                        p = Paragraph(paragraph.text, style)
                        story.append(p)
                        story.append(Spacer(1, 0.1 * inch))

                pdf_doc.build(story)
                return True

            await loop.run_in_executor(None, _convert)
            return True
        except Exception as e:
            logger.error(f"Conversion failed: {e}")
            return False


async def convert_file(input_path: str, output_path: str) -> bool:
    """Convert DOCX to PDF"""
    return await ConversionMethods.convert_with_reportlab(input_path, output_path)


@app.post("/convert")
async def convert_docx_to_pdf(
    file: UploadFile = File(...),
    return_file: bool = Form(default=True)
):
    """Convert DOCX to PDF (sync). Works directly with n8n HTTP Request node."""

    if not file.filename.lower().endswith('.docx'):
        raise HTTPException(status_code=400, detail="File must be a .docx document")

    job_id = str(uuid.uuid4())
    input_filename = f"{job_id}_{file.filename}"
    output_filename = f"{job_id}_{Path(file.filename).stem}.pdf"
    input_path = UPLOAD_DIR / input_filename
    output_path = OUTPUT_DIR / output_filename

    try:
        # Save input file
        async with aiofiles.open(input_path, 'wb') as f:
            content = await file.read()
            await f.write(content)

        logger.info(f"Converting {file.filename}...")

        success = await convert_file(str(input_path), str(output_path))

        if not success or not output_path.exists():
            raise HTTPException(status_code=500, detail="Conversion failed")

        if return_file:
            return FileResponse(
                path=output_path,
                filename=f"{Path(file.filename).stem}.pdf",
                media_type='application/pdf'
            )
        else:
            return {
                "success": True,
                "output_filename": output_filename,
                "download_url": f"/download/{output_filename}"
            }

    finally:
        if input_path.exists():
            input_path.unlink()


@app.get("/download/{filename}")
async def download_file(filename: str):
    file_path = OUTPUT_DIR / filename
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found")

    return FileResponse(
        path=file_path,
        filename=filename,
        media_type='application/pdf'
    )
