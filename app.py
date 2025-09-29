#!/usr/bin/env python3
from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import uuid
import asyncio
from pathlib import Path
import aiofiles
import logging
import subprocess
import shutil

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(
    title="DOCX to PDF Converter API",
    description="Convert DOCX files to PDF format with full formatting preservation",
    version="2.0.0"
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


async def convert_with_libreoffice(input_path: Path, output_dir: Path) -> Path:
    """
    Convert DOCX to PDF using LibreOffice (preserves all formatting)
    """
    try:
        loop = asyncio.get_event_loop()
        
        def _convert():
            # Check if LibreOffice is installed
            if not shutil.which('libreoffice'):
                raise Exception("LibreOffice is not installed")
            
            # Run LibreOffice conversion
            cmd = [
                'libreoffice',
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', str(output_dir),
                str(input_path)
            ]
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=120  # 2 minute timeout
            )
            
            if result.returncode != 0:
                logger.error(f"LibreOffice error: {result.stderr}")
                raise Exception(f"LibreOffice conversion failed: {result.stderr}")
            
            # LibreOffice creates PDF with same name as input
            output_filename = input_path.stem + '.pdf'
            output_path = output_dir / output_filename
            
            if not output_path.exists():
                raise Exception("PDF was not created")
            
            return output_path
        
        output_path = await loop.run_in_executor(None, _convert)
        return output_path
        
    except subprocess.TimeoutExpired:
        raise Exception("Conversion timeout - file may be too large or complex")
    except Exception as e:
        logger.error(f"Conversion failed: {e}")
        raise Exception(f"Conversion error: {str(e)}")


@app.post("/convert")
async def convert_docx_to_pdf(
    file: UploadFile = File(...),
    return_file: bool = Form(default=True)
):
    """
    Convert DOCX to PDF with full formatting preservation.
    Works directly with n8n HTTP Request node.
    """
    if not file.filename.lower().endswith('.docx'):
        raise HTTPException(status_code=400, detail="File must be a .docx document")

    job_id = str(uuid.uuid4())
    input_filename = f"{job_id}_{file.filename}"
    input_path = UPLOAD_DIR / input_filename

    try:
        # Save uploaded DOCX
        async with aiofiles.open(input_path, 'wb') as f:
            content = await file.read()
            await f.write(content)

        logger.info(f"Converting {file.filename} using LibreOffice...")

        # Convert using LibreOffice
        output_path = await convert_with_libreoffice(input_path, OUTPUT_DIR)
        
        if not output_path.exists():
            raise HTTPException(status_code=500, detail="Conversion failed - PDF not created")

        logger.info(f"Conversion successful: {output_path.name}")

        if return_file:
            # Return PDF file directly
            response = FileResponse(
                path=output_path,
                filename=f"{Path(file.filename).stem}.pdf",
                media_type='application/pdf'
            )

            # Schedule cleanup after response is sent
            async def cleanup_files():
                await asyncio.sleep(2)
                if output_path.exists():
                    output_path.unlink()
                    logger.info(f"Cleaned up: {output_path.name}")
            
            asyncio.create_task(cleanup_files())

            return response
        else:
            # Return JSON with download URL
            return {
                "success": True,
                "output_filename": output_path.name,
                "download_url": f"/download/{output_path.name}"
            }

    except Exception as e:
        logger.error(f"Error processing file: {e}")
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Always cleanup uploaded DOCX
        if input_path.exists():
            input_path.unlink()
            logger.info(f"Cleaned up input: {input_filename}")


@app.get("/download/{filename}")
async def download_file(filename: str):
    """Download generated PDF (cleanup after sending)"""
    file_path = OUTPUT_DIR / filename
    
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found or already downloaded")

    response = FileResponse(
        path=file_path,
        filename=filename,
        media_type='application/pdf'
    )

    # Schedule cleanup after download
    async def cleanup_pdf():
        await asyncio.sleep(2)
        if file_path.exists():
            file_path.unlink()
            logger.info(f"Cleaned up downloaded file: {filename}")
    
    asyncio.create_task(cleanup_pdf())

    return response


@app.get("/health")
async def health_check():
    """Health check endpoint"""
    # Check if LibreOffice is available
    libreoffice_available = shutil.which('libreoffice') is not None
    
    return {
        "status": "healthy",
        "libreoffice_installed": libreoffice_available,
        "version": "2.0.0"
    }


@app.get("/", include_in_schema=False)
@app.head("/", include_in_schema=False)
async def root():
    """Root endpoint for Render health checks"""
    return JSONResponse({
        "message": "DOCX to PDF Converter API is running!",
        "docs": "/docs",
        "health": "/health"
    })