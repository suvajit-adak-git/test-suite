from fastapi import APIRouter, UploadFile, File, HTTPException, Form
from fastapi.responses import FileResponse, JSONResponse
from app.services.excel_processor import ExcelHyperlinkProcessor
import shutil
import tempfile
import os
from pathlib import Path

router = APIRouter()

@router.post("/extract-hyperlinks/")
async def extract_hyperlinks(file: UploadFile = File(...)):
    """
    Extract hyperlinks from uploaded Excel file.
    Returns JSON with all hyperlink details.
    """
    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix=Path(file.filename).suffix) as tmp_file:
        shutil.copyfileobj(file.file, tmp_file)
        tmp_path = tmp_file.name
    
    try:
        processor = ExcelHyperlinkProcessor(tmp_path)
        result = processor.extract_hyperlinks()
        processor.close()
        
        return JSONResponse(content=result)
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Cleanup temporary file
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


@router.post("/update-build/")
async def update_build(
    file: UploadFile = File(...),
    new_build: str = Form(...)
):
    """
    Update hyperlinks with new build number and return the updated file.
    """
    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix=Path(file.filename).suffix) as tmp_file:
        shutil.copyfileobj(file.file, tmp_file)
        tmp_path = tmp_file.name
    
    output_path = None
    
    try:
        processor = ExcelHyperlinkProcessor(tmp_path)
        
        # Generate output file path
        output_path = tmp_path.replace(Path(tmp_path).suffix, f"_build_{new_build}{Path(tmp_path).suffix}")
        
        result = processor.update_build_numbers(new_build, output_path)
        processor.close()
        
        if result['status'] == 'success':
            # Return the updated file
            return FileResponse(
                path=output_path,
                filename=f"{Path(file.filename).stem}_build_{new_build}{Path(file.filename).suffix}",
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            raise HTTPException(status_code=500, detail=result['message'])
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Cleanup temporary files
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        if output_path and os.path.exists(output_path):
            # Note: FileResponse handles cleanup, but we do it here for error cases
            pass


@router.post("/extract-and-update/")
async def extract_and_update(
    file: UploadFile = File(...),
    new_build: str = Form(...)
):
    """
    Combined endpoint: Extract hyperlinks info and update build number.
    Returns JSON with extraction details and download link for updated file.
    """
    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix=Path(file.filename).suffix) as tmp_file:
        shutil.copyfileobj(file.file, tmp_file)
        tmp_path = tmp_file.name
    
    output_path = None
    
    try:
        processor = ExcelHyperlinkProcessor(tmp_path)
        
        # Extract hyperlinks
        extraction_result = processor.extract_hyperlinks()
        
        # Generate output file path
        output_path = tmp_path.replace(Path(tmp_path).suffix, f"_build_{new_build}{Path(tmp_path).suffix}")
        
        # Update build numbers
        update_result = processor.update_build_numbers(new_build, output_path)
        processor.close()
        
        combined_result = {
            'extraction': extraction_result,
            'update': update_result
        }
        
        return JSONResponse(content=combined_result)
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
    
    finally:
        # Cleanup temporary files
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
