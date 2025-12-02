from typing import Dict, Any, Optional
from pathlib import Path
import os

from fastapi import APIRouter, UploadFile, File, HTTPException, Body
from app.core.config import UPLOAD_DIR
from app.schemas.models import LocalPathsRequest
from app.services.extractor import (
    save_upload_file,
    extract_from_csv_file,
    extract_from_excel_file,
    extract_hyperlinks_with_versions_from_path
)
from app.services.comparator import compare_data

router = APIRouter()

@router.post("/upload-excel")
async def upload_excel(file: UploadFile = File(...)):
    fname = file.filename or ""
    fname_lower = fname.lower()
    if not fname_lower.endswith((".xls", ".xlsx", ".csv")):
        raise HTTPException(status_code=400, detail="Please upload an .xls, .xlsx or .csv file")

    dest = UPLOAD_DIR / fname
    counter = 0
    base = dest.stem
    ext = dest.suffix
    while dest.exists():
        counter += 1
        dest = UPLOAD_DIR / f"{base}_{counter}{ext}"

    try:
        save_upload_file(file, dest)
        if dest.suffix.lower() == ".csv":
            data = extract_from_csv_file(dest)
        else:
            data = extract_from_excel_file(dest)
    finally:
        try:
            dest.unlink()
        except Exception:
            pass

    return {"status": "ok", "data": data}

@router.post("/upload-review-checklist")
async def upload_review_checklist(file: UploadFile = File(...), sheet_name: Optional[str] = "Test Scenario Remarks"):
    fname = file.filename or ""
    if not fname.lower().endswith((".xls", ".xlsx")):
        raise HTTPException(status_code=400, detail="Please upload an .xls or .xlsx file for review checklist")

    dest = UPLOAD_DIR / fname
    counter = 0
    base = dest.stem
    ext = dest.suffix
    while dest.exists():
        counter += 1
        dest = UPLOAD_DIR / f"{base}_{counter}{ext}"

    try:
        save_upload_file(file, dest)
        results = extract_hyperlinks_with_versions_from_path(str(dest), sheet_name=sheet_name)
    finally:
        try:
            dest.unlink()
        except Exception:
            pass

    return {"status": "ok", "data": results, "count": len(results)}

@router.post("/upload-both")
async def upload_both(svn_file: UploadFile = File(...), checklist_file: UploadFile = File(...), sheet_name: Optional[str] = "Test Scenario Remarks"):
    svn_name = svn_file.filename or "svn_input"
    svn_dest = UPLOAD_DIR / svn_name
    i = 0
    while svn_dest.exists():
        i += 1
        svn_dest = UPLOAD_DIR / f"{Path(svn_name).stem}_{i}{svn_dest.suffix}"

    check_name = checklist_file.filename or "checklist_input.xlsx"
    check_dest = UPLOAD_DIR / check_name
    j = 0
    while check_dest.exists():
        j += 1
        check_dest = UPLOAD_DIR / f"{Path(check_name).stem}_{j}{check_dest.suffix}"

    try:
        save_upload_file(svn_file, svn_dest)
        save_upload_file(checklist_file, check_dest)

        if svn_dest.suffix.lower() == ".csv":
            svn_data = extract_from_csv_file(svn_dest)
        else:
            svn_data = extract_from_excel_file(svn_dest)

        checklist_data = extract_hyperlinks_with_versions_from_path(str(check_dest), sheet_name=sheet_name)
    finally:
        try:
            svn_dest.unlink()
        except Exception:
            pass
        try:
            check_dest.unlink()
        except Exception:
            pass

    return {"status": "ok", "svn": svn_data, "checklist": {"filename": check_name, "data": checklist_data, "count": len(checklist_data)}}

@router.post("/process-local-paths")
async def process_local_paths(req: LocalPathsRequest):
    if not req.svn_path or not os.path.isabs(req.svn_path):
        raise HTTPException(status_code=400, detail="Provide an absolute svn_path on the server")
    if not req.checklist_path or not os.path.isabs(req.checklist_path):
        raise HTTPException(status_code=400, detail="Provide an absolute checklist_path on the server")

    if not Path(req.svn_path).exists():
        raise HTTPException(status_code=404, detail=f"svn_path not found: {req.svn_path}")
    if not Path(req.checklist_path).exists():
        raise HTTPException(status_code=404, detail=f"checklist_path not found: {req.checklist_path}")

    svn_path_obj = Path(req.svn_path)
    if svn_path_obj.suffix.lower() == ".csv":
        svn_data = extract_from_csv_file(svn_path_obj)
    else:
        svn_data = extract_from_excel_file(svn_path_obj)

    checklist_data = extract_hyperlinks_with_versions_from_path(req.checklist_path, sheet_name=req.sheet_name)

    return {"status": "ok", "svn": svn_data, "checklist": {"filename": Path(req.checklist_path).name, "data": checklist_data, "count": len(checklist_data)}}

@router.post("/compare-both")
async def compare_both(payload: Dict[str, Any] = Body(...)):
    """
    Compare svn blob and checklist blob (or use server-local paths).
    Enhancements:
      - filename normalization + strip extension
      - version coercion to integer when possible
      - fuzzy matching (difflib) with threshold (default 0.85)
    Payload options:
      - { "svn": <svn_blob>, "checklist": <checklist_blob>, "fuzzy_threshold": 0.85 }
      - or provide { "svn_path": "/abs/path/to/svn_report.csv", "checklist_path": "/abs/path/to/checklist.xlsx", "sheet_name": "...", "fuzzy_threshold": 0.85 }
    """
    fuzzy_threshold = float(payload.get("fuzzy_threshold", 0.85))

    svn_blob = payload.get("svn")
    checklist_blob = payload.get("checklist")

    # If server-local paths provided, process them first
    if not svn_blob and payload.get("svn_path"):
        svn_path = payload["svn_path"]
        # accept file:// URIs by stripping
        if svn_path.startswith("file://"):
            svn_path = svn_path[len("file://"):]
        if not os.path.isabs(svn_path):
            raise HTTPException(status_code=400, detail="svn_path must be absolute")
        svn_path_obj = Path(svn_path)
        if not svn_path_obj.exists():
            raise HTTPException(status_code=404, detail=f"svn_path not found: {svn_path}")
        if svn_path_obj.suffix.lower() == ".csv":
            svn_blob = extract_from_csv_file(svn_path_obj)
        else:
            svn_blob = extract_from_excel_file(svn_path_obj)

    if not checklist_blob and payload.get("checklist_path"):
        checklist_path = payload["checklist_path"]
        if checklist_path.startswith("file://"):
            checklist_path = checklist_path[len("file://"):]
        if not os.path.isabs(checklist_path):
            raise HTTPException(status_code=400, detail="checklist_path must be absolute")
        if not Path(checklist_path).exists():
            raise HTTPException(status_code=404, detail=f"checklist_path not found: {checklist_path}")
        sheet_name = payload.get("sheet_name", "Test Scenario Remarks")
        checklist_list = extract_hyperlinks_with_versions_from_path(checklist_path, sheet_name=sheet_name)
        checklist_blob = {"filename": Path(checklist_path).name, "data": checklist_list, "count": len(checklist_list)}

    if not svn_blob or not checklist_blob:
        raise HTTPException(status_code=400, detail="Provide either 'svn' and 'checklist' blobs, or 'svn_path' and 'checklist_path'.")

    # Extract svn rows list
    if isinstance(svn_blob, dict) and "preview" in svn_blob:
        svn_rows = svn_blob["preview"]
    elif isinstance(svn_blob, dict) and "data" in svn_blob:
        svn_rows = svn_blob["data"]
    elif isinstance(svn_blob, list):
        svn_rows = svn_blob
    else:
        raise HTTPException(status_code=400, detail="Unrecognized svn blob format")

    # Extract checklist rows list
    if isinstance(checklist_blob, dict) and "data" in checklist_blob:
        checklist_rows = checklist_blob["data"]
    elif isinstance(checklist_blob, list):
        checklist_rows = checklist_blob
    else:
        raise HTTPException(status_code=400, detail="Unrecognized checklist blob format")

    return compare_data(svn_rows, checklist_rows, fuzzy_threshold)

@router.post("/validate-tc-traceability")
async def validate_tc_traceability_endpoint(file: UploadFile = File(...)):
    """
    Validate Test Case traceability from an uploaded Excel matrix.
    """
    from app.services.tc_traceability import validate_tc_traceability
    
    fname = file.filename or "traceability.xlsx"
    if not fname.lower().endswith((".xls", ".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="Please upload an .xls, .xlsx, or .xlsm file")
    
    dest = UPLOAD_DIR / fname
    counter = 0
    base = dest.stem
    ext = dest.suffix
    while dest.exists():
        counter += 1
        dest = UPLOAD_DIR / f"{base}_{counter}{ext}"
    
    try:
        save_upload_file(file, dest)
        result = validate_tc_traceability(dest)
        return result
    finally:
        try:
            dest.unlink()
        except Exception:
            pass

@router.post("/compare-cia")
async def compare_cia_endpoint(
    tc_file: UploadFile = File(...),
    cia_file: UploadFile = File(...)
):
    """
    Compare requirement IDs between TC (Traceability) file and CIA file.
    Both files should be Excel files with a 'General' sheet containing 'Requirements ID' column.
    """
    from app.services.cia_compare import compare_tc_vs_cia
    
    # Validate TC file
    tc_fname = tc_file.filename or "tc_file.xlsx"
    if not tc_fname.lower().endswith((".xls", ".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="TC file must be .xls, .xlsx, or .xlsm")
    
    # Validate CIA file
    cia_fname = cia_file.filename or "cia_file.xlsx"
    if not cia_fname.lower().endswith((".xls", ".xlsx", ".xlsm")):
        raise HTTPException(status_code=400, detail="CIA file must be .xls, .xlsx, or .xlsm")
    
    # Save TC file
    tc_dest = UPLOAD_DIR / tc_fname
    counter = 0
    base = tc_dest.stem
    ext = tc_dest.suffix
    while tc_dest.exists():
        counter += 1
        tc_dest = UPLOAD_DIR / f"{base}_{counter}{ext}"
    
    # Save CIA file
    cia_dest = UPLOAD_DIR / cia_fname
    counter = 0
    base = cia_dest.stem
    ext = cia_dest.suffix
    while cia_dest.exists():
        counter += 1
        cia_dest = UPLOAD_DIR / f"{base}_{counter}{ext}"
    
    try:
        save_upload_file(tc_file, tc_dest)
        save_upload_file(cia_file, cia_dest)
        
        result = compare_tc_vs_cia(
            tc_excel_path=str(tc_dest),
            cia_excel_path=str(cia_dest),
            tc_filename=tc_fname  # Pass original filename for SIT name derivation
        )
        return result
    finally:
        # Clean up temporary files
        try:
            tc_dest.unlink()
        except Exception:
            pass
        try:
            cia_dest.unlink()
        except Exception:
            pass
