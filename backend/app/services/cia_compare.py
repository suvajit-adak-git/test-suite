"""
CIA Comparison Service
Compares requirement IDs between TC (Traceability) file and CIA file.
"""
import re
import pandas as pd
from typing import List, Dict, Tuple, Optional
from pathlib import Path
from fastapi import HTTPException


# Compiled regex patterns for better performance
RE_CL_PREFIX = re.compile(r'(?i)^\s*CL\b')
RE_NUMERIC = re.compile(r'^[+-]?\d+(\.\d+)?$')

def normalize_req_id(value) -> Optional[str]:
    """
    Normalize a Requirements ID to a canonical float-string with 3 decimal places.
    Returns None for values that should be ignored (e.g. d_ICR_xyz).
    
    Examples:
      "CL 12345678.001" -> "12345678.001"
      "12345678.001"    -> "12345678.001"
      "CL12345678.1"    -> "12345678.100"
      "d_ICR_xyz"       -> None
      
    Args:
        value: Raw requirement ID value
        
    Returns:
        Normalized requirement ID string or None if invalid
    """
    if pd.isna(value):
        return None

    s = str(value).strip()

    # If it begins with CL (case-insensitive) remove it and continue
    if RE_CL_PREFIX.match(s):
        s = RE_CL_PREFIX.sub('', s).strip()

    # Try parsing as float-like numeric id
    if RE_NUMERIC.match(s):
        try:
            val = float(s)
            # Format with exactly 3 decimals for canonical comparison
            return f"{val:.3f}"
        except Exception:
            return None

    # Handle comma-separated numbers (e.g. '12,345.001')
    s_nocomma = s.replace(',', '')
    if RE_NUMERIC.match(s_nocomma):
        try:
            val = float(s_nocomma)
            return f"{val:.3f}"
        except Exception:
            return None

    # Anything else (like 'd_ICR_xyz', 'some_text') -> ignore for comparison
    return None


def extract_sit_name_from_tc_filename(tc_filename: str) -> str:
    """
    Extract SIT name from TC filename by replacing _TC with _SIT.
    
    Examples:
        "CCPU_ICR_Exit_Cond_TC.xlsm" -> "CCPU_ICR_Exit_Cond_SIT"
        "Some_Test_TC.xlsx" -> "Some_Test_SIT"
    
    Args:
        tc_filename: TC file name
        
    Returns:
        SIT name to search for in CIA file
    """
    # Remove file extension
    name_without_ext = Path(tc_filename).stem
    
    # Replace _TC with _SIT (case-insensitive)
    sit_name = re.sub(r'_TC$', '_SIT', name_without_ext, flags=re.IGNORECASE)
    
    return sit_name


def extract_requirements_from_tc(
    excel_path: str, 
    sheet_name: str = 'General', 
    req_col_name: str = 'Requirements ID'
) -> Tuple[List[str], Dict[str, Optional[str]]]:
    """
    Read TC Excel workbook and extract normalized requirement IDs from General sheet.
    
    Args:
        excel_path: Path to TC Excel file
        sheet_name: Name of sheet to read (default: 'General')
        req_col_name: Name of the requirements column
        
    Returns:
        Tuple of (normalized_ids_list, original_to_normalized_mapping)
    """
    try:
        # Read without forcing header to auto-detect header row
        raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    except Exception as e:
        raise HTTPException(
            status_code=400, 
            detail=f"Failed to read sheet '{sheet_name}' from TC file: {str(e)}"
        )

    # Find header row
    header_row = None
    for idx, row in raw.iterrows():
        row_vals = row.astype(str).str.strip().str.lower().tolist()
        if req_col_name.lower() in row_vals:
            header_row = idx
            break

    if header_row is None:
        # Fallback to default header=0 read
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
        except Exception as e:
            raise HTTPException(
                status_code=400,
                detail=f"Could not find header row with '{req_col_name}' in TC sheet '{sheet_name}'"
            )
    else:
        cols = raw.iloc[header_row].astype(str).str.strip().tolist()
        df = raw.iloc[header_row + 1:].copy().reset_index(drop=True)
        df.columns = cols

    df.columns = df.columns.str.strip()
    
    if req_col_name not in df.columns:
        raise HTTPException(
            status_code=400,
            detail=f"Column '{req_col_name}' not found in TC sheet '{sheet_name}'. Available columns: {df.columns.tolist()}"
        )

    req_values = df[req_col_name].dropna().tolist()
    normalized = []
    mapping = {}
    
    for v in req_values:
        norm = normalize_req_id(v)
        mapping[str(v)] = norm
        if norm is not None:
            normalized.append(norm)

    # Deduplicate while keeping order
    seen = set()
    normalized_unique = []
    for n in normalized:
        if n not in seen:
            seen.add(n)
            normalized_unique.append(n)

    return normalized_unique, mapping


def extract_requirements_from_cia(
    excel_path: str,
    sit_name: str,
    sheet_name: str = 'HLR Change and Impact',
    hlr_col_name: str = 'New HLR ID',
    sit_col_name: str = 'SIT name'
) -> Tuple[List[str], Dict[str, Optional[str]]]:
    """
    Read CIA Excel workbook and extract normalized HLR IDs filtered by SIT name.
    
    Args:
        excel_path: Path to CIA Excel file
        sit_name: SIT name to filter by
        sheet_name: Name of sheet to read (default: 'HLR Change and Impact')
        hlr_col_name: Name of the HLR ID column (default: 'New HLR ID')
        sit_col_name: Name of the SIT name column (default: 'SIT name')
        
    Returns:
        Tuple of (normalized_ids_list, original_to_normalized_mapping)
    """
    try:
        # Read without forcing header to auto-detect header row
        raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    except Exception as e:
        raise HTTPException(
            status_code=400,
            detail=f"Failed to read sheet '{sheet_name}' from CIA file: {str(e)}"
        )

    # Find header row by looking for 'SIT name' column
    header_row = None
    for idx, row in raw.iterrows():
        row_vals = row.astype(str).str.strip().str.lower().tolist()
        if sit_col_name.lower() in row_vals:
            header_row = idx
            break

    if header_row is None:
        # Fallback to default header=0 read
        try:
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
        except Exception:
            raise HTTPException(
                status_code=400,
                detail=f"Could not find header row with '{sit_col_name}' in CIA sheet '{sheet_name}'"
            )
    else:
        cols = raw.iloc[header_row].astype(str).str.strip().tolist()
        df = raw.iloc[header_row + 1:].copy().reset_index(drop=True)
        df.columns = cols

    df.columns = df.columns.str.strip()
    
    # Check if required columns exist
    if sit_col_name not in df.columns:
        raise HTTPException(
            status_code=400,
            detail=f"Column '{sit_col_name}' not found in CIA sheet '{sheet_name}'. Available columns: {df.columns.tolist()}"
        )
    
    if hlr_col_name not in df.columns:
        raise HTTPException(
            status_code=400,
            detail=f"Column '{hlr_col_name}' not found in CIA sheet '{sheet_name}'. Available columns: {df.columns.tolist()}"
        )

    # Filter rows where 'SIT name' matches the input (handling semicolon-separated values)
    target_sit = sit_name.strip().lower()
    
    def check_sit_match(cell_value):
        if pd.isna(cell_value):
            return False
        # Split by semicolon, strip whitespace, and lower case
        sits = [s.strip().lower() for s in str(cell_value).split(';')]
        return target_sit in sits

    filtered_df = df[df[sit_col_name].apply(check_sit_match)]
    
    if len(filtered_df) == 0:
        # Provide helpful error message with available SIT names
        unique_sit_names = df[sit_col_name].dropna().unique()
        # Flatten the list of SIT names for the error message
        all_sits = set()
        for val in unique_sit_names:
            for s in str(val).split(';'):
                all_sits.add(s.strip())
        
        available_names = ', '.join([f"'{name}'" for name in sorted(list(all_sits))[:10]])
        raise HTTPException(
            status_code=400,
            detail=f"No matching records found for SIT name: '{sit_name}'. Available SIT names (first 10): {available_names}"
        )

    # Extract HLR IDs
    raw_hlr_ids = filtered_df[hlr_col_name].dropna().tolist()
    
    normalized = []
    mapping = {}
    
    for v in raw_hlr_ids:
        norm = normalize_req_id(v)
        mapping[str(v)] = norm
        if norm is not None:
            normalized.append(norm)

    # Deduplicate while keeping order
    seen = set()
    normalized_unique = []
    for n in normalized:
        if n not in seen:
            seen.add(n)
            normalized_unique.append(n)

    return normalized_unique, mapping


def compare_tc_vs_cia(
    tc_excel_path: str, 
    cia_excel_path: str,
    tc_filename: str = None
) -> Dict:
    """
    Compare requirement IDs between TC and CIA Excel files.
    
    Args:
        tc_excel_path: Path to TC Excel file
        cia_excel_path: Path to CIA Excel file
        tc_filename: Original TC filename (used to derive SIT name)
        
    Returns:
        Dictionary with comparison results in format expected by frontend
    """
    # Extract SIT name from TC filename
    if tc_filename is None:
        tc_filename = Path(tc_excel_path).name
    
    sit_name = extract_sit_name_from_tc_filename(tc_filename)
    
    # Extract requirements from TC file (General sheet, Requirements ID column)
    tc_ids, tc_map = extract_requirements_from_tc(tc_excel_path)
    
    # Extract requirements from CIA file (HLR Change and Impact sheet, New HLR ID column, filtered by SIT name)
    cia_ids, cia_map = extract_requirements_from_cia(cia_excel_path, sit_name)

    set_tc = set(tc_ids)
    set_cia = set(cia_ids)

    only_in_tc = sorted(list(set_tc - set_cia))
    only_in_cia = sorted(list(set_cia - set_tc))
    common = sorted(list(set_tc & set_cia))

    # Build results array for frontend
    results = []
    
    # Add common requirements (passed)
    for req_id in common:
        results.append({
            "tc_requirement": req_id,
            "cia_requirement": req_id,
            "status": "pass"
        })
    
    # Intelligent alignment for mismatches
    # Group by integer part (base ID) to find potential matches
    # e.g. 12345.001 and 12345.002 should be paired
    
    tc_groups = {}
    cia_groups = {}
    
    for req_id in only_in_tc:
        base_id = req_id.split('.')[0]
        if base_id not in tc_groups:
            tc_groups[base_id] = []
        tc_groups[base_id].append(req_id)
        
    for req_id in only_in_cia:
        base_id = req_id.split('.')[0]
        if base_id not in cia_groups:
            cia_groups[base_id] = []
        cia_groups[base_id].append(req_id)
        
    # Find all unique base IDs involved in mismatches
    all_bases = set(tc_groups.keys()) | set(cia_groups.keys())
    
    for base_id in all_bases:
        tc_items = sorted(tc_groups.get(base_id, []))
        cia_items = sorted(cia_groups.get(base_id, []))
        
        # Pair them up
        max_len = max(len(tc_items), len(cia_items))
        
        for i in range(max_len):
            tc_req = tc_items[i] if i < len(tc_items) else ""
            cia_req = cia_items[i] if i < len(cia_items) else ""
            
            results.append({
                "tc_requirement": tc_req,
                "cia_requirement": cia_req,
                "status": "fail"
            })

    # Sort results by requirement ID (TC if present, else CIA)
    results.sort(key=lambda x: x['tc_requirement'] or x['cia_requirement'])

    # Calculate summary
    total = len(results)
    passed = len(common)
    failed = total - passed

    return {
        "summary": {
            "total_requirements": total,
            "passed": passed,
            "failed": failed
        },
        "results": results,
        "details": {
            "sit_name": sit_name,
            "tc_count": len(tc_ids),
            "cia_count": len(cia_ids),
            "common_count": len(common),
            "only_in_tc_count": len(only_in_tc),
            "only_in_cia_count": len(only_in_cia)
        }
    }
