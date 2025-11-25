from typing import Dict, Any, List
from app.utils.common import (
    normalize_filename_for_match,
    normalize_version_string,
    extract_int_from_version,
    fuzzy_best_match
)

def compare_data(svn_rows: List[Dict[str, Any]], checklist_rows: List[Dict[str, Any]], fuzzy_threshold: float = 0.85) -> Dict[str, Any]:
    # Build canonical maps keyed by normalized filename (no extension)
    # Value is a LIST of entries to handle collisions (e.g. same name, different extension)
    svn_map: Dict[str, List[Dict[str, Any]]] = {}
    for r in svn_rows:
        # support both "File" or lowercase/other keys
        filename = r.get("File") if "File" in r else r.get("file") or r.get("Filename") or r.get("filename")
        if not filename:
            continue
        norm = normalize_filename_for_match(filename)
        svn_rev_raw = r.get("Last Changed Revision") or r.get("Last Changed Revision".lower()) or r.get("WC Revision") or r.get("revision") or r.get("Revision")
        svn_auth = r.get("Last Changed Author") or r.get("Last Changed Author".lower()) or r.get("last changed author") or r.get("Last Changed Author")
        svn_date = r.get("Last Changed Date") or r.get("Last Changed Date".lower()) or r.get("last changed date")
        
        entry = {
            "raw": r,
            "norm_name": norm,
            "filename_original": filename,
            "last_changed_revision_raw": normalize_version_string(svn_rev_raw),
            "last_changed_revision_int": extract_int_from_version(svn_rev_raw),
            "last_changed_author": svn_auth,
            "last_changed_date": svn_date,
            "matched": False
        }
        
        if norm not in svn_map:
            svn_map[norm] = []
        svn_map[norm].append(entry)

    checklist_map: Dict[str, List[Dict[str, Any]]] = {}
    checklist_norm_list: List[str] = []
    for r in checklist_rows:
        filename = r.get("filename") or r.get("Filename") or r.get("File")
        if not filename:
            continue
        norm = normalize_filename_for_match(filename)
        checklist_norm_list.append(norm)
        
        entry = {
            "raw": r,
            "norm_name": norm,
            "filename_original": filename,
            "version_closed_raw": normalize_version_string(r.get("version_closed") or r.get("Version") or r.get("version")),
            "version_closed_int": extract_int_from_version(r.get("version_closed") or r.get("Version") or r.get("version")),
            "matched": False
        }
        
        if norm not in checklist_map:
            checklist_map[norm] = []
        checklist_map[norm].append(entry)

    # Matching process
    matches = []
    mismatches = []
    only_in_svn = []
    only_in_checklist = []

    # First pass: exact normalized matches
    for s_key, s_entries in svn_map.items():
        if s_key in checklist_map:
            c_entries = checklist_map[s_key]
            
            # Try to match up entries within this group
            # Strategy: 
            # 1. Exact filename match (case-insensitive)
            # 2. If 1-to-1 remaining, match them
            
            # Helper to find match
            for s_entry in s_entries:
                if s_entry["matched"]:
                    continue
                    
                best_c_match = None
                
                # 1. Try exact filename match
                for c_entry in c_entries:
                    if not c_entry["matched"] and s_entry["filename_original"].lower() == c_entry["filename_original"].lower():
                        best_c_match = c_entry
                        break
                
                # 2. If no exact match, and both have only 1 unmatched entry, match them
                if not best_c_match:
                    unmatched_s = [x for x in s_entries if not x["matched"]]
                    unmatched_c = [x for x in c_entries if not x["matched"]]
                    if len(unmatched_s) == 1 and len(unmatched_c) == 1:
                        best_c_match = unmatched_c[0]
                
                if best_c_match:
                    # Found a match
                    s_entry["matched"] = True
                    best_c_match["matched"] = True
                    
                    s_ver_int = s_entry["last_changed_revision_int"]
                    c_ver_int = best_c_match["version_closed_int"]
                    s_ver_raw = s_entry["last_changed_revision_raw"]
                    c_ver_raw = best_c_match["version_closed_raw"]

                    version_equal = False
                    if s_ver_int is not None and c_ver_int is not None:
                        version_equal = (s_ver_int == c_ver_int)
                    else:
                        version_equal = (s_ver_raw != "" and c_ver_raw != "" and s_ver_raw == c_ver_raw)

                    result_entry = {
                        "filename": s_entry["filename_original"],
                        "normalized_filename": s_key,
                        "matched_checklist_filename": best_c_match["filename_original"],
                        "svn_revision_raw": s_ver_raw,
                        "svn_revision_int": s_ver_int,
                        "checklist_version_raw": c_ver_raw,
                        "checklist_version_int": c_ver_int,
                        "last_changed_author": s_entry["last_changed_author"],
                        "last_changed_date": s_entry["last_changed_date"],
                        "match_type": "exact",
                        "score": 1.0
                    }
                    
                    if version_equal:
                        matches.append(result_entry)
                    else:
                        mismatches.append(result_entry)

    # Second pass: Fuzzy matching for unmatched SVN entries
    # Get all unmatched SVN entries
    unmatched_svn_entries = []
    for s_list in svn_map.values():
        for s in s_list:
            if not s["matched"]:
                unmatched_svn_entries.append(s)
                
    # Get all unmatched Checklist entries (candidates)
    unmatched_checklist_entries = []
    for c_list in checklist_map.values():
        for c in c_list:
            if not c["matched"]:
                unmatched_checklist_entries.append(c)
                
    # Map normalized names to list of unmatched checklist entries for fuzzy search
    # We need a list of unique normalized keys from unmatched checklist entries to run fuzzy match against
    checklist_candidate_keys = list(set(c["norm_name"] for c in unmatched_checklist_entries))

    for s_entry in unmatched_svn_entries:
        s_key = s_entry["norm_name"]
        best_candidate_key, score = fuzzy_best_match(s_key, checklist_candidate_keys)
        
        if best_candidate_key and score >= fuzzy_threshold:
            # Find an unmatched checklist entry with this key
            # If multiple, we just take the first one (fuzzy match is already imprecise)
            # Or we could try to find the "best" one among them?
            
            # Get candidates with this key
            candidates = [c for c in unmatched_checklist_entries if c["norm_name"] == best_candidate_key and not c["matched"]]
            
            if candidates:
                c_entry = candidates[0]
                c_entry["matched"] = True
                s_entry["matched"] = True
                
                # If we used up all candidates for this key, remove from search list (optional optimization)
                if len(candidates) == 1:
                    checklist_candidate_keys.remove(best_candidate_key)

                s_ver_int = s_entry["last_changed_revision_int"]
                c_ver_int = c_entry["version_closed_int"]
                s_ver_raw = s_entry["last_changed_revision_raw"]
                c_ver_raw = c_entry["version_closed_raw"]

                version_equal = False
                if s_ver_int is not None and c_ver_int is not None:
                    version_equal = (s_ver_int == c_ver_int)
                else:
                    version_equal = (s_ver_raw != "" and c_ver_raw != "" and s_ver_raw == c_ver_raw)

                result_entry = {
                    "filename": s_entry["filename_original"],
                    "normalized_filename": s_key,
                    "matched_checklist_filename": c_entry["filename_original"],
                    "matched_checklist_normalized": best_candidate_key,
                    "svn_revision_raw": s_ver_raw,
                    "svn_revision_int": s_ver_int,
                    "checklist_version_raw": c_ver_raw,
                    "checklist_version_int": c_ver_int,
                    "last_changed_author": s_entry["last_changed_author"],
                    "last_changed_date": s_entry["last_changed_date"],
                    "match_type": "fuzzy",
                    "score": score
                }
                if version_equal:
                    matches.append(result_entry)
                else:
                    mismatches.append(result_entry)
            else:
                # Should not happen if logic is correct
                pass
        
        if not s_entry["matched"]:
            # No fuzzy match found
            only_in_svn.append({
                "filename": s_entry["filename_original"],
                "normalized_filename": s_entry["norm_name"],
                "last_changed_revision_raw": s_entry["last_changed_revision_raw"],
                "last_changed_revision_int": s_entry["last_changed_revision_int"],
                "last_changed_author": s_entry["last_changed_author"],
                "last_changed_date": s_entry["last_changed_date"]
            })

    # Collect remaining unmatched checklist entries
    for c_list in checklist_map.values():
        for c in c_list:
            if not c["matched"]:
                only_in_checklist.append({
                    "filename": c["filename_original"],
                    "normalized_filename": c["norm_name"],
                    "version_closed_raw": c["version_closed_raw"],
                    "version_closed_int": c["version_closed_int"],
                    "raw": c["raw"]
                })

    return {
        "status": "ok",
        "summary": {
            "matches": len(matches),
            "mismatches": len(mismatches),
            "only_in_svn": len(only_in_svn),
            "only_in_checklist": len(only_in_checklist)
        },
        "matches": matches,
        "mismatches": mismatches,
        "only_in_svn": only_in_svn,
        "only_in_checklist": only_in_checklist
    }
