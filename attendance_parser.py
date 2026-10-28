import re
import pandas as pd
from typing import Tuple, Dict, Any, Optional

def normalize_text(text: Any) -> str:
    """Normalize Arabic and English text for fuzzy column matching."""
    if text is None:
        return ""
    t = str(text).strip().lower()
    t = re.sub(r'[\u064B-\u065F\u0670]', '', t) # remove tashkeel
    t = t.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا').replace('ة', 'ه').replace('ى', 'ي')
    t = re.sub(r'[_\s\-]+', ' ', t)
    return t

COLUMN_SYNONYMS = {
    "student_id": [
        "id", "student id", "std id", "seat no", "كود", "كود الطالب", "رقم الطالب", 
        "الرقم التعريفي", "رقم الجلوس", "رقم القيد", "الرقم الجامعي", "الرقم", "academic id"
    ],
    "student_name": [
        "name", "student name", "std name", "full name", "اسم", "اسم الطالب", 
        "اسم الطالب رباعي", "الاسم ثلاثي", "الاسم", "student"
    ],
    "group": [
        "group", "grp", "الفرقة", "المجموعة", "جروب", "المجموعه"
    ],
    "section": [
        "section", "sec", "sec no", "السكشن", "سكشن", "رقم السكشن"
    ]
}

def detect_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    """
    Intelligently maps DataFrame columns to standard schema:
    student_id, student_name, group, section
    """
    matched = {"student_id": None, "student_name": None, "group": None, "section": None}
    cols = list(df.columns)
    normalized_cols = {c: normalize_text(c) for c in cols}
    
    # Pass 1: Exact / Synonym matching
    for target_key, syns in COLUMN_SYNONYMS.items():
        for col, norm in normalized_cols.items():
            if norm in syns:
                matched[target_key] = col
                break
                
    # Pass 2: Substring matching if still unmatched
    for target_key, syns in COLUMN_SYNONYMS.items():
        if matched[target_key] is None:
            for col, norm in normalized_cols.items():
                if any(syn in norm for syn in syns):
                    # make sure not already taken
                    if col not in matched.values():
                        matched[target_key] = col
                        break
                        
    # Pass 3: Heuristics based on content if ID or Name still missing
    if matched["student_id"] is None:
        for col in cols:
            # Check if column is numeric or looks like student ID
            non_na = df[col].dropna()
            if len(non_na) > 0:
                sample_vals = non_na.astype(str).tolist()[:10]
                if all(re.match(r'^\d{4,14}$', v.strip()) for v in sample_vals):
                    matched["student_id"] = col
                    break

    if matched["student_name"] is None:
        for col in cols:
            if col != matched["student_id"]:
                non_na = df[col].dropna()
                if len(non_na) > 0:
                    sample_vals = non_na.astype(str).tolist()[:10]
                    # Check if contains letters and spaces (names)
                    if any(" " in v for v in sample_vals):
                        matched["student_name"] = col
                        break
                        
    return matched

def parse_session_metadata_from_filename(filename: str) -> Dict[str, Any]:
    """
    Extracts session type (Lecture/Section), week number, section number from file name.
    """
    norm = normalize_text(filename)
    meta = {
        "type": "Lecture" if ("محاضره" in norm or "lec" in norm) else ("Section" if ("سكشن" in norm or "sec" in norm) else "Session"),
        "week": "Unknown",
        "section_assigned": "General",
        "raw_name": filename
    }
    
    # Extract week
    week_match = re.search(r'(?:week|اسبوع|w)[\s_\-]*(\d+)', norm)
    if week_match:
        meta["week"] = f"Week {week_match.group(1)}"
        
    # Extract section number
    sec_match = re.search(r'(?:section|سكشن|sec)[\s_\-]*(\d+)', norm)
    if sec_match:
        meta["section_assigned"] = f"Section {sec_match.group(1)}"
        
    return meta
