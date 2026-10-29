import sys
import os
import io
from typing import List, Dict, Any, Optional

# Ensure project root is in sys.path
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

import pandas as pd
import numpy as np
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Response
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from analytics_engine import AttendanceEngine
from report_generator import export_attendance_matrix_excel, generate_pdf_summary_report

app = FastAPI(title="UniAttend AI API", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def to_serializable(val: Any) -> Any:
    """Recursively convert numpy types to native Python types for clean JSON serialization."""
    if isinstance(val, (np.integer, int)):
        return int(val)
    elif isinstance(val, (np.floating, float)):
        if np.isnan(val) or np.isinf(val):
            return 0.0
        return float(val)
    elif isinstance(val, (np.bool_, bool)):
        return bool(val)
    elif isinstance(val, dict):
        return {k: to_serializable(v) for k, v in val.items()}
    elif isinstance(val, (list, tuple)):
        return [to_serializable(v) for v in val]
    return val

def clean_session_result(s: Dict[str, Any]) -> Dict[str, Any]:
    """Strip non-serializable objects like raw_df from session dictionary."""
    clean_s = {k: v for k, v in s.items() if k != "raw_df"}
    return to_serializable(clean_s)

def clean_cumulative(cum: Dict[str, Any]) -> Dict[str, Any]:
    """Convert cumulative calculations and matrix DataFrame to JSON-friendly records."""
    if not cum:
        return {}
    matrix_df = cum.get("matrix_df", pd.DataFrame())
    matrix_records = matrix_df.to_dict(orient="records") if not matrix_df.empty else []
    matrix_columns = list(matrix_df.columns) if not matrix_df.empty else []

    clean_cum = {
        "matrix_records": matrix_records,
        "matrix_columns": matrix_columns,
        "session_names": cum.get("session_names", []),
        "group_stats": cum.get("group_stats", {}),
        "section_stats": cum.get("section_stats", {}),
        "session_trends": cum.get("session_trends", []),
        "all_cross_sections": cum.get("all_cross_sections", []),
        "all_unregistered": cum.get("all_unregistered", []),
        "total_students": cum.get("total_students", 0),
        "overall_attendance_rate": cum.get("overall_attendance_rate", 0.0)
    }
    return to_serializable(clean_cum)

# Pydantic schemas for export endpoints
class ExcelExportPayload(BaseModel):
    matrix_records: List[Dict[str, Any]] = []
    matrix_columns: List[str] = []
    session_names: List[str] = []

class PDFExportPayload(BaseModel):
    course_title: str = "مقرر الفصل الدراسي الحالي"
    overall_stats: Dict[str, Any] = {}
    matrix_records: List[Dict[str, Any]] = []
    cross_sections: List[Dict[str, Any]] = []

@app.get("/api/health")
async def health_check():
    return {"status": "ok", "app": "UniAttend AI", "engine": "FastAPI Stateless"}

@app.post("/api/demo")
async def load_demo_data():
    """Load sample data files bundled with the project."""
    engine = AttendanceEngine()
    sample_dir = os.path.join(ROOT_DIR, "sample_data")
    if not os.path.exists(sample_dir):
        raise HTTPException(status_code=404, detail="sample_data directory not found")

    master_file = os.path.join(sample_dir, "00_Master_Roster.xlsx")
    m_res = {"status": "skipped", "total_students": 0}
    if os.path.exists(master_file):
        m_res = engine.load_master_roster(master_file)

    sample_files = sorted([f for f in os.listdir(sample_dir) if f.endswith(".xlsx") and f != "00_Master_Roster.xlsx"])
    results = []
    for sf in sample_files:
        path = os.path.join(sample_dir, sf)
        s_res = engine.process_session_file(path, sf)
        if s_res.get("status") == "success":
            results.append(s_res)

    cum_data = engine.build_cumulative_matrix(results) if results else {}

    return {
        "status": "success",
        "master_loaded": m_res.get("status") == "success",
        "master_total": m_res.get("total_students", 0),
        "sessions": [clean_session_result(s) for s in results],
        "cumulative": clean_cumulative(cum_data),
        "errors": []
    }

@app.post("/api/check-master")
async def check_master(master_file: UploadFile = File(...)):
    """Validate master roster file and return student count and divisions."""
    try:
        content = await master_file.read()
        engine = AttendanceEngine()
        res = engine.load_master_roster(io.BytesIO(content), master_file.filename or "Master_Roster.xlsx")
        if res.get("status") == "success":
            return {
                "status": "success",
                "filename": master_file.filename,
                "total_students": res.get("total_students", 0),
                "groups": res.get("groups", []),
                "sections": res.get("sections", [])
            }
        else:
            return {
                "status": "error",
                "filename": master_file.filename,
                "message": res.get("message", "لم يتم العثور على بيانات صالحة")
            }
    except Exception as e:
        return {
            "status": "error",
            "filename": master_file.filename,
            "message": f"خطأ أثناء قراءة الملف: {str(e)}"
        }

@app.post("/api/analyze")
async def analyze_sessions(
    master_file: Optional[UploadFile] = File(None),
    session_files: List[UploadFile] = File(...)
):
    """Process uploaded sessions against master roster (if provided) and compute analytics."""
    engine = AttendanceEngine()
    master_loaded = False
    master_total = 0

    if master_file and master_file.filename:
        try:
            m_content = await master_file.read()
            if m_content:
                m_res = engine.load_master_roster(io.BytesIO(m_content), master_file.filename)
                if m_res.get("status") == "success":
                    master_loaded = True
                    master_total = m_res.get("total_students", 0)
        except Exception:
            pass

    results = []
    errors = []
    for sf in session_files:
        if not sf.filename:
            continue
        try:
            s_content = await sf.read()
            s_res = engine.process_session_file(io.BytesIO(s_content), sf.filename)
            if s_res.get("status") == "success":
                results.append(s_res)
            else:
                errors.append({
                    "filename": sf.filename,
                    "message": s_res.get("message", "فشل معالجة الكشف")
                })
        except Exception as e:
            errors.append({
                "filename": sf.filename,
                "message": str(e)
            })

    cum_data = engine.build_cumulative_matrix(results) if results else {}

    return {
        "status": "success" if results else "error",
        "master_loaded": master_loaded,
        "master_total": master_total,
        "sessions": [clean_session_result(s) for s in results],
        "cumulative": clean_cumulative(cum_data),
        "errors": errors
    }

@app.post("/api/export/excel")
async def export_excel(payload: ExcelExportPayload):
    """Generate and return styled Excel attendance matrix workbook."""
    try:
        matrix_df = pd.DataFrame(payload.matrix_records)
        if payload.matrix_columns and not matrix_df.empty:
            cols = [c for c in payload.matrix_columns if c in matrix_df.columns]
            matrix_df = matrix_df[cols]

        excel_bytes = export_attendance_matrix_excel(matrix_df, payload.session_names)
        return Response(
            content=excel_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": 'attachment; filename="Attendance_Matrix_Report.xlsx"'
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Export failed: {str(e)}")

@app.post("/api/export/pdf")
async def export_pdf(payload: PDFExportPayload):
    """Generate and return executive PDF summary report."""
    try:
        top_absent_df = pd.DataFrame(payload.matrix_records)
        pdf_bytes = generate_pdf_summary_report(
            course_title=payload.course_title,
            overall_stats=payload.overall_stats,
            top_absent_df=top_absent_df,
            cross_sections=payload.cross_sections
        )
        return Response(
            content=pdf_bytes,
            media_type="application/pdf",
            headers={
                "Content-Disposition": 'attachment; filename="Attendance_Executive_Summary.pdf"'
            }
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"PDF generation failed: {str(e)}")

# Serve frontend
@app.get("/")
async def serve_index():
    index_path = os.path.join(ROOT_DIR, "public", "index.html")
    if os.path.exists(index_path):
        return FileResponse(index_path, media_type="text/html")
    return {"message": "UniAttend AI Backend Active"}

# Mount public directory for static assets if needed
public_dir = os.path.join(ROOT_DIR, "public")
if os.path.exists(public_dir):
    app.mount("/public", StaticFiles(directory=public_dir), name="public")
