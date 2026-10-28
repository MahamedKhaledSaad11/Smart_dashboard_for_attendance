import os
import io
import pandas as pd
import numpy as np
from typing import List, Dict, Any, Optional
from attendance_parser import detect_columns, parse_session_metadata_from_filename

class AttendanceEngine:
    def __init__(self):
        self.master_df: Optional[pd.DataFrame] = None
        self.sessions: List[Dict[str, Any]] = []
        
    def load_master_roster(self, file_content_or_path, filename: str = "Master_Roster.xlsx") -> Dict[str, Any]:
        """Loads master student list and standardizes columns."""
        try:
            df = pd.read_excel(file_content_or_path)
        except Exception as e:
            return {"status": "error", "message": f"فشل قراءة ملف الماستر شيت ({filename}): {str(e)}"}
            
        mapping = detect_columns(df)
        if not mapping["student_id"]:
            return {"status": "error", "message": "لم يتم العثور على عمود كود الطالب (ID) في الماستر شيت."}
            
        # Standardize DataFrame
        id_col = mapping["student_id"]
        # Drop rows where student_id is NaN
        df = df.dropna(subset=[id_col]).copy()
        clean_ids = df[id_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        valid_id_mask = clean_ids.ne("") & clean_ids.str.lower().ne("nan") & clean_ids.str.lower().ne("none")
        df = df[valid_id_mask].copy()
        clean_ids = clean_ids[valid_id_mask]

        clean_df = pd.DataFrame()
        clean_df["student_id"] = clean_ids

        name_col = mapping["student_name"]
        if name_col and name_col in df.columns:
            clean_df["student_name"] = df[name_col].fillna("Unknown").astype(str).str.strip()
            clean_df["student_name"] = clean_df["student_name"].replace(["nan", "None", ""], "Unknown")
        else:
            clean_df["student_name"] = "Unknown"

        grp_col = mapping["group"]
        if grp_col and grp_col in df.columns:
            clean_df["group"] = df[grp_col].fillna("General").astype(str).str.strip()
            clean_df["group"] = clean_df["group"].replace(["nan", "None", ""], "General")
        else:
            clean_df["group"] = "General"

        sec_col = mapping["section"]
        if sec_col and sec_col in df.columns:
            clean_df["section"] = df[sec_col].fillna("General").astype(str).str.strip()
            clean_df["section"] = clean_df["section"].replace(["nan", "None", ""], "General")
        else:
            clean_df["section"] = "General"
        
        # Remove duplicate IDs in master roster
        clean_df = clean_df.drop_duplicates(subset=["student_id"]).reset_index(drop=True)
        self.master_df = clean_df
        
        return {
            "status": "success",
            "total_students": len(clean_df),
            "groups": clean_df["group"].unique().tolist(),
            "sections": clean_df["section"].unique().tolist(),
            "mapping": mapping
        }

    def process_session_file(self, file_content_or_path, filename: str) -> Dict[str, Any]:
        """
        Parses single session file and cross-references against master roster
        to detect present students, absentees, cross-section attendees, and external visitors.
        """
        try:
            df = pd.read_excel(file_content_or_path)
        except Exception as e:
            return {"status": "error", "filename": filename, "message": f"فشل قراءة ملف الحضور ({filename}): {str(e)}"}

        meta = parse_session_metadata_from_filename(filename)
        mapping = detect_columns(df)
        
        if not mapping["student_id"]:
            return {"status": "error", "filename": filename, "message": "لم يتم التعرف على كود الطالب (ID)"}
            
        raw_ids = df[mapping["student_id"]].dropna().astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
        attendee_ids = [sid for sid in raw_ids.unique() if sid and sid.lower() not in ("nan", "none")]
        
        # Cross reference with master
        if self.master_df is not None:
            master = self.master_df.copy()
            # Mark attendance
            master["attended"] = master["student_id"].isin(attendee_ids)
            
            # Cross section attendees:
            # If session is specific to a Section (e.g. Section 1), who attended from other sections?
            assigned_sec = meta["section_assigned"]
            is_specific_section = (assigned_sec != "General" and meta["type"] == "Section")

            cross_section_students = []
            if is_specific_section:
                # Students who attended this session, but their assigned section in master is DIFFERENT
                cross_df = master[master["attended"] & (master["section"] != assigned_sec)]
                cross_section_students = cross_df[["student_id", "student_name", "group", "section"]].to_dict(orient="records")
            
            # Out of roster (unregistered / visitor)
            master_ids = set(master["student_id"])
            unregistered_ids = [sid for sid in attendee_ids if sid not in master_ids]
            unregistered_records = []
            if unregistered_ids:
                id_col = mapping["student_id"]
                unreg_df = df.dropna(subset=[id_col]).copy()
                unreg_df["_clean_id"] = unreg_df[id_col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                unreg_df = unreg_df[unreg_df["_clean_id"].isin(unregistered_ids)]
                for _, row in unreg_df.iterrows():
                    sid = row["_clean_id"]
                    sname = "Unknown"
                    if mapping["student_name"] and pd.notna(row[mapping["student_name"]]):
                        cname = str(row[mapping["student_name"]]).strip()
                        if cname and cname.lower() not in ("nan", "none"):
                            sname = cname
                    unregistered_records.append({
                        "student_id": sid,
                        "student_name": sname,
                        "raw_info": "غير مقيد في الماستر شيت"
                    })
                    
            # Target audience calculations
            target_audience_df = master[master["section"] == assigned_sec] if is_specific_section else master
            target_total = len(target_audience_df)
            target_present = int(target_audience_df["attended"].sum())
            target_absent = target_total - target_present
            rate = (target_present / target_total * 100) if target_total > 0 else 0.0

            # Present students: all students in master who attended this session (including cross-section)
            present_students = master[master["attended"]][["student_id", "student_name", "group", "section"]].to_dict(orient="records")
            
            # Target present students (only those belonging to assigned section, or all if General)
            target_present_students = target_audience_df[target_audience_df["attended"]][["student_id", "student_name", "group", "section"]].to_dict(orient="records")

            # Absent students: ONLY students in target audience who did NOT attend!
            absent_students = target_audience_df[~target_audience_df["attended"]][["student_id", "student_name", "group", "section"]].to_dict(orient="records")
        else:
            # No master provided: Attendance based only on file
            present_students = []
            id_col = mapping["student_id"]
            name_col = mapping["student_name"]
            grp_col = mapping["group"]
            sec_col = mapping["section"]
            clean_file_df = df.dropna(subset=[id_col]).copy()
            for _, row in clean_file_df.iterrows():
                sid = str(row[id_col]).strip().replace(".0", "")
                if not sid or sid.lower() in ("nan", "none"):
                    continue
                sname = str(row[name_col]).strip() if name_col and pd.notna(row[name_col]) else "Unknown"
                if sname.lower() in ("nan", "none", ""): sname = "Unknown"
                sgrp = str(row[grp_col]).strip() if grp_col and pd.notna(row[grp_col]) else "General"
                if sgrp.lower() in ("nan", "none", ""): sgrp = "General"
                ssec = str(row[sec_col]).strip() if sec_col and pd.notna(row[sec_col]) else "General"
                if ssec.lower() in ("nan", "none", ""): ssec = "General"
                present_students.append({
                    "student_id": sid,
                    "student_name": sname,
                    "group": sgrp,
                    "section": ssec,
                })
            target_present_students = present_students
            absent_students = []
            cross_section_students = []
            unregistered_records = []
            target_total = len(present_students)
            target_present = target_total
            target_absent = 0
            rate = 100.0

        session_result = {
            "status": "success",
            "filename": filename,
            "session_name": filename.replace(".xlsx", "").replace(".xls", ""),
            "metadata": meta,
            "total_attendees": len(attendee_ids),
            "target_total": target_total,
            "target_present": int(target_present),
            "target_absent": int(target_absent),
            "attendance_rate": round(rate, 1),
            "present_students": present_students,
            "target_present_students": target_present_students,
            "absent_students": absent_students,
            "cross_section_students": cross_section_students,
            "unregistered_records": unregistered_records,
            "raw_df": df
        }
        return session_result

    def build_cumulative_matrix(self, session_results: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Builds attendance matrix across all uploaded sessions.
        Columns: Student ID, Student Name, Group, Section, [Session 1, Session 2, ...], Total Attended, Total Sessions, Attendance %
        """
        if not session_results:
            return {}
            
        session_names = [s["session_name"] for s in session_results]
        
        # Base roster: Either master_df or union of attendees
        if self.master_df is not None:
            matrix_df = self.master_df.copy()
        else:
            all_students = {}
            for s in session_results:
                for p in s.get("present_students", []):
                    sid = p["student_id"]
                    if sid not in all_students:
                        all_students[sid] = p
            matrix_df = pd.DataFrame(list(all_students.values()))
            
        for s in session_results:
            s_name = s["session_name"]
            meta = s.get("metadata", {})
            assigned_sec = meta.get("section_assigned", "General")
            s_type = meta.get("type", "Session")
            is_specific_section = (assigned_sec != "General" and s_type == "Section")
            attended_ids = set([p["student_id"] for p in s.get("present_students", [])])
            
            def get_status(row):
                sid = row["student_id"]
                if sid in attended_ids:
                    return "حاضر"
                if is_specific_section:
                    sec = str(row.get("section", "General")).strip()
                    if sec == assigned_sec:
                        return "غائب"
                    else:
                        return "-"
                else:
                    return "غائب"
                    
            matrix_df[s_name] = matrix_df.apply(get_status, axis=1)
            
        # Add summary stats
        attendance_cols = session_names
        matrix_df["إجمالي الحضور"] = (matrix_df[attendance_cols] == "حاضر").sum(axis=1)
        matrix_df["إجمالي الجلسات"] = ((matrix_df[attendance_cols] == "حاضر") | (matrix_df[attendance_cols] == "غائب")).sum(axis=1)
        applicable_sessions = matrix_df["إجمالي الجلسات"].replace(0, np.nan)
        matrix_df["نسبة الحضور %"] = ((matrix_df["إجمالي الحضور"] / applicable_sessions) * 100).fillna(0.0).round(1)
        matrix_df["نسبة الغياب %"] = np.where(matrix_df["إجمالي الجلسات"] == 0, 0.0, (100 - matrix_df["نسبة الحضور %"]).round(1))
        
        # Sort by attendance rate ascending (at-risk on top)
        matrix_df = matrix_df.sort_values(by="نسبة الحضور %", ascending=True).reset_index(drop=True)
        
        # Aggregate statistics
        group_stats = matrix_df.groupby("group")["نسبة الحضور %"].mean().round(1).to_dict() if "group" in matrix_df else {}
        section_stats = matrix_df.groupby("section")["نسبة الحضور %"].mean().round(1).to_dict() if "section" in matrix_df else {}
        
        # Session by session trend
        session_trends = []
        for s in session_results:
            session_trends.append({
                "session": s["session_name"],
                "type": s["metadata"]["type"],
                "week": s["metadata"]["week"],
                "present_count": s["target_present"],
                "attendance_rate": s["attendance_rate"]
            })
            
        # Collect all cross section events
        all_cross_sections = []
        for s in session_results:
            for cs in s.get("cross_section_students", []):
                all_cross_sections.append({
                    "session": s["session_name"],
                    "student_id": cs["student_id"],
                    "student_name": cs["student_name"],
                    "assigned_section": cs["section"],
                    "attended_section": s["metadata"]["section_assigned"]
                })
                
        # Collect all unregistered visits
        all_unregistered = []
        for s in session_results:
            for unreg in s.get("unregistered_records", []):
                all_unregistered.append({
                    "session": s["session_name"],
                    "student_id": unreg["student_id"],
                    "student_name": unreg["student_name"],
                    "note": unreg["raw_info"]
                })

        overall_rate = round(float(matrix_df["نسبة الحضور %"].mean()), 1) if not matrix_df.empty and not pd.isna(matrix_df["نسبة الحضور %"].mean()) else 0.0

        return {
            "matrix_df": matrix_df,
            "session_names": session_names,
            "group_stats": group_stats,
            "section_stats": section_stats,
            "session_trends": session_trends,
            "all_cross_sections": all_cross_sections,
            "all_unregistered": all_unregistered,
            "total_students": len(matrix_df),
            "overall_attendance_rate": overall_rate
        }
