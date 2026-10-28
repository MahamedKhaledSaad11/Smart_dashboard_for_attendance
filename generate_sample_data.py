import os
import pandas as pd
import numpy as np

def generate_sample_datasets(output_dir="sample_data"):
    os.makedirs(output_dir, exist_ok=True)
    
    # 1. Master Roster (100 students across 2 groups and 4 sections)
    first_names = ["أحمد", "محمد", "محمود", "عمر", "علي", "يوسف", "إبراهيم", "كريم", "خالد", "حسن",
                   "مريم", "سارة", "نور", "فاطمة", "آية", "ياسمين", "هدى", "رنا", "ندى", "شهد"]
    last_names = ["السيد", "محمود", "علي", "عثمان", "حسن", "إبراهيم", "فاروق", "عبدالرحمن", "مصطفى", "شريف"]
    
    np.random.seed(42)
    students = []
    student_id_base = 2024001
    
    for i in range(100):
        fname = np.random.choice(first_names)
        lname = np.random.choice(last_names)
        name = f"{fname} {lname} {i+1}"
        sid = str(student_id_base + i)
        # Groups: G1, G2
        group = "Group 1" if i < 50 else "Group 2"
        # Sections: Sec 1, Sec 2 (for G1), Sec 3, Sec 4 (for G2)
        if group == "Group 1":
            section = "Section 1" if i < 25 else "Section 2"
        else:
            section = "Section 3" if i < 75 else "Section 4"
            
        students.append({
            "Student ID": sid,
            "Student Name": name,
            "Group": group,
            "Section": section,
            "Email": f"std_{sid}@univ.edu.eg"
        })
        
    df_master = pd.DataFrame(students)
    master_path = os.path.join(output_dir, "00_Master_Roster.xlsx")
    df_master.to_excel(master_path, index=False)
    print(f"Created: {master_path}")
    
    # Generate 3 Lectures (all students invited)
    # Some students attend, some skip, and 2 unregistered students show up
    for week in [1, 2, 3]:
        # Attendance probability varies
        prob = 0.85 if week == 1 else (0.75 if week == 2 else 0.65)
        attendees = []
        for s in students:
            if np.random.rand() < prob:
                attendees.append({
                    "كود الطالب": s["Student ID"],
                    "اسم الطالب": s["Student Name"],
                    "المجموعة": s["Group"],
                    "السكشن": s["Section"],
                    "تاريخ المحاضرة": f"2026-10-0{week*7 % 30 + 1}"
                })
        # Add 2 out-of-roster students
        attendees.append({
            "كود الطالب": f"99990{week}",
            "اسم الطالب": f"طالب زائر من جامعة أخرى {week}",
            "المجموعة": "External",
            "السكشن": "External",
            "تاريخ المحاضرة": f"2026-10-0{week*7 % 30 + 1}"
        })
        df_lec = pd.DataFrame(attendees)
        lec_path = os.path.join(output_dir, f"Lecture_Week_{week}.xlsx")
        df_lec.to_excel(lec_path, index=False)
        print(f"Created: {lec_path}")

    # Generate 3 Sections for each section (Sec 1, Sec 2, Sec 3, Sec 4)
    # Also add someone attending a wrong section!
    for week in [1, 2, 3]:
        for sec_num in [1, 2, 3, 4]:
            sec_name = f"Section {sec_num}"
            # Expected students for this section
            sec_students = [s for s in students if s["Section"] == sec_name]
            attendees = []
            for s in sec_students:
                if np.random.rand() < 0.80:
                    attendees.append({
                        "Student ID": s["Student ID"],
                        "Name": s["Student Name"],
                        "Assigned Section": s["Section"],
                        "Week": f"Week {week}"
                    })
            
            # Cross-section attendee (Student from another section attends here!)
            wrong_s = [s for s in students if s["Section"] != sec_name][np.random.randint(0, 50)]
            attendees.append({
                "Student ID": wrong_s["Student ID"],
                "Name": wrong_s["Student Name"],
                "Assigned Section": f"Swapped from {wrong_s['Section']}",
                "Week": f"Week {week}"
            })
            
            df_sec = pd.DataFrame(attendees)
            sec_path = os.path.join(output_dir, f"Section_{sec_num}_Week_{week}.xlsx")
            df_sec.to_excel(sec_path, index=False)
            print(f"Created: {sec_path}")

if __name__ == "__main__":
    generate_sample_datasets()
