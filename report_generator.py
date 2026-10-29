import io
import os
import re
import pandas as pd
from typing import Dict, Any, List
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as ReportLabImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Register Arabic-capable Unicode TrueType font (Tajawal bundled or Windows system fonts)
FONT_NAME = "Helvetica"
FONT_BOLD = "Helvetica-Bold"

FONTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
tajawal_reg = os.path.join(FONTS_DIR, "Tajawal-Regular.ttf")
tajawal_bold = os.path.join(FONTS_DIR, "Tajawal-Bold.ttf")

try:
    if os.path.exists(tajawal_reg):
        pdfmetrics.registerFont(TTFont('ArabicFont', tajawal_reg))
        if os.path.exists(tajawal_bold):
            pdfmetrics.registerFont(TTFont('ArabicFont-Bold', tajawal_bold))
            FONT_BOLD = "ArabicFont-Bold"
        else:
            FONT_BOLD = "ArabicFont"
        FONT_NAME = "ArabicFont"
    elif os.path.exists("C:/Windows/Fonts/arial.ttf"):
        pdfmetrics.registerFont(TTFont('ArabicFont', 'C:/Windows/Fonts/arial.ttf'))
        if os.path.exists("C:/Windows/Fonts/arialbd.ttf"):
            pdfmetrics.registerFont(TTFont('ArabicFont-Bold', 'C:/Windows/Fonts/arialbd.ttf'))
            FONT_BOLD = "ArabicFont-Bold"
        else:
            FONT_BOLD = "ArabicFont"
        FONT_NAME = "ArabicFont"
    elif os.path.exists("C:/Windows/Fonts/tahoma.ttf"):
        pdfmetrics.registerFont(TTFont('ArabicFont', 'C:/Windows/Fonts/tahoma.ttf'))
        if os.path.exists("C:/Windows/Fonts/tahomabd.ttf"):
            pdfmetrics.registerFont(TTFont('ArabicFont-Bold', 'C:/Windows/Fonts/tahomabd.ttf'))
            FONT_BOLD = "ArabicFont-Bold"
        else:
            FONT_BOLD = "ArabicFont"
        FONT_NAME = "ArabicFont"
except Exception as e:
    print(f"Font registration notice: {e}")

# Arabic reshaper & bidi for RTL text in ReportLab
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    HAS_ARABIC_SUPPORT = True
except ImportError:
    HAS_ARABIC_SUPPORT = False

def format_ar(text: Any) -> str:
    """Reshape Arabic letters and apply BiDi algorithm so it prints connected RTL properly."""
    if text is None:
        return ""
    s = str(text).strip()
    if not s:
        return ""
    if re.search(r'[\u0600-\u06FF]', s) and HAS_ARABIC_SUPPORT:
        reshaped = arabic_reshaper.reshape(s)
        return get_display(reshaped)
    return s

def export_attendance_matrix_excel(matrix_df: pd.DataFrame, session_names: List[str]) -> bytes:
    """Exports beautifully formatted Excel attendance sheet with conditional styling."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        matrix_df.to_excel(writer, sheet_name='تقرير الحضور الشامل', index=False)
        workbook = writer.book
        worksheet = writer.sheets['تقرير الحضور الشامل']
        
        # Formats
        header_fmt = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'fg_color': '#0F172A',
            'font_color': '#FFFFFF',
            'border': 1
        })
        present_fmt = workbook.add_format({
            'bg_color': '#DCFCE7',
            'font_color': '#166534',
            'align': 'center'
        })
        absent_fmt = workbook.add_format({
            'bg_color': '#FEE2E2',
            'font_color': '#991B1B',
            'align': 'center'
        })

        # Apply header formatting
        num_rows = len(matrix_df)
        for col_num, value in enumerate(matrix_df.columns.values):
            worksheet.write(0, col_num, value, header_fmt)
            if num_rows > 0:
                col_data_len = matrix_df[value].astype(str).map(len).max()
                if pd.isna(col_data_len):
                    col_data_len = 0
            else:
                col_data_len = 0
            max_len = max(len(str(value)), int(col_data_len)) + 3
            worksheet.set_column(col_num, col_num, min(max(max_len, 14), 32))

        # Conditional formatting for present/absent/exempt
        if num_rows > 0:
            exempt_fmt = workbook.add_format({
                'bg_color': '#F1F5F9',
                'font_color': '#94A3B8',
                'align': 'center'
            })
            for s_col in session_names:
                if s_col in matrix_df.columns:
                    col_idx = list(matrix_df.columns).index(s_col)
                    worksheet.conditional_format(1, col_idx, num_rows, col_idx, {
                        'type': 'cell',
                        'criteria': '==',
                        'value': '"حاضر"',
                        'format': present_fmt
                    })
                    worksheet.conditional_format(1, col_idx, num_rows, col_idx, {
                        'type': 'cell',
                        'criteria': '==',
                        'value': '"غائب"',
                        'format': absent_fmt
                    })
                    worksheet.conditional_format(1, col_idx, num_rows, col_idx, {
                        'type': 'cell',
                        'criteria': '==',
                        'value': '"-"',
                        'format': exempt_fmt
                    })
            
    output.seek(0)
    return output.getvalue()

def generate_pdf_summary_report(course_title: str,
                                overall_stats: Dict[str, Any],
                                top_absent_df: pd.DataFrame,
                                cross_sections: List[Dict[str, Any]],
                                chart_images: List[bytes] = None) -> bytes:
    """
    Generates a clean, executive PDF report containing summary stats,
    high absence students table, cross-section attendees, and embedded graphs
    with full Arabic Unicode font support.
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=30
    )
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'DocTitle',
        parent=styles['Heading1'],
        fontName=FONT_BOLD,
        fontSize=18,
        leading=22,
        textColor=colors.HexColor("#0F172A"),
        alignment=1, # Center
        spaceAfter=10
    )
    subtitle_style = ParagraphStyle(
        'DocSubtitle',
        parent=styles['Normal'],
        fontName=FONT_NAME,
        fontSize=10,
        textColor=colors.HexColor("#64748B"),
        alignment=1,
        spaceAfter=18
    )
    section_h2 = ParagraphStyle(
        'SecH2',
        parent=styles['Heading2'],
        fontName=FONT_BOLD,
        fontSize=13,
        leading=16,
        textColor=colors.HexColor("#1E293B"),
        spaceBefore=14,
        spaceAfter=8
    )
    
    story = []
    
    # Header & Subtitle
    story.append(Paragraph(f"<b>{format_ar('تقرير تحليلات الحضور الأكاديمي الشامل')}</b>", title_style))
    course_txt = format_ar(f"المقرر: {course_title} | منصة UniAttend AI")
    story.append(Paragraph(course_txt, subtitle_style))
    story.append(Spacer(1, 10))
    
    # KPI Table
    tot_std = format_ar(f"إجمالي الطلاب: {overall_stats.get('total_students', 0)}")
    tot_sess = format_ar(f"إجمالي الجلسات: {len(overall_stats.get('session_names', []))}")
    avg_att = format_ar(f"متوسط الحضور: {overall_stats.get('overall_attendance_rate', 0)}%")
    
    kpi_data = [[tot_std, tot_sess, avg_att]]
    kpi_table = Table(kpi_data, colWidths=[175, 175, 175])
    kpi_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F8FAFC")),
        ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor("#0F172A")),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,-1), FONT_BOLD),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ('TOPPADDING', (0,0), (-1,-1), 10),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#CBD5E1")),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 15))
    
    # Embed Charts if available
    if chart_images:
        story.append(Paragraph(format_ar("الرسوم البيانية والإحصائية"), section_h2))
        for img_bytes in chart_images:
            img_io = io.BytesIO(img_bytes)
            story.append(ReportLabImage(img_io, width=520, height=260))
            story.append(Spacer(1, 10))
            
    # Cross-Section / Swapped Attendees Table (if any)
    if cross_sections:
        story.append(Paragraph(format_ar("كشف الطلاب الحاضرين في سكشن بديل (للمعرفة الإحصائية)"), section_h2))
        cs_rows = [[
            format_ar("كود الطالب"),
            format_ar("اسم الطالب"),
            format_ar("السكشن الأصلي"),
            format_ar("حضر في")
        ]]
        for cs in cross_sections[:15]:
            cs_rows.append([
                str(cs.get("student_id", "")),
                format_ar(cs.get("student_name", "")),
                format_ar(cs.get("assigned_section", "")),
                format_ar(cs.get("attended_section", ""))
            ])
        cs_table = Table(cs_rows, colWidths=[85, 205, 120, 120])
        cs_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0284C7")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,-1), FONT_NAME),
            ('FONTNAME', (0,0), (-1,0), FONT_BOLD),
            ('FONTSIZE', (0,0), (-1,-1), 8.5),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")])
        ]))
        story.append(cs_table)
        story.append(Spacer(1, 15))
        
    # Top Students with Low Attendance Table
    if not top_absent_df.empty:
        story.append(Paragraph(format_ar("قائمة الطلاب الأكثر غياباً (تحتاج للمتابعة)"), section_h2))
        sub_df = top_absent_df.head(15)
        table_rows = [[
            format_ar("كود الطالب"),
            format_ar("اسم الطالب"),
            format_ar("المجموعة"),
            format_ar("السكشن"),
            format_ar("عدد الحضور"),
            format_ar("النسبة %")
        ]]
        for _, row in sub_df.iterrows():
            table_rows.append([
                str(row.get("student_id", "")),
                format_ar(row.get("student_name", "")),
                format_ar(row.get("group", "")),
                format_ar(row.get("section", "")),
                f"{row.get('إجمالي الحضور', 0)} / {row.get('إجمالي الجلسات', 0)}",
                f"{row.get('نسبة الحضور %', 0)}%"
            ])
        t = Table(table_rows, colWidths=[75, 195, 75, 75, 55, 55])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0F172A")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,-1), FONT_NAME),
            ('FONTNAME', (0,0), (-1,0), FONT_BOLD),
            ('FONTSIZE', (0,0), (-1,-1), 8.5),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E1")),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")])
        ]))
        story.append(t)
        
    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()
