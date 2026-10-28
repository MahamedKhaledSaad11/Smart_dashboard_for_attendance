import os
import io
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from attendance_parser import detect_columns, parse_session_metadata_from_filename
from analytics_engine import AttendanceEngine
from report_generator import export_attendance_matrix_excel, generate_pdf_summary_report

# Page Config
st.set_page_config(
    page_title="UniAttend AI | منصة متابعة الحضور الأكاديمي",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# RTL & Refined High-Contrast Light-Mode Styling
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=Tajawal:wght@400;500;700;800&display=swap');

/* Global RTL Support */
html, body, [class*="css"], .stApp {
    font-family: 'Tajawal', 'Plus Jakarta Sans', sans-serif !important;
    direction: rtl !important;
    text-align: right !important;
    color: #0F172A !important;
    background-color: #F8FAFC !important;
}

/* Fix RTL for controls and headers */
h1, h2, h3, h4, h5, h6, p, label, .stMarkdown, .stSelectbox, .stMultiSelect {
    text-align: right !important;
    direction: rtl !important;
}

/* Sidebar adjustments for RTL */
section[data-testid="stSidebar"] {
    direction: rtl !important;
    text-align: right !important;
    border-left: 1px solid #E2E8F0 !important;
    border-right: none !important;
    background-color: #FFFFFF !important;
}

section[data-testid="stSidebar"] .stMarkdown, 
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p {
    text-align: right !important;
}

/* Header Hero Banner */
.hero-banner {
    background: linear-gradient(135deg, #1E1B4B 0%, #312E81 50%, #4338CA 100%);
    padding: 2rem 2.5rem;
    border-radius: 18px;
    color: #FFFFFF !important;
    margin-bottom: 2rem;
    border: 1px solid rgba(255, 255, 255, 0.15);
    box-shadow: 0 10px 25px -5px rgba(30, 27, 75, 0.25);
}

.hero-title {
    font-size: 2.1rem;
    font-weight: 800;
    margin: 0;
    line-height: 1.3;
    color: #FFFFFF !important;
}

.hero-subtitle {
    color: #E0E7FF !important;
    margin-top: 0.5rem;
    font-size: 1.05rem;
    max-width: 850px;
    line-height: 1.6;
}

/* Stat KPI Cards - High Contrast in Light Mode */
.stat-card {
    background: #FFFFFF !important;
    border-radius: 16px;
    padding: 1.5rem;
    border: 1.5px solid #E2E8F0;
    box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    transition: all 0.2s ease;
    text-align: right;
    direction: rtl;
}

.stat-card:hover {
    transform: translateY(-2px);
    border-color: #CBD5E1;
    box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08);
}

.stat-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 0.75rem;
}

.stat-label {
    font-size: 0.95rem;
    font-weight: 700;
    color: #475569 !important;
}

.stat-icon {
    width: 42px;
    height: 42px;
    border-radius: 12px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1.3rem;
}

.stat-value {
    font-size: 2.2rem;
    font-weight: 800;
    color: #0F172A !important;
    line-height: 1.1;
}

.stat-footer {
    margin-top: 0.6rem;
    font-size: 0.85rem;
    color: #64748B !important;
    font-weight: 500;
}

/* Tabs Styling */
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
    background-color: #E2E8F0;
    padding: 6px;
    border-radius: 14px;
    direction: rtl !important;
}

.stTabs [data-baseweb="tab"] {
    border-radius: 10px;
    padding: 10px 22px;
    font-weight: 700;
    font-size: 0.95rem;
    color: #475569 !important;
}

.stTabs [aria-selected="true"] {
    background-color: #FFFFFF !important;
    color: #4F46E5 !important;
    box-shadow: 0 2px 5px rgba(0, 0, 0, 0.08);
}

/* Glass Section Container */
.glass-container {
    background: #FFFFFF !important;
    border: 1.5px solid #E2E8F0;
    border-radius: 16px;
    padding: 1.5rem;
    margin-bottom: 1.5rem;
    box-shadow: 0 2px 4px rgba(0,0,0,0.02);
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# Session State Initialization
if "engine" not in st.session_state:
    st.session_state.engine = AttendanceEngine()
if "master_loaded" not in st.session_state:
    st.session_state.master_loaded = False
if "uploaded_sessions" not in st.session_state:
    st.session_state.uploaded_sessions = []
if "session_results" not in st.session_state:
    st.session_state.session_results = []

# Hero Banner (without version or impeccable tags)
st.markdown("""
<div class="hero-banner">
    <div class="hero-title">🎓 منصة متابعة وتحليل الحضور الأكاديمي</div>
    <div class="hero-subtitle">
        نظام مؤتمت للمعيدين وأعضاء هيئة التدريس لقراءة كشوف الإكسيل، كشف حالات الحضور في مواعيد سكاشن بديلة، وتوليد مصفوفة الحضور والتقارير الفورية.
    </div>
</div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown("### ⚙️ إدارة البيانات والملفات")
    
    # 1-Click Sample Loader
    if st.button("🚀 تحميل بيانات تجريبية (Load Demo)", use_container_width=True, type="secondary"):
        sample_dir = "sample_data"
        if os.path.exists(sample_dir):
            master_file = os.path.join(sample_dir, "00_Master_Roster.xlsx")
            if os.path.exists(master_file):
                m_res = st.session_state.engine.load_master_roster(master_file)
                if m_res.get("status") == "success":
                    st.session_state.master_loaded = True
                else:
                    st.error(m_res.get("message"))
                
                sample_files = sorted([f for f in os.listdir(sample_dir) if f.endswith(".xlsx") and f != "00_Master_Roster.xlsx"])
                results = []
                for sf in sample_files:
                    path = os.path.join(sample_dir, sf)
                    s_res = st.session_state.engine.process_session_file(path, sf)
                    if s_res.get("status") == "success":
                        results.append(s_res)
                    else:
                        st.error(f"{sf}: {s_res.get('message')}")
                if results:
                    st.session_state.session_results = results
                    st.success(f"تم تحميل {len(results)} جلسة تجريبية بنجاح!")
                    st.rerun()

    st.markdown("---")
    st.markdown("#### 1️⃣ القائمة المرجعية (Master Roster)")
    st.caption("ملف إكسيل يضم جميع الطلاب المسجلين بالـ ID والاسم والمجموعة والسكشن.")
    master_upload = st.file_uploader("رفع ملف الماستر شيت", type=["xlsx", "xls"], key="master_uploader")
    if master_upload is not None:
        with st.spinner("جاري قراءة الماستر شيت..."):
            res = st.session_state.engine.load_master_roster(master_upload, master_upload.name)
            if res.get("status") == "success":
                st.session_state.master_loaded = True
                st.success(f"✅ تم تحميل {res['total_students']} طالب بنجاح")
            else:
                st.error(res.get("message"))

    st.markdown("---")
    st.markdown("#### 2️⃣ ملفات الحضور (محاضرات / سكاشن)")
    st.caption("يمكنك سحب وإفلات ملف واحد أو عدة ملفات إكسيل دفعة واحدة.")
    session_files = st.file_uploader("رفع كشوف الحضور", type=["xlsx", "xls"], accept_multiple_files=True, key="sessions_uploader")
    if session_files:
        if st.button("معالجة الملفات المرفوعة", type="primary", use_container_width=True):
            with st.spinner("جاري المطابقة الآلية..."):
                results = []
                for f in session_files:
                    s_res = st.session_state.engine.process_session_file(f, f.name)
                    if s_res.get("status") == "success":
                        results.append(s_res)
                    else:
                        st.error(f"{f.name}: {s_res.get('message')}")
                if results:
                    st.session_state.session_results = results
                    st.success(f"تمت معالجة {len(results)} ملف حضور!")
                    st.rerun()

    if st.session_state.session_results:
        st.markdown("---")
        if st.button("🗑️ مسح كل البيانات والبدء من جديد", use_container_width=True):
            st.session_state.engine = AttendanceEngine()
            st.session_state.master_loaded = False
            st.session_state.session_results = []
            st.rerun()

# Empty State
if not st.session_state.session_results:
    st.markdown("""
    <div class="glass-container" style="text-align: center; padding: 3.5rem 2rem;">
        <div style="font-size: 3.5rem; margin-bottom: 1rem;">📁</div>
        <h3 style="color: #0F172A; margin-bottom: 0.5rem; font-weight: 700;">لا توجد كشوف حضور مضافة حالياً</h3>
        <p style="color: #475569; max-width: 520px; margin: 0 auto 1.5rem auto; font-size: 1rem;">
            ابدأ برفع ملف الماستر شيت وكشوف المحاضرات أو السكاشن من القائمة الجانبية، أو اضغط على <b>تحميل بيانات تجريبية</b> للاطلاع على التحليلات فوراً.
        </p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Build Cumulative Calculations
results = st.session_state.session_results
cum_data = st.session_state.engine.build_cumulative_matrix(results)
matrix_df = cum_data.get("matrix_df", pd.DataFrame())

# Main Navigation Tabs
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 لوحة التحليل التراكمي الشامل",
    "⚡ تحليل جلسة واحدة سريعة",
    "🔄 كشف الحضور في سكشن بديل",
    "👤 ملف الطالب الشامل",
    "📥 مركز التصدير المتقدم (Excel & PDF)"
])

# ----------------- TAB 1: CUMULATIVE DASHBOARD -----------------
with tab1:
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    
    # 4 Stat Cards
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-header">
                <span class="stat-label">الطلاب المسجلين</span>
                <div class="stat-icon" style="background: #EEF2FF; color: #4338CA;">👥</div>
            </div>
            <div class="stat-value">{cum_data.get('total_students', 0)}</div>
            <div class="stat-footer">طالب في الماستر شيت</div>
        </div>
        """, unsafe_allow_html=True)
    with k2:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-header">
                <span class="stat-label">الجلسات المسجلة</span>
                <div class="stat-icon" style="background: #ECFDF5; color: #047857;">📅</div>
            </div>
            <div class="stat-value">{len(results)}</div>
            <div class="stat-footer">محاضرات وسكاشن</div>
        </div>
        """, unsafe_allow_html=True)
    with k3:
        rate = cum_data.get('overall_attendance_rate', 0)
        rate_color = "#059669" if rate >= 75 else ("#D97706" if rate >= 50 else "#DC2626")
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-header">
                <span class="stat-label">متوسط الحضور العام</span>
                <div class="stat-icon" style="background: #FEF3C7; color: #B45309;">📈</div>
            </div>
            <div class="stat-value" style="color: {rate_color};">{rate}%</div>
            <div class="stat-footer">عبر كافة الجلسات</div>
        </div>
        """, unsafe_allow_html=True)
    with k4:
        cross_count = len(cum_data.get('all_cross_sections', []))
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-header">
                <span class="stat-label">حضور بسكشن بديل</span>
                <div class="stat-icon" style="background: #F3E8FF; color: #7E22CE;">🔄</div>
            </div>
            <div class="stat-value" style="color: #7E22CE;">{cross_count}</div>
            <div class="stat-footer">طالب بدل موعد سكشنه</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)
    
    # Visual Analytics Row
    c1, c2 = st.columns([3, 2])
    with c1:
        st.markdown("##### 📉 تطور نسب الحضور مع مرور الأسابيع")
        trend_df = pd.DataFrame(cum_data.get("session_trends", []))
        if not trend_df.empty:
            fig_trend = px.line(
                trend_df,
                x="session",
                y="attendance_rate",
                color="type",
                markers=True,
                labels={"session": "الجلسة", "attendance_rate": "نسبة الحضور %", "type": "النوع"},
                color_discrete_map={"Lecture": "#4338CA", "Section": "#0284C7", "Session": "#059669"}
            )
            fig_trend.update_layout(
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#FFFFFF",
                font=dict(color="#0F172A", family="Tajawal, sans-serif"),
                margin=dict(l=20, r=20, t=30, b=20),
                yaxis=dict(range=[0, 105], gridcolor="#E2E8F0", tickfont=dict(color="#0F172A")),
                xaxis=dict(gridcolor="#E2E8F0", tickfont=dict(color="#0F172A")),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(color="#0F172A")),
                hovermode="x unified"
            )
            st.plotly_chart(fig_trend, use_container_width=True)
            
    with c2:
        st.markdown("##### 👥 مقارنة الحضور حسب السكاشن")
        sec_stats = cum_data.get("section_stats", {})
        if sec_stats:
            sec_df = pd.DataFrame(list(sec_stats.items()), columns=["السكشن", "نسبة الحضور %"]).sort_values("نسبة الحضور %", ascending=False)
            fig_bar = px.bar(
                sec_df,
                x="السكشن",
                y="نسبة الحضور %",
                text="نسبة الحضور %",
                color="نسبة الحضور %",
                color_continuous_scale=["#6366F1", "#312E81"]
            )
            fig_bar.update_traces(
                texttemplate='%{text}%',
                textposition='outside',
                textfont=dict(color="#0F172A", size=12)
            )
            fig_bar.update_layout(
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#FFFFFF",
                font=dict(color="#0F172A", family="Tajawal, sans-serif"),
                margin=dict(l=20, r=20, t=30, b=20),
                yaxis=dict(range=[0, 115], gridcolor="#E2E8F0", tickfont=dict(color="#0F172A")),
                xaxis=dict(gridcolor="#E2E8F0", tickfont=dict(color="#0F172A")),
                coloraxis_showscale=False
            )
            st.plotly_chart(fig_bar, use_container_width=True)

    # Matrix Table with Multi-Filters (Group, Section, Session)
    st.markdown("---")
    st.markdown("##### 📋 مصفوفة الحضور التراكمية الشاملة")
    
    mf1, mf2, mf3, mf4 = st.columns([1, 1, 1, 1.5])
    with mf1:
        group_filter = st.multiselect("المجموعة (Group):", options=matrix_df["group"].unique() if "group" in matrix_df else [])
    with mf2:
        sec_filter = st.multiselect("السكشن (Section):", options=matrix_df["section"].unique() if "section" in matrix_df else [])
    with mf3:
        all_sessions = cum_data.get("session_names", [])
        session_filter = st.multiselect("تصفية بالأعمدة (محاضرات / سكاشن):", options=all_sessions)
    with mf4:
        search_query = st.text_input("🔍 بحث باسم الطالب أو الكود:", "")
        
    filtered_df = matrix_df.copy()
    if group_filter:
        filtered_df = filtered_df[filtered_df["group"].isin(group_filter)]
    if sec_filter:
        filtered_df = filtered_df[filtered_df["section"].isin(sec_filter)]
    if search_query:
        filtered_df = filtered_df[
            filtered_df["student_id"].str.contains(search_query, case=False, na=False) |
            filtered_df["student_name"].str.contains(search_query, case=False, na=False)
        ]
        
    # If user filtered specific sessions, display base columns + only selected sessions + stats
    if session_filter:
        base_cols = ["student_id", "student_name", "group", "section"]
        cols_to_show = [c for c in base_cols if c in filtered_df.columns] + session_filter + ["إجمالي الحضور", "نسبة الحضور %"]
        filtered_df = filtered_df[[c for c in cols_to_show if c in filtered_df.columns]]
        
    st.dataframe(filtered_df, use_container_width=True, height=380)

# ----------------- TAB 2: SINGLE SESSION QUICK VIEW -----------------
with tab2:
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    session_options = {s["session_name"]: s for s in results}
    selected_session_name = st.selectbox("📌 اختر الجلسة المراد فحصها (محاضرة أو سكشن):", list(session_options.keys()))
    s_curr = session_options[selected_session_name]
    
    cs_students = s_curr.get("cross_section_students", [])
    cs_count = len(cs_students)
    is_sec = s_curr.get("metadata", {}).get("type") == "Section" and s_curr.get("metadata", {}).get("section_assigned") != "General"

    # 3 Stat Cards (Removed visitors/unregistered per user request)
    sk1, sk2, sk3 = st.columns(3)
    with sk1:
        footer_text = f"{s_curr['target_present']} مقيد بالسكشن + {cs_count} سكشن بديل" if (is_sec and cs_count > 0) else "طالب حضر الجلسة"
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-header">
                <span class="stat-label">عدد الحاضرين</span>
                <div class="stat-icon" style="background: #ECFDF5; color: #047857;">✅</div>
            </div>
            <div class="stat-value" style="color: #047857;">{s_curr['target_present']}</div>
            <div class="stat-footer">{footer_text}</div>
        </div>
        """, unsafe_allow_html=True)
    with sk2:
        absent_footer = "طالب غائب من السكشن المستهدف" if is_sec else "طالب لم يحضر"
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-header">
                <span class="stat-label">عدد الغائبين</span>
                <div class="stat-icon" style="background: #FEF2F2; color: #B91C1C;">❌</div>
            </div>
            <div class="stat-value" style="color: #B91C1C;">{s_curr['target_absent']}</div>
            <div class="stat-footer">{absent_footer}</div>
        </div>
        """, unsafe_allow_html=True)
    with sk3:
        st.markdown(f"""
        <div class="stat-card">
            <div class="stat-header">
                <span class="stat-label">نسبة الحضور</span>
                <div class="stat-icon" style="background: #EEF2FF; color: #4338CA;">📊</div>
            </div>
            <div class="stat-value" style="color: #4338CA;">{s_curr['attendance_rate']}%</div>
            <div class="stat-footer">نسبة التزام الطلاب في الجلسة</div>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("<div style='height: 20px;'></div>", unsafe_allow_html=True)

    if is_sec and cs_count > 0:
        st.info(f"💡 تنبيه: حضر في هذا السكشن {cs_count} طالب/طلاب مقيدين بسكاشن بديلة أخرى، تم توضيحهم في تبويب منفصل أدناه.")
    
    # Donut Chart & Lists
    sc1, sc2 = st.columns([1, 2])
    with sc1:
        fig_donut = go.Figure(data=[go.Pie(
            labels=["حاضر", "غائب"],
            values=[s_curr['target_present'], s_curr['target_absent']],
            hole=.6,
            marker_colors=["#10B981", "#EF4444"],
            textinfo='percent+label',
            textfont=dict(color="#0F172A", size=13, family="Tajawal, sans-serif")
        )])
        fig_donut.update_layout(
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF",
            showlegend=False,
            margin=dict(l=10, r=10, t=10, b=10),
            height=280
        )
        st.plotly_chart(fig_donut, use_container_width=True)
        
    with sc2:
        if is_sec and cs_count > 0:
            tab_p, tab_a, tab_cs = st.tabs([
                f"✅ الطلاب الحاضرون ({s_curr['target_present']})",
                f"❌ الطلاب الغائبون ({s_curr['target_absent']})",
                f"🔄 سكشن بديل ({cs_count})"
            ])
            with tab_p:
                target_present = s_curr.get("target_present_students", s_curr["present_students"])
                st.dataframe(pd.DataFrame(target_present), use_container_width=True, height=250)
            with tab_a:
                st.dataframe(pd.DataFrame(s_curr["absent_students"]), use_container_width=True, height=250)
            with tab_cs:
                st.dataframe(pd.DataFrame(cs_students), use_container_width=True, height=250)
        else:
            tab_p, tab_a = st.tabs([
                f"✅ الطلاب الحاضرون ({s_curr['target_present']})",
                f"❌ الطلاب الغائبون ({s_curr['target_absent']})"
            ])
            with tab_p:
                st.dataframe(pd.DataFrame(s_curr["present_students"]), use_container_width=True, height=250)
            with tab_a:
                st.dataframe(pd.DataFrame(s_curr["absent_students"]), use_container_width=True, height=250)

# ----------------- TAB 3: CROSS-SECTION ATTENDANCE -----------------
with tab3:
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    all_cs = cum_data.get("all_cross_sections", [])
    if all_cs:
        cs_df = pd.DataFrame(all_cs)
        col_c1, col_c2 = st.columns([1, 2])
        with col_c1:
            cs_summary = cs_df.groupby("attended_section")["student_id"].count().reset_index()
            cs_summary.columns = ["السكشن البديل", "عدد الطلاب"]
            fig_cs = px.bar(
                cs_summary,
                x="السكشن البديل",
                y="عدد الطلاب",
                text="عدد الطلاب",
                color="عدد الطلاب",
                color_continuous_scale=["#0D9488", "#115E59"]
            )
            fig_cs.update_traces(
                textposition='outside',
                textfont=dict(color="#0F172A", size=12)
            )
            fig_cs.update_layout(
                paper_bgcolor="#FFFFFF",
                plot_bgcolor="#FFFFFF",
                font=dict(color="#0F172A", family="Tajawal, sans-serif"),
                coloraxis_showscale=False,
                margin=dict(l=10, r=10, t=25, b=10),
                yaxis=dict(gridcolor="#E2E8F0", tickfont=dict(color="#0F172A")),
                xaxis=dict(gridcolor="#E2E8F0", tickfont=dict(color="#0F172A"))
            )
            st.plotly_chart(fig_cs, use_container_width=True)
            
        with col_c2:
            st.markdown("##### 📝 قائمة الطلاب الذين حضروا في سكشن بديل:")
            st.dataframe(cs_df, use_container_width=True, height=300)
    else:
        st.success("🎉 التزم جميع الطلاب بمواعيد سكاشنهم المخصصة.")

# ----------------- TAB 4: STUDENT 360 PROFILE -----------------
with tab4:
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    if not matrix_df.empty:
        student_list = (matrix_df["student_id"] + " - " + matrix_df["student_name"]).tolist()
        selected_std = st.selectbox("🔍 ابحث عن طالب بالاسم أو الكود:", student_list)
        std_id = selected_std.split(" - ")[0].strip()
        std_row = matrix_df[matrix_df["student_id"] == std_id].iloc[0]
        
        sp1, sp2, sp3, sp4 = st.columns(4)
        with sp1:
            st.markdown(f"**الاسم:** {std_row.get('student_name', '')}")
            st.markdown(f"**الكود (ID):** {std_row.get('student_id', '')}")
        with sp2:
            st.markdown(f"**المجموعة:** {std_row.get('group', '')}")
            st.markdown(f"**السكشن:** {std_row.get('section', '')}")
        with sp3:
            st.metric("إجمالي الحضور", f"{std_row.get('إجمالي الحضور', 0)} / {std_row.get('إجمالي الجلسات', 0)}")
        with sp4:
            att_rate = std_row.get('نسبة الحضور %', 0)
            st.metric("نسبة الحضور", f"{att_rate}%")

        st.markdown("##### 📅 السجل الأسبوعي المفصل للجلسات:")
        session_names = cum_data.get("session_names", [])
        show_all = st.checkbox("عرض جميع الجلسات (بما فيها سكاشن الدفعة الأخرى غير المطالب بها)", value=False)
        records = []
        for sname in session_names:
            status = std_row.get(sname, "-")
            if status == "-":
                if show_all:
                    records.append({
                        "الجلسة": sname,
                        "الحالة": "⚪ غير مطالب (سكشن آخر)"
                    })
            elif status == "حاضر":
                records.append({
                    "الجلسة": sname,
                    "الحالة": "✅ حاضر"
                })
            else: # "غائب"
                records.append({
                    "الجلسة": sname,
                    "الحالة": "❌ غائب"
                })
        st.dataframe(pd.DataFrame(records), use_container_width=True)

# ----------------- TAB 5: EXPORTS (EXCEL & PDF) -----------------
with tab5:
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    exp1, exp2 = st.columns(2)
    with exp1:
        st.markdown("""
        <div class="glass-container" style="text-align: center; padding: 2rem;">
            <div style="font-size: 2.8rem; margin-bottom: 0.5rem;">📊</div>
            <div style="font-size: 1.25rem; font-weight: 700; color: #0F172A;">مصفوفة الحضور إكسيل (.xlsx)</div>
            <p style="color: #475569; font-size: 0.9rem; margin: 0.5rem auto 1.5rem auto; max-width: 400px;">
                ملف Excel بتنسيق شرطي ملون (أخضر للحاضرين / أحمر للغائبين) جاهز فوراً للطباعة والفرز.
            </p>
        """, unsafe_allow_html=True)
        excel_bytes = export_attendance_matrix_excel(matrix_df, cum_data.get("session_names", []))
        st.download_button(
            label="تحميل ملف الإكسيل الشامل (.xlsx)",
            data=excel_bytes,
            file_name="Attendance_Matrix_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
        st.markdown("</div>", unsafe_allow_html=True)
        
    with exp2:
        st.markdown("""
        <div class="glass-container" style="text-align: center; padding: 2rem;">
            <div style="font-size: 2.8rem; margin-bottom: 0.5rem;">📄</div>
            <div style="font-size: 1.25rem; font-weight: 700; color: #0F172A;">تقرير تنفيذي رسمي (.pdf)</div>
            <p style="color: #475569; font-size: 0.9rem; margin: 0.5rem auto 1.5rem auto; max-width: 400px;">
                تقرير PDF معتمد ومصمم باللغة العربية، يضم المؤشرات الرئيسية وكشف الغياب وحالات السكاشن البديلة.
            </p>
        """, unsafe_allow_html=True)
        
        pdf_bytes = generate_pdf_summary_report(
            course_title="مقرر الفصل الدراسي الحالي",
            overall_stats=cum_data,
            top_absent_df=matrix_df,
            cross_sections=cum_data.get("all_cross_sections", [])
        )
        st.download_button(
            label="تحميل التقرير التنفيذي (.pdf)",
            data=pdf_bytes,
            file_name="Attendance_Executive_Summary.pdf",
            mime="application/pdf",
            use_container_width=True,
            type="primary"
        )
        st.markdown("</div>", unsafe_allow_html=True)
