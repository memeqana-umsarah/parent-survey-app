import os
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
import plotly.express as px

# دعم العربي في PDF بشكل اختياري
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    ARABIC_SUPPORT = True
except ImportError:
    ARABIC_SUPPORT = False

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# =========================
# إعدادات الصفحة
# =========================
st.set_page_config(
    page_title="استبانة أولياء الأمور 2025-2026",
    page_icon="📝",
    layout="wide"
)

# =========================
# الملفات
# =========================
STUDENTS_FILE = "students.xlsx"
RESULTS_FILE = "survey_results.xlsx"
SCHOOL_TOTALS_FILE = "school_totals.xlsx"
LOGO_FILE = "logo.png"
BANNER_FILE = "banner.png"

# حطي واحد من هالخطوط داخل نفس مجلد المشروع
ARABIC_FONT_CANDIDATES = [
    "Amiri-Regular.ttf",
    "NotoNaskhArabic-Regular.ttf",
    "Cairo-Regular.ttf"
    "Simpo.ttf"
]

# =========================
# بيانات المشرف
# =========================
APP_TITLE = "استبانة أولياء الأمور 2025-2026"
SCHOOL_NAME = "مدارس الكلية العلمية الإسلامية"
ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "mmmm"

# =========================
# الهوية البصرية
# =========================
PRIMARY_COLOR = "#0B3D91"
SECONDARY_COLOR = "#2563EB"
ACCENT_COLOR = "#D4A017"
BACKGROUND_COLOR = "#F3F6FB"
CARD_COLOR = "#FFFFFF"
TEXT_COLOR = "#1F2937"

# =========================
# قوالب الاستبانات
# =========================
SURVEY_TEMPLATES = {
    "E1": {
        "المحور الأول: البيئة المدرسية": [
            "هل تشعر أن المدرسة توفر بيئة آمنة لابنك/ابنتك؟",
            "هل المدرسة نظيفة ومرتبة بشكل مناسب؟",
            "هل يتم الاهتمام براحة الطلبة داخل المدرسة؟",
            "هل المدرسة تتواصل بوضوح عند وجود ملاحظات؟",
            "هل ترى أن مرافق المدرسة مناسبة للطلبة؟",
        ],
        "المحور الثاني: المستوى الأكاديمي": [
            "هل مستوى التدريس مناسب لابنك/ابنتك؟",
            "هل الواجبات المنزلية مناسبة من حيث الكم والمحتوى؟",
            "هل يتم متابعة تقدم الطالب بشكل مستمر؟",
            "هل تشعر بوجود تحسن في المستوى الدراسي؟",
            "هل الشرح داخل الصف واضح ومفيد؟",
        ],
        "المحور الثالث: المعلمون": [
            "هل المعلمون يتعاملون باحترام مع الطلبة؟",
            "هل المعلمون متعاونون مع أولياء الأمور؟",
            "هل يوضح المعلمون نقاط القوة والضعف للطالب؟",
            "هل المعلمون يراعون الفروق الفردية بين الطلبة؟",
            "هل ترى أن المعلمين ملتزمون بمهامهم؟",
        ],
        "المحور الرابع: التواصل مع ولي الأمر": [
            "هل إدارة المدرسة تتواصل معك بشكل كافٍ؟",
            "هل تستقبل الملاحظات الأكاديمية والسلوكية بانتظام؟",
            "هل المدرسة تستجيب لاستفساراتك بسرعة؟",
            "هل قنوات التواصل مع المدرسة فعالة؟",
            "هل تشعر أن رأيك كولي أمر محل اهتمام؟",
        ],
        "المحور الخامس: الرضا العام": [
            "هل أنت راضٍ بشكل عام عن المدرسة؟",
            "هل تنصح أولياء الأمور الآخرين بهذه المدرسة؟",
            "هل تعتقد أن المدرسة تلبي احتياجات الطالب؟",
            "هل ترى أن مستوى الانضباط في المدرسة جيد؟",
            "هل ترغب في استمرار ابنك/ابنتك في هذه المدرسة؟",
        ]
    },

    "E2": {
        "المحور الأول: البيئة المدرسية": [
            "هل البيئة المدرسية مناسبة لطلاب المرحلة؟",
            "هل الانضباط المدرسي جيد داخل المدرسة؟",
            "هل تشعر أن مرافق المدرسة مناسبة؟",
            "هل مستوى النظافة في المدرسة جيد؟",
            "هل الأجواء العامة مشجعة للتعلم؟",
        ],
        "المحور الثاني: التحصيل الأكاديمي": [
            "هل المنهاج مناسب لمستوى الطالب؟",
            "هل الشرح داخل الصف واضح؟",
            "هل يتم متابعة الواجبات بشكل مناسب؟",
            "هل يتم إعلامكم بمستوى الطالب باستمرار؟",
            "هل يوجد تحسن في أداء الطالب؟",
        ],
        "المحور الثالث: الكادر التعليمي": [
            "هل المعلمون ملتزمون بواجباتهم؟",
            "هل يتعامل المعلمون باحترام مع الطلبة؟",
            "هل يوجد تعاون جيد بين المعلمين وأولياء الأمور؟",
            "هل يتم مراعاة الفروق الفردية؟",
            "هل يتم توجيه الطالب أكاديميًا بشكل مناسب؟",
        ],
        "المحور الرابع: التواصل": [
            "هل المدرسة تتواصل معكم عند الحاجة؟",
            "هل يتم الرد على استفساراتكم بسرعة؟",
            "هل قنوات التواصل واضحة؟",
            "هل يتم تزويدكم بالملاحظات الأكاديمية بانتظام؟",
            "هل تشعر أن رأيك مهم بالنسبة للمدرسة؟",
        ],
        "المحور الخامس: الرضا العام": [
            "هل أنت راضٍ عن المدرسة بشكل عام؟",
            "هل توصي بهذه المدرسة للآخرين؟",
            "هل ترى أن المدرسة تلبي احتياجات الطالب؟",
            "هل ترغب باستمرار ابنك/ابنتك فيها؟",
            "هل مستوى الإدارة المدرسية مرضٍ؟",
        ]
    }
}

answer_options = ["موافق جدًا", "موافق", "محايد", "غير موافق", "غير موافق جدًا"]

answer_score_map = {
    "موافق جدًا": 5,
    "موافق": 4,
    "محايد": 3,
    "غير موافق": 2,
    "غير موافق جدًا": 1
}

# =========================
# CSS
# =========================
st.markdown(f"""
<style>
html, body, [class*="css"] {{
    direction: rtl !important;
    text-align: right !important;
    background-color: {BACKGROUND_COLOR};
    color: {TEXT_COLOR};
    font-size: 19px !important;
}}

body {{
    unicode-bidi: bidi-override;
}}

.block-container {{
    padding-top: 1rem;
    padding-bottom: 2rem;
    max-width: 1280px;
}}

.main-title {{
    text-align: right !important;
    font-size: 40px;
    font-weight: 900;
    color: {PRIMARY_COLOR};
    margin-bottom: 6px;
    line-height: 1.5;
}}

.sub-title {{
    text-align: right !important;
    font-size: 24px;
    font-weight: 700;
    color: {SECONDARY_COLOR};
    margin-bottom: 18px;
    line-height: 1.5;
}}

.info-box {{
    background-color: {CARD_COLOR};
    border-right: 7px solid {PRIMARY_COLOR};
    padding: 18px 22px;
    border-radius: 16px;
    box-shadow: 0 3px 12px rgba(0,0,0,0.06);
    margin-bottom: 18px;
    font-size: 19px;
    font-weight: 600;
    text-align: right !important;
}}

.section-card {{
    background-color: {CARD_COLOR};
    padding: 24px;
    border-radius: 18px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.07);
    margin-bottom: 22px;
    text-align: right !important;
}}

.axis-title {{
    color: #0A2F6B;
    font-size: 32px;
    font-weight: 900;
    margin-bottom: 10px;
    line-height: 1.6;
    text-align: right !important;
}}

.stTextInput input, .stTextArea textarea {{
    direction: rtl !important;
    text-align: right !important;
    border-radius: 10px !important;
    font-size: 18px !important;
}}

div[data-baseweb="select"] > div {{
    direction: rtl !important;
    text-align: right !important;
    border-radius: 10px !important;
    font-size: 18px !important;
}}

.stRadio > div {{
    direction: rtl !important;
    text-align: right !important;
}}

.stRadio label, .stRadio div {{
    direction: rtl !important;
    text-align: right !important;
    font-size: 18px !important;
    font-weight: 600 !important;
}}

.stMarkdown, .stAlert, .stCaption, label, p, div, span {{
    text-align: right !important;
}}

.stButton > button {{
    width: 100%;
    border-radius: 12px;
    font-weight: 800;
    border: none;
    padding: 0.75rem 1rem;
    font-size: 18px;
}}

div[data-testid="stMetric"] {{
    background-color: {CARD_COLOR};
    border-radius: 16px;
    padding: 12px;
    box-shadow: 0 2px 10px rgba(0,0,0,0.05);
}}

div[data-testid="stMetric"] label {{
    font-size: 18px !important;
    font-weight: 700 !important;
    text-align: right !important;
}}

div[data-testid="stMetricValue"] {{
    font-size: 28px !important;
    font-weight: 900 !important;
    text-align: right !important;
}}

[data-testid="stDataFrame"] {{
    background-color: {CARD_COLOR};
    border-radius: 14px;
    padding: 8px;
}}

[data-testid="stDataFrame"] * {{
    direction: rtl !important;
    text-align: right !important;
}}

h1, h2, h3, h4, h5, h6 {{
    font-weight: 900 !important;
    color: {PRIMARY_COLOR} !important;
    text-align: right !important;
}}

img {{
    border-radius: 14px;
}}

.element-container, .stPlotlyChart {{
    direction: rtl !important;
}}
</style>
""", unsafe_allow_html=True)

# =========================
# دوال مساعدة
# =========================
def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()

def score_to_percentage(score):
    if pd.isna(score):
        return 0.0
    return round((float(score) / 5) * 100, 2)

def ar_text(text):
    if pd.isna(text):
        text = ""
    text = str(text)

    if ARABIC_SUPPORT:
        try:
            reshaped = arabic_reshaper.reshape(text)
            return get_display(reshaped)
        except Exception:
            return text

    return text

def register_arabic_font():
    for font_file in ARABIC_FONT_CANDIDATES:
        font_path = os.path.join(os.getcwd(), font_file)
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont("ArabicFont", font_path))
                return "ArabicFont"
            except Exception:
                continue
    return "Helvetica"

def get_student_survey_type(student):
    survey_type = normalize_text(student.get("survey_type", "E1"))
    if survey_type not in SURVEY_TEMPLATES:
        survey_type = list(SURVEY_TEMPLATES.keys())[0]
    return survey_type

def get_survey_questions_by_student(student):
    survey_type = get_student_survey_type(student)
    return SURVEY_TEMPLATES[survey_type]

def get_max_questions_count():
    max_count = 0
    for template in SURVEY_TEMPLATES.values():
        count = sum(len(qs) for qs in template.values())
        max_count = max(max_count, count)
    return max_count

def get_max_axes_count():
    max_axes = 0
    for template in SURVEY_TEMPLATES.values():
        max_axes = max(max_axes, len(template.keys()))
    return max_axes

def render_bar_chart(df, x_col, y_col, title, color_col=None):
    if df.empty:
        return

    fig = px.bar(
        df,
        x=x_col,
        y=y_col,
        color=color_col if color_col else None,
        text=y_col,
        title=title,
        barmode="group"
    )

    fig.update_traces(textposition="outside", cliponaxis=False)

    fig.update_layout(
        height=520,
        xaxis_title="",
        yaxis_title=y_col,
        font=dict(size=16),
        title=dict(font=dict(size=22)),
        xaxis=dict(tickfont=dict(size=15)),
        yaxis=dict(tickfont=dict(size=15)),
        plot_bgcolor="white",
        paper_bgcolor="white",
        margin=dict(l=20, r=20, t=60, b=80),
    )

    st.plotly_chart(fig, use_container_width=True)

def build_pdf_report_bytes(filtered_df, axis_summary_df, question_summary_df, school_summary_df):
    output = BytesIO()
    font_name = register_arabic_font()

    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        rightMargin=1.0 * cm,
        leftMargin=1.0 * cm,
        topMargin=1.0 * cm,
        bottomMargin=1.0 * cm
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        name="ArabicTitle",
        parent=styles["Title"],
        fontName=font_name,
        fontSize=18,
        leading=24,
        alignment=TA_CENTER
    )

    heading_style = ParagraphStyle(
        name="ArabicHeading",
        parent=styles["Heading2"],
        fontName=font_name,
        fontSize=13,
        leading=18,
        alignment=TA_RIGHT
    )

    normal_style = ParagraphStyle(
        name="ArabicNormal",
        parent=styles["BodyText"],
        fontName=font_name,
        fontSize=9,
        leading=12,
        alignment=TA_RIGHT
    )

    elements = []

    banner_path = os.path.join(os.getcwd(), BANNER_FILE)
    if os.path.exists(banner_path):
        try:
            banner = Image(banner_path, width=26 * cm, height=4.0 * cm)
            elements.append(banner)
            elements.append(Spacer(1, 0.3 * cm))
        except Exception:
            pass

    elements.append(Paragraph(ar_text("تقرير نتائج استبانة أولياء الأمور"), title_style))
    elements.append(Spacer(1, 0.3 * cm))
    elements.append(
        Paragraph(ar_text(f"تاريخ التصدير: {datetime.now().strftime('%Y-%m-%d %H:%M')}"), normal_style)
    )
    elements.append(Spacer(1, 0.4 * cm))

    def shorten_text(value, max_len=35):
        if pd.isna(value):
            return ""
        value = str(value)
        if len(value) > max_len:
            return value[:max_len] + "..."
        return value

    def make_table_from_df(df, title, selected_cols=None, max_rows=20):
        if df is None or df.empty:
            return

        work_df = df.copy()

        if selected_cols:
            available_cols = [col for col in selected_cols if col in work_df.columns]
            if not available_cols:
                return
            work_df = work_df[available_cols]

        work_df = work_df.head(max_rows).fillna("")

        for col in work_df.columns:
            work_df[col] = work_df[col].apply(lambda x: shorten_text(x, 40))

        elements.append(Paragraph(ar_text(title), heading_style))
        elements.append(Spacer(1, 0.2 * cm))

        headers = [Paragraph(ar_text(col), normal_style) for col in work_df.columns]
        data = [headers]

        for _, row in work_df.iterrows():
            row_cells = [Paragraph(ar_text(val), normal_style) for val in row]
            data.append(row_cells)

        page_width = 27 * cm
        num_cols = len(work_df.columns)
        col_width = page_width / max(num_cols, 1)
        col_widths = [col_width] * num_cols

        tbl = Table(data, colWidths=col_widths, repeatRows=1)

        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0B3D91")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, -1), font_name),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.HexColor("#F8FAFC")]),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ]))

        elements.append(tbl)
        elements.append(Spacer(1, 0.4 * cm))

    make_table_from_df(
        axis_summary_df,
        "ملخص متوسطات المحاور",
        selected_cols=["المحور", "المتوسط", "النسبة المئوية"],
        max_rows=15
    )

    make_table_from_df(
        school_summary_df,
        "ملخص المدارس",
        selected_cols=[
            "اسم المدرسة",
            "نوع الاستبانة",
            "عدد الاستجابات",
            "عدد الطلبة الكلي",
            "نسبة الاستجابة",
            "المتوسط الكلي",
            "النسبة المئوية"
        ],
        max_rows=20
    )

    make_table_from_df(
        question_summary_df,
        "ملخص الفقرات",
        selected_cols=["رقم الفقرة", "المحور", "الفقرة", "المتوسط", "النسبة المئوية"],
        max_rows=15
    )

    raw_pdf_cols = [
        "student_id",
        "student_name",
        "school",
        "survey_type",
        "respondent_type",
        "overall_avg",
        "overall_pct",
        "contact_phone"
    ]

    make_table_from_df(
        filtered_df,
        "النتائج الخام المختصرة",
        selected_cols=raw_pdf_cols,
        max_rows=20
    )

    doc.build(elements)
    output.seek(0)
    return output.getvalue()

# =========================
# Session State
# =========================
def init_session():
    defaults = {
        "page": "home",
        "logged_in_parent": False,
        "logged_in_admin": False,
        "student_data": None,
        "current_axis": 0,
        "answers": {},
        "notes": "",
        "respondent_type": "",
        "father_job": "",
        "mother_job": "",
        "contact_phone": "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

def reset_parent_session():
    st.session_state.logged_in_parent = False
    st.session_state.student_data = None
    st.session_state.current_axis = 0
    st.session_state.answers = {}
    st.session_state.notes = ""
    st.session_state.respondent_type = ""
    st.session_state.father_job = ""
    st.session_state.mother_job = ""
    st.session_state.contact_phone = ""
    st.session_state.page = "home"

def reset_admin_session():
    st.session_state.logged_in_admin = False
    st.session_state.page = "home"

# =========================
# Header
# =========================
def render_header():
    banner_path = os.path.join(os.getcwd(), BANNER_FILE)
    logo_path = os.path.join(os.getcwd(), LOGO_FILE)

    if os.path.exists(banner_path):
        st.image(banner_path, use_container_width=True)

    right_col, left_col = st.columns([1, 4])

    with right_col:
        if os.path.exists(logo_path):
            st.image(logo_path, width=150)

    with left_col:
        st.markdown(f'<div class="main-title">{APP_TITLE}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="sub-title">{SCHOOL_NAME}</div>', unsafe_allow_html=True)

# =========================
# تحميل الملفات
# =========================
def ensure_results_file_exists():
    if not os.path.exists(RESULTS_FILE):
        columns = [
            "student_id",
            "student_name",
            "grade",
            "school",
            "survey_type",
            "respondent_type",
            "father_job",
            "mother_job",
            "contact_phone",
            "overall_avg",
            "overall_pct",
            "notes",
            "timestamp"
        ]

        max_axes = get_max_axes_count()
        for i in range(1, max_axes + 1):
            columns.append(f"axis{i}_name")
            columns.append(f"axis{i}_avg")
            columns.append(f"axis{i}_pct")

        max_questions = get_max_questions_count()
        for i in range(1, max_questions + 1):
            columns.extend([
                f"Q{i}",
                f"Q{i}_text",
                f"Q{i}_axis"
            ])

        pd.DataFrame(columns=columns).to_excel(RESULTS_FILE, index=False)

def load_students():
    if not os.path.exists(STUDENTS_FILE):
        return None, f"ملف الطلاب غير موجود: {STUDENTS_FILE}"

    try:
        df = pd.read_excel(STUDENTS_FILE)

        required_cols = ["student_id", "password", "student_name", "grade", "school", "survey_type"]
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            return None, f"الأعمدة التالية مفقودة في ملف الطلاب: {missing}"

        df["student_id"] = pd.to_numeric(df["student_id"], errors="coerce").fillna(0).astype(int).astype(str)
        df["password"] = pd.to_numeric(df["password"], errors="coerce").fillna(0).astype(int).astype(str)

        for col in ["student_name", "grade", "school", "survey_type"]:
            df[col] = df[col].astype(str).str.strip()

        return df, None
    except Exception as e:
        return None, f"حدث خطأ أثناء قراءة ملف الطلاب: {e}"

def load_results():
    ensure_results_file_exists()
    try:
        df = pd.read_excel(RESULTS_FILE)
        return df, None
    except Exception as e:
        return None, f"حدث خطأ أثناء قراءة ملف النتائج: {e}"

def load_school_totals():
    if not os.path.exists(SCHOOL_TOTALS_FILE):
        return None, "ملف أعداد الطلبة الكلي غير موجود"

    try:
        df = pd.read_excel(SCHOOL_TOTALS_FILE)

        required_cols = ["school", "total_students"]
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            return None, f"الأعمدة التالية مفقودة في ملف أعداد الطلبة: {missing}"

        df["school"] = df["school"].astype(str).str.strip()
        df["total_students"] = pd.to_numeric(df["total_students"], errors="coerce").fillna(0)

        return df, None
    except Exception as e:
        return None, f"حدث خطأ أثناء قراءة ملف أعداد الطلبة: {e}"

def student_already_submitted(student_id):
    df, error = load_results()
    if error or df is None or df.empty or "student_id" not in df.columns:
        return False

    df["student_id"] = df["student_id"].astype(str).str.strip()
    return str(student_id).strip() in df["student_id"].values

# =========================
# المتوسطات
# =========================
def get_axis_average(student, axis_name):
    current_survey_questions = get_survey_questions_by_student(student)
    questions = current_survey_questions[axis_name]

    scores = []
    for q in questions:
        answer = st.session_state.answers.get(q, "")
        if answer in answer_score_map:
            scores.append(answer_score_map[answer])

    if not scores:
        return 0.0
    return round(sum(scores) / len(scores), 2)

def get_overall_average(student):
    current_survey_questions = get_survey_questions_by_student(student)
    all_questions = [q for axis in current_survey_questions.values() for q in axis]

    scores = []
    for q in all_questions:
        answer = st.session_state.answers.get(q, "")
        if answer in answer_score_map:
            scores.append(answer_score_map[answer])

    if not scores:
        return 0.0
    return round(sum(scores) / len(scores), 2)

# =========================
# حفظ النتائج
# =========================
def save_survey():
    student = st.session_state.student_data
    student_id = normalize_text(student.get("student_id", ""))

    if student_already_submitted(student_id):
        return False, "هذا الطالب قام بتعبئة الاستبانة مسبقًا"

    current_survey_questions = get_survey_questions_by_student(student)
    current_axes_list = list(current_survey_questions.keys())
    overall_avg = get_overall_average(student)

    row_data = {
        "student_id": student.get("student_id", ""),
        "student_name": student.get("student_name", ""),
        "grade": student.get("grade", ""),
        "school": student.get("school", ""),
        "survey_type": get_student_survey_type(student),
        "respondent_type": st.session_state.respondent_type,
        "father_job": st.session_state.father_job.strip(),
        "mother_job": st.session_state.mother_job.strip(),
        "contact_phone": st.session_state.contact_phone.strip(),
        "overall_avg": overall_avg,
        "overall_pct": score_to_percentage(overall_avg),
        "notes": st.session_state.notes.strip(),
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    for i, axis_name in enumerate(current_axes_list, start=1):
        axis_avg = get_axis_average(student, axis_name)
        row_data[f"axis{i}_name"] = axis_name
        row_data[f"axis{i}_avg"] = axis_avg
        row_data[f"axis{i}_pct"] = score_to_percentage(axis_avg)

    q_num = 1
    for axis_name, questions in current_survey_questions.items():
        for q in questions:
            answer_text = st.session_state.answers.get(q, "")
            row_data[f"Q{q_num}"] = answer_score_map.get(answer_text, "")
            row_data[f"Q{q_num}_text"] = q
            row_data[f"Q{q_num}_axis"] = axis_name
            q_num += 1

    try:
        ensure_results_file_exists()
        existing_df = pd.read_excel(RESULTS_FILE)
        updated_df = pd.concat([existing_df, pd.DataFrame([row_data])], ignore_index=True)
        updated_df.to_excel(RESULTS_FILE, index=False)
        return True, "تم حفظ الاستبانة بنجاح"
    except Exception as e:
        return False, f"حدث خطأ أثناء حفظ النتائج: {e}"

# =========================
# التحليل
# =========================
def dataframe_to_excel_bytes(df_dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in df_dict.items():
            safe_name = str(sheet_name)[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
    output.seek(0)
    return output.getvalue()

def build_question_summary(filtered_df):
    if filtered_df.empty:
        return pd.DataFrame(columns=["رقم الفقرة", "المحور", "الفقرة", "المتوسط", "النسبة المئوية"])

    rows = []
    question_cols = [col for col in filtered_df.columns if col.startswith("Q") and "_text" not in col and "_axis" not in col]

    question_numbers = []
    for col in question_cols:
        if col[1:].isdigit():
            question_numbers.append(int(col[1:]))

    for q_num in sorted(question_numbers):
        col_name = f"Q{q_num}"
        q_text_col = f"Q{q_num}_text"
        q_axis_col = f"Q{q_num}_axis"

        valid_rows = filtered_df[filtered_df[q_text_col].notna()] if q_text_col in filtered_df.columns else pd.DataFrame()
        if valid_rows.empty:
            continue

        question_text = valid_rows[q_text_col].dropna().astype(str).iloc[0] if q_text_col in valid_rows.columns else ""
        axis_name = valid_rows[q_axis_col].dropna().astype(str).iloc[0] if q_axis_col in valid_rows.columns else ""
        avg_val = round(pd.to_numeric(valid_rows[col_name], errors="coerce").mean(), 2) if col_name in valid_rows.columns else 0

        rows.append({
            "رقم الفقرة": q_num,
            "المحور": axis_name,
            "الفقرة": question_text,
            "المتوسط": 0 if pd.isna(avg_val) else avg_val,
            "النسبة المئوية": score_to_percentage(avg_val)
        })

    return pd.DataFrame(rows)

def build_axis_summary(filtered_df):
    if filtered_df.empty:
        return pd.DataFrame(columns=["المحور", "المتوسط", "النسبة المئوية"])

    rows = []
    axis_avg_cols = sorted(
        [col for col in filtered_df.columns if col.startswith("axis") and col.endswith("_avg")],
        key=lambda x: int(x.replace("axis", "").replace("_avg", ""))
    )

    for avg_col in axis_avg_cols:
        axis_num = avg_col.replace("axis", "").replace("_avg", "")
        name_col = f"axis{axis_num}_name"

        valid_rows = filtered_df[filtered_df[name_col].notna()] if name_col in filtered_df.columns else pd.DataFrame()
        if valid_rows.empty:
            continue

        axis_name = valid_rows[name_col].dropna().astype(str).iloc[0]
        avg_val = round(pd.to_numeric(valid_rows[avg_col], errors="coerce").mean(), 2)

        rows.append({
            "المحور": axis_name,
            "المتوسط": 0 if pd.isna(avg_val) else avg_val,
            "النسبة المئوية": score_to_percentage(avg_val)
        })

    return pd.DataFrame(rows)

def build_school_summary(df):
    if "school" not in df.columns or df.empty:
        return pd.DataFrame(columns=[
            "اسم المدرسة",
            "نوع الاستبانة",
            "عدد الاستجابات",
            "عدد الطلبة الكلي",
            "نسبة الاستجابة",
            "المتوسط الكلي",
            "النسبة المئوية"
        ])

    work_df = df.copy()
    work_df["overall_avg"] = pd.to_numeric(work_df["overall_avg"], errors="coerce")

    group_cols = ["school"]
    if "survey_type" in work_df.columns:
        group_cols.append("survey_type")

    summary = (
        work_df.groupby(group_cols, dropna=False)
        .agg(
            عدد_الاستجابات=("student_id", "count"),
            المتوسط_الكلي=("overall_avg", "mean")
        )
        .reset_index()
    )

    summary = summary.rename(columns={
        "school": "اسم المدرسة",
        "survey_type": "نوع الاستبانة",
        "عدد_الاستجابات": "عدد الاستجابات",
        "المتوسط_الكلي": "المتوسط الكلي"
    })

    summary["المتوسط الكلي"] = summary["المتوسط الكلي"].round(2)
    summary["النسبة المئوية"] = summary["المتوسط الكلي"].apply(score_to_percentage)

    if "نوع الاستبانة" not in summary.columns:
        summary["نوع الاستبانة"] = ""

    totals_df, _ = load_school_totals()

    if totals_df is not None and not totals_df.empty:
        totals_df = totals_df.rename(columns={
            "school": "اسم المدرسة",
            "total_students": "عدد الطلبة الكلي"
        })

        summary = summary.merge(totals_df, on="اسم المدرسة", how="left")
        summary["عدد الطلبة الكلي"] = pd.to_numeric(summary["عدد الطلبة الكلي"], errors="coerce").fillna(0)

        summary["نسبة الاستجابة"] = summary.apply(
            lambda row: round((row["عدد الاستجابات"] / row["عدد الطلبة الكلي"]) * 100, 2)
            if row["عدد الطلبة الكلي"] > 0 else 0,
            axis=1
        )
    else:
        summary["عدد الطلبة الكلي"] = 0
        summary["نسبة الاستجابة"] = 0

    return summary[[
        "اسم المدرسة",
        "نوع الاستبانة",
        "عدد الاستجابات",
        "عدد الطلبة الكلي",
        "نسبة الاستجابة",
        "المتوسط الكلي",
        "النسبة المئوية"
    ]]

# =========================
# الصفحات
# =========================
def render_home():
    render_header()

    st.markdown(
        '<div class="section-card" style="text-align:right;">اختر نوع الدخول المناسب</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:
        st.markdown(f"""
        <div class="section-card" style="border-top: 5px solid {PRIMARY_COLOR}; min-height: 180px; text-align:right;">
            <div class="axis-title">دخول ولي الأمر</div>
            <div>لتعبئة الاستبانة باستخدام رقم الطالب والباسوورد.</div>
        </div>
        """, unsafe_allow_html=True)

        if st.button("فتح صفحة ولي الأمر", key="parent_btn", use_container_width=True):
            st.session_state.page = "parent_login"
            st.rerun()

    with col2:
        st.markdown(f"""
        <div class="section-card" style="border-top: 5px solid {ACCENT_COLOR}; min-height: 180px; text-align:right;">
            <div class="axis-title">دخول المشرف / الإدارة</div>
            <div>لعرض النتائج والتحليل وتنزيل الملفات.</div>
        </div>
        """, unsafe_allow_html=True)

        if st.button("فتح صفحة المشرف", key="admin_btn", use_container_width=True):
            st.session_state.page = "admin_login"
            st.rerun()

def render_parent_login():
    render_header()
    st.markdown(
        '<div class="section-card"><div class="axis-title">تسجيل دخول ولي الأمر</div></div>',
        unsafe_allow_html=True
    )

    with st.form("parent_login_form"):
        student_id = st.text_input("رقم الطالب")
        password = st.text_input("الباسوورد", type="password")
        submitted = st.form_submit_button("دخول")

    if st.button("العودة للرئيسية"):
        st.session_state.page = "home"
        st.rerun()

    if submitted:
        if not student_id.strip() or not password.strip():
            st.warning("يرجى إدخال رقم الطالب والباسوورد")
            return

        students_df, error = load_students()
        if error:
            st.error(error)
            return

        sid = str(student_id).strip()
        pwd = str(password).strip()

        result = students_df[
            (students_df["student_id"] == sid) &
            (students_df["password"] == pwd)
        ]

        if result.empty:
            st.error("رقم الطالب أو الباسوورد غير صحيح")
            return

        if student_already_submitted(sid):
            st.warning("هذا الطالب قام بتعبئة الاستبانة مسبقًا")
            return

        st.session_state.student_data = result.iloc[0].to_dict()
        st.session_state.logged_in_parent = True
        st.session_state.current_axis = 0
        st.session_state.answers = {}
        st.session_state.notes = ""
        st.session_state.page = "student_info"
        st.rerun()

def render_student_info_page():
    student = st.session_state.student_data
    render_header()

    st.markdown(
        '<div class="section-card"><div class="axis-title">بيانات ولي الأمر والطالب</div></div>',
        unsafe_allow_html=True
    )

    st.markdown(
        f'''
        <div class="info-box">
        <b>رقم الطالب:</b> {student.get('student_id', '')}
        <br><b>اسم الطالب:</b> {student.get('student_name', '')}
        <br><b>الصف:</b> {student.get('grade', '')}
        <br><b>المدرسة:</b> {student.get('school', '')}
        <br><b>نوع الاستبانة:</b> {get_student_survey_type(student)}
        </div>
        ''',
        unsafe_allow_html=True
    )

    respondent_options = ["الأم", "الأب", "الاثنان معًا"]
    previous_respondent = st.session_state.respondent_type
    respondent_index = respondent_options.index(previous_respondent) if previous_respondent in respondent_options else 0

    st.session_state.respondent_type = st.radio(
        "من يقوم بتعبئة الاستبانة؟",
        respondent_options,
        index=respondent_index,
        horizontal=True
    )

    st.session_state.father_job = st.text_input("عمل الأب", value=st.session_state.father_job)
    st.session_state.mother_job = st.text_input("عمل الأم", value=st.session_state.mother_job)
    st.session_state.contact_phone = st.text_input("رقم الهاتف للتواصل", value=st.session_state.contact_phone)

    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        if st.button("خروج", use_container_width=True):
            reset_parent_session()
            st.rerun()

    with col2:
        if st.button("التالي إلى الاستبانة", use_container_width=True):
            if not st.session_state.respondent_type:
                st.warning("يرجى تحديد من يقوم بتعبئة الاستبانة")
                return
            if not st.session_state.father_job.strip():
                st.warning("يرجى إدخال عمل الأب")
                return
            if not st.session_state.mother_job.strip():
                st.warning("يرجى إدخال عمل الأم")
                return
            if not st.session_state.contact_phone.strip():
                st.warning("يرجى إدخال رقم الهاتف للتواصل")
                return

            st.session_state.page = "survey"
            st.rerun()

def render_survey_page():
    student = st.session_state.student_data
    current_survey_questions = get_survey_questions_by_student(student)
    current_axes_list = list(current_survey_questions.keys())

    axis_index = st.session_state.current_axis
    axis_name = current_axes_list[axis_index]
    questions = current_survey_questions[axis_name]

    render_header()

    st.markdown(
        f'''
        <div class="info-box">
        <b>رقم الطالب:</b> {student.get('student_id', '')}
        &nbsp;&nbsp; | &nbsp;&nbsp;
        <b>الاسم:</b> {student.get('student_name', '')}
        &nbsp;&nbsp; | &nbsp;&nbsp;
        <b>الصف:</b> {student.get('grade', '')}
        &nbsp;&nbsp; | &nbsp;&nbsp;
        <b>المدرسة:</b> {student.get('school', '')}
        &nbsp;&nbsp; | &nbsp;&nbsp;
        <b>نوع الاستبانة:</b> {get_student_survey_type(student)}
        </div>
        ''',
        unsafe_allow_html=True
    )

    total_axes = len(current_axes_list)
    st.progress((axis_index + 1) / total_axes, text=f"المحور {axis_index + 1} من {total_axes}")

    st.markdown(
        f'<div class="section-card"><div class="axis-title">{axis_name}</div></div>',
        unsafe_allow_html=True
    )

    start_q_num = sum(len(current_survey_questions[a]) for a in current_axes_list[:axis_index]) + 1

    for i, q in enumerate(questions):
        q_num = start_q_num + i
        previous_answer = st.session_state.answers.get(q, None)

        answer = st.radio(
            label=f"{q_num}. {q}",
            options=answer_options,
            index=answer_options.index(previous_answer) if previous_answer in answer_options else 0,
            key=f"radio_{axis_index}_{i}",
            horizontal=True
        )

        st.session_state.answers[q] = answer

    if axis_index == total_axes - 1:
        st.session_state.notes = st.text_area(
            "ملاحظات إضافية",
            value=st.session_state.notes,
            height=120
        )

    st.divider()

    col1, col2, col3 = st.columns(3)

    with col1:
        if st.button("خروج", use_container_width=True):
            reset_parent_session()
            st.rerun()

    with col2:
        if axis_index > 0:
            if st.button("السابق", use_container_width=True):
                st.session_state.current_axis -= 1
                st.rerun()

    with col3:
        current_axis_answered = all(st.session_state.answers.get(q, "") for q in questions)

        if axis_index < total_axes - 1:
            if st.button("التالي", use_container_width=True):
                if not current_axis_answered:
                    st.warning("يرجى الإجابة على جميع أسئلة هذا المحور قبل الانتقال")
                else:
                    st.session_state.current_axis += 1
                    st.rerun()
        else:
            if st.button("حفظ الاستبانة", use_container_width=True):
                all_questions = [q for axis in current_axes_list for q in current_survey_questions[axis]]
                unanswered = [q for q in all_questions if not st.session_state.answers.get(q, "")]

                if unanswered:
                    st.warning("يرجى الإجابة على جميع الأسئلة قبل الحفظ")
                    return

                success, msg = save_survey()
                if success:
                    st.success(msg)
                    st.balloons()
                    reset_parent_session()
                    st.rerun()
                else:
                    st.error(msg)

def render_admin_login():
    render_header()
    st.markdown(
        '<div class="section-card"><div class="axis-title">دخول المشرف / الإدارة</div></div>',
        unsafe_allow_html=True
    )

    with st.form("admin_login_form"):
        username = st.text_input("اسم المستخدم")
        password = st.text_input("كلمة المرور", type="password")
        submitted = st.form_submit_button("دخول")

    if st.button("العودة للرئيسية"):
        st.session_state.page = "home"
        st.rerun()

    if submitted:
        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            st.session_state.logged_in_admin = True
            st.session_state.page = "admin_dashboard"
            st.rerun()
        else:
            st.error("بيانات دخول المشرف غير صحيحة")

def render_admin_dashboard():
    render_header()
    st.markdown(
        '<div class="section-card"><div class="axis-title">لوحة المشرف / التحليل</div></div>',
        unsafe_allow_html=True
    )

    top1, top2 = st.columns([1, 1])

    with top1:
        if st.button("تسجيل خروج المشرف"):
            reset_admin_session()
            st.rerun()

    with top2:
        if st.button("العودة للرئيسية"):
            reset_admin_session()
            st.rerun()

    results_df, error = load_results()
    if error:
        st.error(error)
        return

    if results_df is None or results_df.empty:
        st.warning("لا توجد نتائج محفوظة بعد.")
        return

    totals_df, _ = load_school_totals()
    if totals_df is None:
        st.info("لم يتم العثور على ملف أعداد الطلبة الكلي للمدارس، لذلك لن تظهر نسبة الاستجابة بشكل فعلي.")

    schools = []
    survey_types = []

    if "school" in results_df.columns:
        schools = sorted(results_df["school"].dropna().astype(str).unique().tolist())

    if "survey_type" in results_df.columns:
        survey_types = sorted(results_df["survey_type"].dropna().astype(str).unique().tolist())

    f1, f2 = st.columns(2)

    with f1:
        selected_school = st.selectbox("اختر المدرسة", ["جميع المدارس"] + schools)

    with f2:
        selected_survey_type = st.selectbox("اختر نوع الاستبانة", ["جميع الأنواع"] + survey_types)

    filtered_df = results_df.copy()

    if selected_school != "جميع المدارس":
        filtered_df = filtered_df[filtered_df["school"].astype(str) == selected_school]

    if selected_survey_type != "جميع الأنواع" and "survey_type" in filtered_df.columns:
        filtered_df = filtered_df[filtered_df["survey_type"].astype(str) == selected_survey_type]

    if filtered_df.empty:
        st.warning("لا توجد بيانات مطابقة للفلاتر المختارة.")
        return

    filtered_df["overall_avg"] = pd.to_numeric(filtered_df["overall_avg"], errors="coerce")
    overall_avg = round(filtered_df["overall_avg"].mean(), 2) if "overall_avg" in filtered_df.columns else 0
    overall_pct = score_to_percentage(overall_avg)
    responses_count = len(filtered_df)
    unique_schools = filtered_df["school"].nunique() if "school" in filtered_df.columns else 0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("عدد الاستجابات", responses_count)
    m2.metric("المتوسط الكلي", overall_avg if not pd.isna(overall_avg) else 0)
    m3.metric("النسبة المئوية", f"{overall_pct}%")
    m4.metric("عدد المدارس", unique_schools if selected_school == "جميع المدارس" else 1)

    st.markdown("## متوسطات المحاور")
    axis_summary_df = build_axis_summary(filtered_df)
    st.dataframe(axis_summary_df, use_container_width=True)

    if not axis_summary_df.empty:
        st.markdown("### الرسم البياني لمتوسطات المحاور")
        render_bar_chart(
            axis_summary_df,
            x_col="المحور",
            y_col="النسبة المئوية",
            title="النسبة المئوية لمتوسطات المحاور"
        )

    st.markdown("## متوسطات الفقرات")
    question_summary_df = build_question_summary(filtered_df)
    st.dataframe(question_summary_df, use_container_width=True, height=500)

    st.markdown("## ملخص جميع المدارس")
    school_summary_df = build_school_summary(results_df)
    school_summary_df = school_summary_df[
        ["اسم المدرسة", "نوع الاستبانة", "عدد الاستجابات", "عدد الطلبة الكلي", "نسبة الاستجابة", "المتوسط الكلي", "النسبة المئوية"]
    ]
    st.dataframe(school_summary_df, use_container_width=True)

    if not school_summary_df.empty:
        st.markdown("### الرسم البياني لمتوسطات المدارس")
        school_chart_df = school_summary_df.copy()
        school_chart_df["المدرسة والنوع"] = (
            school_chart_df["اسم المدرسة"].astype(str) + " - " + school_chart_df["نوع الاستبانة"].astype(str)
        )

        render_bar_chart(
            school_chart_df,
            x_col="المدرسة والنوع",
            y_col="النسبة المئوية",
            title="النسبة المئوية لمتوسطات المدارس"
        )

        st.markdown("### الرسم البياني لنسبة الاستجابة لكل مدرسة")
        render_bar_chart(
            school_chart_df,
            x_col="المدرسة والنوع",
            y_col="نسبة الاستجابة",
            title="نسبة الاستجابة لكل مدرسة"
        )

    with st.expander("عرض بيانات التواصل والمعبئ", expanded=False):
        cols_to_show = [
            "student_id",
            "student_name",
            "school",
            "survey_type",
            "respondent_type",
            "father_job",
            "mother_job",
            "contact_phone"
        ]
        available_cols = [col for col in cols_to_show if col in filtered_df.columns]
        st.dataframe(filtered_df[available_cols], use_container_width=True)

    with st.expander("عرض النتائج الخام", expanded=False):
        st.dataframe(filtered_df, use_container_width=True, height=350)

    st.markdown("## تنزيل الملفات")

    pdf_bytes = None
    pdf_error = None

    try:
        pdf_bytes = build_pdf_report_bytes(
            filtered_df=filtered_df,
            axis_summary_df=axis_summary_df,
            question_summary_df=question_summary_df,
            school_summary_df=school_summary_df
        )
    except Exception as e:
        pdf_error = str(e)

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.download_button(
            label="تنزيل النتائج الخام Excel",
            data=dataframe_to_excel_bytes({"Results": filtered_df}),
            file_name="filtered_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col2:
        st.download_button(
            label="تنزيل ملخص التحليل Excel",
            data=dataframe_to_excel_bytes({
                "Axis Summary": axis_summary_df,
                "Question Summary": question_summary_df,
                "School Summary": school_summary_df
            }),
            file_name="analysis_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col3:
        csv_data = filtered_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="تنزيل النتائج CSV",
            data=csv_data,
            file_name="filtered_results.csv",
            mime="text/csv",
            use_container_width=True
        )

    with col4:
        if pdf_bytes is not None:
            st.download_button(
                label="تنزيل التقرير PDF",
                data=pdf_bytes,
                file_name="survey_report.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        else:
            st.warning("تعذر إنشاء PDF")
            if pdf_error:
                st.caption(pdf_error)

# =========================
# التشغيل
# =========================
init_session()

if st.session_state.page == "home":
    render_home()
elif st.session_state.page == "parent_login":
    render_parent_login()
elif st.session_state.page == "student_info":
    render_student_info_page()
elif st.session_state.page == "survey":
    render_survey_page()
elif st.session_state.page == "admin_login":
    render_admin_login()
elif st.session_state.page == "admin_dashboard":
    render_admin_dashboard()
    