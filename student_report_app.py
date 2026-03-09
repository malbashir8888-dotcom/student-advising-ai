import io
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="تقرير الطالب الأكاديمي", page_icon="🎓", layout="wide")

DEFAULT_FILE = Path("/mnt/data/students_progress_report_PLUS_FIXED.xlsx")

RISK_COLORS = {
    "High": "#fde2e1",
    "Medium": "#fff1cc",
    "Low": "#e6f4ea",
}


def norm_id(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text.replace(" ", "")


@st.cache_data(show_spinner=False)
def load_workbook(uploaded_file_bytes: bytes | None, fallback_path: str):
    if uploaded_file_bytes:
        excel_source = io.BytesIO(uploaded_file_bytes)
    else:
        excel_source = fallback_path

    xls = pd.ExcelFile(excel_source)
    sheets = {name: pd.read_excel(excel_source, sheet_name=name) for name in xls.sheet_names}

    for name, df in sheets.items():
        sheets[name] = df.copy()
        if "رقم الطالب" in df.columns:
            sheets[name]["رقم الطالب"] = df["رقم الطالب"].apply(norm_id)

    return sheets


def get_student_row(df: pd.DataFrame, student_id: str):
    if df is None or df.empty or "رقم الطالب" not in df.columns:
        return None
    match = df[df["رقم الطالب"] == student_id]
    if match.empty:
        return None
    return match.iloc[0]



def fmt_num(value):
    if pd.isna(value):
        return "—"
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.2f}"
    return str(value)



def clean_text(value):
    if pd.isna(value) or value is None or str(value).strip() == "":
        return "لا يوجد"
    return str(value)



def build_arabic_summary(row_analytics, row_action):
    if row_analytics is None and row_action is None:
        return "لا تتوفر بيانات كافية لتوليد الملخص."

    risk = clean_text((row_action.get("Risk_Level") if row_action is not None else None) or (row_analytics.get("Risk_Level") if row_analytics is not None else None))
    gpa = fmt_num((row_action.get("GPA(B)") if row_action is not None else None) or (row_analytics.get("GPA(B)") if row_analytics is not None else None))
    warnings = fmt_num((row_action.get("Warnings(B)") if row_action is not None else None) or (row_analytics.get("Warnings(B)") if row_analytics is not None else None))
    failed_count = fmt_num((row_action.get("Failed_in_B_count") if row_action is not None else None) or (row_analytics.get("Failed_in_B_count") if row_analytics is not None else None))
    failed_courses = clean_text((row_action.get("Failed_in_B_courses") if row_action is not None else None) or (row_analytics.get("Failed_in_B_courses") if row_analytics is not None else None))
    next_load = fmt_num((row_action.get("Next_Term_Total_Load") if row_action is not None else None) or (row_analytics.get("Next_Term_Total_Load") if row_analytics is not None else None))
    action = clean_text(row_action.get("Recommended_Action") if row_action is not None else None)

    if risk == "High":
        intro = "الطالب ضمن فئة الخطورة العالية ويحتاج إلى تدخل إرشادي عاجل."
    elif risk == "Medium":
        intro = "الطالب ضمن فئة الخطورة المتوسطة ويحتاج إلى متابعة أكاديمية منتظمة."
    elif risk == "Low":
        intro = "الطالب ضمن الفئة منخفضة الخطورة مع الحاجة إلى متابعة اعتيادية."
    else:
        intro = "تم توليد ملخص أكاديمي بناءً على البيانات المتاحة."

    return (
        f"{intro} المعدل الحالي {gpa}، وعدد الإنذارات {warnings}، وعدد المواد الراسبة {failed_count}. "
        f"المواد الراسبة: {failed_courses}. العبء الدراسي المتوقع للفصل القادم: {next_load}. "
        f"التوصية الإرشادية الحالية: {action}."
    )



def build_student_report(student_id: str, sheets: dict):
    analytics = sheets.get("Student_Analytics")
    actions = sheets.get("Advising_Action_List")
    summary = sheets.get("Student_Summary")
    reg_details = sheets.get("Details_Registered")
    new_regs = sheets.get("Details_NewRegistrations")

    row_analytics = get_student_row(analytics, student_id)
    row_action = get_student_row(actions, student_id)
    row_summary = get_student_row(summary, student_id)

    if row_analytics is None and row_action is None and row_summary is None:
        return None

    name = None
    for row in [row_action, row_analytics, row_summary]:
        if row is not None and "اسم الطالب" in row and pd.notna(row["اسم الطالب"]):
            name = str(row["اسم الطالب"])
            break

    details_registered = pd.DataFrame()
    if reg_details is not None and not reg_details.empty and "رقم الطالب" in reg_details.columns:
        details_registered = reg_details[reg_details["رقم الطالب"] == student_id].copy()

    details_new = pd.DataFrame()
    if new_regs is not None and not new_regs.empty and "رقم الطالب" in new_regs.columns:
        details_new = new_regs[new_regs["رقم الطالب"] == student_id].copy()

    report = {
        "student_id": student_id,
        "name": name or "غير متوفر",
        "gpa": fmt_num((row_action.get("GPA(B)") if row_action is not None else None) or (row_analytics.get("GPA(B)") if row_analytics is not None else None) or (row_summary.get("GPA(2S)") if row_summary is not None else None)),
        "warnings": fmt_num((row_action.get("Warnings(B)") if row_action is not None else None) or (row_analytics.get("Warnings(B)") if row_analytics is not None else None) or (row_summary.get("Warnings(2S)") if row_summary is not None else None)),
        "failed_count": fmt_num((row_action.get("Failed_in_B_count") if row_action is not None else None) or (row_analytics.get("Failed_in_B_count") if row_analytics is not None else None) or (row_summary.get("Failed_in_S1_count") if row_summary is not None else None)),
        "failed_courses": clean_text((row_action.get("Failed_in_B_courses") if row_action is not None else None) or (row_analytics.get("Failed_in_B_courses") if row_analytics is not None else None) or (row_summary.get("Failed_in_S1_courses") if row_summary is not None else None)),
        "registered_count": fmt_num((row_analytics.get("Registered_in_A_count") if row_analytics is not None else None) or (row_summary.get("Registered_in_S1_count") if row_summary is not None else None)),
        "registered_courses": clean_text((row_analytics.get("Registered_in_A_courses") if row_analytics is not None else None) or (row_summary.get("Registered_in_S1_courses") if row_summary is not None else None)),
        "completed_courses": clean_text((row_analytics.get("Completed_in_B_courses") if row_analytics is not None else None) or (row_summary.get("Completed_in_S2_courses") if row_summary is not None else None)),
        "new_reg_count": fmt_num((row_analytics.get("New_R_in_B_count") if row_analytics is not None else None) or (row_summary.get("New_R_in_S2_count") if row_summary is not None else None)),
        "new_reg_courses": clean_text((row_analytics.get("New_R_in_B_courses") if row_analytics is not None else None) or (row_summary.get("New_R_in_S2_courses") if row_summary is not None else None)),
        "next_load": fmt_num((row_action.get("Next_Term_Total_Load") if row_action is not None else None) or (row_analytics.get("Next_Term_Total_Load") if row_analytics is not None else None)),
        "core_failed_count": fmt_num((row_action.get("Failed_Core_Count") if row_action is not None else None) or (row_analytics.get("Failed_Core_Count") if row_analytics is not None else None)),
        "core_failed_courses": clean_text((row_action.get("Failed_Core_Courses") if row_action is not None else None) or (row_analytics.get("Failed_Core_Courses") if row_analytics is not None else None)),
        "gened_failed_count": fmt_num((row_action.get("Failed_GenEd_Count") if row_action is not None else None) or (row_analytics.get("Failed_GenEd_Count") if row_analytics is not None else None)),
        "gened_failed_courses": clean_text((row_action.get("Failed_GenEd_Courses") if row_action is not None else None) or (row_analytics.get("Failed_GenEd_Courses") if row_analytics is not None else None)),
        "risk_score": fmt_num((row_action.get("Academic_Risk_Score") if row_action is not None else None) or (row_analytics.get("Academic_Risk_Score") if row_analytics is not None else None)),
        "risk_level": clean_text((row_action.get("Risk_Level") if row_action is not None else None) or (row_analytics.get("Risk_Level") if row_analytics is not None else None)),
        "recommended_action": clean_text(row_action.get("Recommended_Action") if row_action is not None else None),
        "summary_text": build_arabic_summary(row_analytics, row_action),
        "details_registered": details_registered,
        "details_new": details_new,
    }

    return report



def report_to_text(report: dict):
    return f"""
تقرير الطالب الأكاديمي

رقم الطالب: {report['student_id']}
اسم الطالب: {report['name']}
المعدل: {report['gpa']}
عدد الإنذارات: {report['warnings']}
عدد المواد الراسبة: {report['failed_count']}
المواد الراسبة: {report['failed_courses']}
عدد المواد المسجلة سابقًا: {report['registered_count']}
المواد المسجلة سابقًا: {report['registered_courses']}
المواد المكتملة: {report['completed_courses']}
عدد المواد المسجلة الجديدة: {report['new_reg_count']}
المواد المسجلة الجديدة: {report['new_reg_courses']}
العبء الدراسي القادم: {report['next_load']}
عدد المواد الأساسية الراسبة: {report['core_failed_count']}
المواد الأساسية الراسبة: {report['core_failed_courses']}
عدد مواد المتطلبات العامة الراسبة: {report['gened_failed_count']}
مواد المتطلبات العامة الراسبة: {report['gened_failed_courses']}
درجة الخطورة: {report['risk_level']}
مؤشر الخطورة: {report['risk_score']}
التوصية الإرشادية: {report['recommended_action']}

الملخص الآلي:
{report['summary_text']}
""".strip()


st.title("🎓 أداة تقرير الطالب الأكاديمي")
st.caption("أدخل رقم الطالب أو ارفع ملف Excel نفسه، وسيتم توليد تقرير مختصر ومباشر." )

with st.sidebar:
    st.header("مصدر البيانات")
    uploaded_file = st.file_uploader("رفع ملف Excel", type=["xlsx"])
    use_default = DEFAULT_FILE.exists()
    if use_default:
        st.success("تم العثور على الملف الافتراضي داخل نفس المسار.")
    st.markdown("**الأوراق المدعومة:** Student_Analytics، Advising_Action_List، Student_Summary، Details_Registered، Details_NewRegistrations")

file_bytes = uploaded_file.getvalue() if uploaded_file is not None else None

if not file_bytes and not use_default:
    st.error("يرجى رفع ملف Excel أولًا.")
    st.stop()

try:
    sheets = load_workbook(file_bytes, str(DEFAULT_FILE))
except Exception as e:
    st.error(f"تعذر قراءة الملف: {e}")
    st.stop()

student_ids = []
for sheet_name in ["Student_Analytics", "Advising_Action_List", "Student_Summary"]:
    df = sheets.get(sheet_name)
    if df is not None and "رقم الطالب" in df.columns:
        student_ids.extend(df["رقم الطالب"].dropna().astype(str).tolist())
student_ids = sorted(set([norm_id(x) for x in student_ids if norm_id(x)]))

col1, col2 = st.columns([2, 1])
with col1:
    student_id_input = st.text_input("رقم الطالب", placeholder="مثال: 23120113")
with col2:
    selected_id = st.selectbox("أو اختر من القائمة", options=[""] + student_ids)

student_id = norm_id(student_id_input or selected_id)

if student_id:
    report = build_student_report(student_id, sheets)
    if report is None:
        st.warning("لم يتم العثور على هذا الطالب في الأوراق المعتمدة داخل الملف.")
    else:
        bg = RISK_COLORS.get(report["risk_level"], "#f3f4f6")
        st.markdown(
            f"""
            <div style='background:{bg};padding:16px;border-radius:14px;border:1px solid #ddd'>
                <h3 style='margin:0'>{report['name']}</h3>
                <div style='margin-top:8px'>رقم الطالب: <strong>{report['student_id']}</strong></div>
                <div>مستوى الخطورة: <strong>{report['risk_level']}</strong></div>
                <div>التوصية: <strong>{report['recommended_action']}</strong></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        a, b, c, d = st.columns(4)
        a.metric("GPA", report["gpa"])
        b.metric("الإنذارات", report["warnings"])
        c.metric("المواد الراسبة", report["failed_count"])
        d.metric("العبء القادم", report["next_load"])

        st.subheader("الملخص الآلي")
        st.info(report["summary_text"])

        c1, c2 = st.columns(2)
        with c1:
            st.subheader("بيانات التعثر")
            st.write(f"**المواد الراسبة:** {report['failed_courses']}")
            st.write(f"**المواد الأساسية الراسبة:** {report['core_failed_courses']}")
            st.write(f"**مواد المتطلبات العامة الراسبة:** {report['gened_failed_courses']}")
            st.write(f"**مؤشر الخطورة:** {report['risk_score']}")
        with c2:
            st.subheader("التسجيل والإنجاز")
            st.write(f"**المواد المسجلة سابقًا:** {report['registered_courses']}")
            st.write(f"**المواد المكتملة:** {report['completed_courses']}")
            st.write(f"**المواد المسجلة الجديدة:** {report['new_reg_courses']}")
            st.write(f"**عدد المواد المسجلة الجديدة:** {report['new_reg_count']}")

        if not report["details_registered"].empty:
            st.subheader("تفاصيل المقررات المسجلة سابقًا")
            st.dataframe(report["details_registered"], use_container_width=True, hide_index=True)

        if not report["details_new"].empty:
            st.subheader("تفاصيل التسجيلات الجديدة")
            st.dataframe(report["details_new"], use_container_width=True, hide_index=True)

        st.download_button(
            "تنزيل التقرير كنص",
            data=report_to_text(report).encode("utf-8-sig"),
            file_name=f"student_report_{report['student_id']}.txt",
            mime="text/plain",
        )
else:
    st.subheader("معاينة سريعة")
    dist = sheets.get("Dist_Risk_Levels")
    prog = sheets.get("Program_Summary")
    if prog is not None and not prog.empty:
        st.dataframe(prog, use_container_width=True, hide_index=True)
    if dist is not None and not dist.empty:
        st.dataframe(dist, use_container_width=True, hide_index=True)
    st.caption("اكتب رقم الطالب لعرض تقريره التفصيلي.")
