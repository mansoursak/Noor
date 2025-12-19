import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
import io
from datetime import datetime

# 1. دالة قوية لتحويل الأرقام وإجبارها على التنسيق الإنجليزي
def force_english_numbers(text):
    arabic_numbers = "٠١٢٣٤٥٦٧٨٩"
    english_numbers = "0123456789"
    translation_table = str.maketrans(arabic_numbers, english_numbers)
    return str(text).translate(translation_table)

# إعداد واجهة التطبيق
st.set_page_config(page_title="نظام النماذج الذكي", layout="wide")
st.title("🎓 نظام إصدار النماذج - مدرسة الإمام النووي")

# 2. منطقة رفع الملفات
col1, col2 = st.columns(2)
with col1:
    up_excel = st.file_uploader("1. ارفع ملف الطلاب (Excel)", type="xlsx")
with col2:
    up_template = st.file_uploader("2. ارفع نموذج الوورد (Word)", type="docx")

if up_excel and up_template:
    # قراءة أسماء الفصول كما هي في الإكسل
    excel_data = pd.ExcelFile(up_excel)
    sheet_names = excel_data.sheet_names 
    
    st.divider()
    
    # 3. واجهة مدخلات المستخدم
    col_input1, col_input2 = st.columns(2)
    with col_input1:
        selected_sheet = st.selectbox("📁 اختر الفصل الدراسي:", sheet_names)
        df = pd.read_excel(up_excel, sheet_name=selected_sheet)
        selected_students = st.multiselect("👥 اختر الطلاب المطلوبين:", df['الاسم'].tolist())

    with col_input2:
        reasons_list = st.multiselect(
            "أسباب التحويل (سيتم وضع ✔️ في المربعات):", 
            ["عدم أداء الواجب", "ضعف دراسي", "مشاغبة", "تأخر عن الحصة", "أخرى"]
        )
        other_text = st.text_input("في حال اخترت 'أخرى' اذكر السبب هنا [F]:")

    # 4. بيانات إضافية
    st.subheader("✍️ تعبئة بيانات النموذج")
    col_a, col_b = st.columns(2)
    with col_a:
        problem_desc = st.text_area("إيضاح المشكلة [S]:")
    with col_b:
        # توليد التاريخ وتحويل أرقامه فوراً
        today_raw = datetime.now().strftime("%d / %m / 1446 هـ")
        today_auto = force_english_numbers(today_raw)
        doc_date = st.text_input("التاريخ [T]:", value=today_auto)

    if st.button("🚀 إنشاء وتحميل النماذج"):
        if not selected_students:
            st.warning("يرجى اختيار طالب واحد على الأقل.")
        else:
            # تنظيف التاريخ قبل الاستخدام
            final_date = force_english_numbers(doc_date)
            
            for student_name in selected_students:
                doc = Document(up_template)
                check_mark = "✔️"
                
                # قاموس الاستبدال بالحروف الإنجليزية (A, B, C...)
                replacements = {
                    "[A]": str(student_name),
                    "[B]": str(selected_sheet),
                    "[S]": problem_desc,
                    "[T]": final_date,
                    "[F]": other_text,
                    "[C]": check_mark if "عدم أداء الواجب" in reasons_list else "  ",
                    "[D]": check_mark if "ضعف دراسي" in reasons_list else "  ",
                    "[E]": check_mark if "مشاغبة" in reasons_list else "  ",
                    "[G]": check_mark if "تأخر عن الحصة" in reasons_list else "  ",
                    "[R]": check_mark if "أخرى" in reasons_list else "  ",
                }
                
                # وظيفة الاستبدال الذكية للحفاظ على التنسيق واللون
                def process_content(target):
                    for paragraph in target.paragraphs:
                        for key, value in replacements.items():
                            if key in paragraph.text:
                                for run in paragraph.runs:
                                    if key in run.text:
                                        run.text = run.text.replace(key, value)
                                        # حل مشكلة الأرقام: منع تحويلها لهندية (عربية)
                                        run.font.complex_script = False

                # تنفيذ العملية على النصوص والجداول
                process_content(doc)
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            process_content(cell)
                
                # حفظ الملف في الذاكرة للتحميل
                target_stream = io.BytesIO()
                doc.save(target_stream)
                st.download_button(
                    label=f"⬇️ تحميل نموذج: {student_name}",
                    data=target_stream.getvalue(),
                    file_name=f"نموذج_{student_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            st.success("✅ تم الانتهاء! النماذج جاهزة الآن مع الحفاظ على الخطوط والأرقام.")