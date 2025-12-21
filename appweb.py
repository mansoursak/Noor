import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

# إعدادات واجهة البرنامج
st.set_page_config(page_title="نظام النماذج الموحد", layout="centered")
st.title("📝 مدرسة الإمام النووي")
st.subheader("تصدير نماذج الطلاب المحددين في ملف واحد")

# 1. رفع ملف الإكسل
uploaded_excel = st.file_uploader("ارفع ملف الطلاب (Excel)", type=["xlsx"])

if uploaded_excel:
    try:
        df = pd.read_excel(uploaded_excel)
        
        # اختيار الطلاب (يمكن اختيار أكثر من اسم)
        selected_students = st.multiselect("اختر الطلاب المراد تصدير نماذجهم:", df['اسم الطالب'].tolist())
        
        # إدخال سبب النموذج
        reason = st.text_input("سبب النموذج (سيطبق على جميع المختارين):")

        if st.button("تجهيز الملف الموحد"):
            if not selected_students:
                st.error("الرجاء اختيار طالب واحد على الأقل.")
            elif not reason:
                st.warning("الرجاء كتابة السبب.")
            else:
                # إنشاء مستند جديد لجمع كل الصفحات فيه
                combined_doc = Document()
                
                for i, name in enumerate(selected_students):
                    # فتح قالب الوورد لكل طالب
                    template = Document("template.docx")
                    
                    # استبدال الكلمات المحجوزة
                    for p in template.paragraphs:
                        if '[A]' in p.text:
                            p.text = p.text.replace('[A]', name)
                        if '[T]' in p.text:
                            p.text = p.text.replace('[T]', reason)
                    
                    # إضافة محتوى القالب المعدل للمستند الرئيسي
                    for element in template.element.body:
                        combined_doc.element.body.append(element)
                    
                    # إضافة فاصل صفحات إلا بعد الطالب الأخير
                    if i < len(selected_students) - 1:
                        combined_doc.add_page_break()

                # حفظ الملف الموحد في الذاكرة
                file_stream = BytesIO()
                combined_doc.save(file_stream)
                file_stream.seek(0)

                st.success(f"تم بنجاح تجهيز نماذج ({len(selected_students)}) طلاب في ملف واحد.")
                st.download_button(
                    label="📥 تحميل ملف النماذج الموحد (Word)",
                    data=file_stream,
                    file_name="نماذج_الطلاب_الموحدة.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
    except Exception as e:
        st.error(f"حدث خطأ في قراءة الملف: {e}")