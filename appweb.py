import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

# إعداد واجهة البرنامج
st.set_page_config(page_title="نظام النماذج الموحد", layout="centered")
st.title("📝 مدرسة الإمام النووي")
st.subheader("تصدير جميع الطلاب المحددين في ملف واحد")

# 1. رفع ملف الإكسل
uploaded_excel = st.file_uploader("ارفع ملف الطلاب (Excel)", type=["xlsx"])

if uploaded_excel:
    df = pd.read_excel(uploaded_excel)
    
    # اختيار الطلاب
    selected_students = st.multiselect("اختر الطلاب (يمكنك اختيار الجميع):", df['اسم الطالب'].tolist())
    
    # اختيار نوع المخالفة
    reason = st.text_input("سبب النموذج (مثال: تأخر عن الطابور):")

    if st.button("تجهيز ملف PDF الموحد"):
        if not selected_students:
            st.error("الرجاء اختيار طالب واحد على الأقل")
        else:
            # فتح القالب المرفوع على GitHub
            # تأكد أن ملف template.docx موجود في نفس المجلد على GitHub
            
            output_doc = Document() # إنشاء مستند جديد للدمج
            
            for index, name in enumerate(selected_students):
                # فتح نسخة من القالب لكل طالب
                template = Document("template.docx")
                
                # استبدال البيانات في القالب
                for p in template.paragraphs:
                    if '[A]' in p.text:
                        p.text = p.text.replace('[A]', name)
                    if '[T]' in p.text:
                        p.text = p.text.replace('[T]', reason)
                
                # إضافة محتوى القالب المعدل إلى المستند الرئيسي
                for element in template.element.body:
                    output_doc.element.body.append(element)
                
                # إضافة فاصل صفحات بين الطلاب (إلا الطالب الأخير)
                if index < len(selected_students) - 1:
                    output_doc.add_page_break()

            # حفظ الملف الموحد في الذاكرة
            target_file = BytesIO()
            output_doc.save(target_file)
            target_file.seek(0)

            st.success(f"تم تجهيز نماذج ({len(selected_students)}) طلاب بنجاح!")
            st.download_button(
                label="تحميل الملف الموحد (Word)",
                data=target_file,
                file_name="جميع_النماذج.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

st.info("ملاحظة: بعد تحميل الملف، افتحه من جوالك واختر (طباعة -> حفظ كـ PDF) للحصول على ملف واحد بصيغة PDF.")