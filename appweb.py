import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

st.set_page_config(page_title="نظام النماذج الموحد", layout="centered")
st.title("📝 مدرسة الإمام النووي")
st.subheader("تصدير جميع الطلاب المحددين في ملف واحد")

uploaded_excel = st.file_uploader("ارفع ملف الطلاب (Excel)", type=["xlsx"])

if uploaded_excel:
    try:
        df = pd.read_excel(uploaded_excel)
        
        # التأكد من اسم العمود الصحيح
        columns = df.columns.tolist()
        student_col = st.selectbox("اختر العمود الذي يحتوي على أسماء الطلاب:", columns)
        
        selected_students = st.multiselect("اختر الطلاب:", df[student_col].tolist())
        reason = st.text_input("سبب النموذج:")

        if st.button("تجهيز ملف الـ PDF الموحد"):
            if not selected_students:
                st.error("الرجاء اختيار طالب واحد على الأقل")
            else:
                output_doc = Document()
                for index, name in enumerate(selected_students):
                    # فتح القالب (تأكد من وجود ملف template.docx في GitHub)
                    template = Document("template.docx")
                    
                    for p in template.paragraphs:
                        if '[A]' in p.text:
                            p.text = p.text.replace('[A]', str(name))
                        if '[T]' in p.text:
                            p.text = p.text.replace('[T]', reason)
                    
                    for element in template.element.body:
                        output_doc.element.body.append(element)
                    
                    if index < len(selected_students) - 1:
                        output_doc.add_page_break()

                target_file = BytesIO()
                output_doc.save(target_file)
                target_file.seek(0)

                st.success(f"تم تجهيز نماذج ({len(selected_students)}) طلاب بنجاح!")
                st.download_button(
                    label="📥 تحميل الملف الموحد (Word)",
                    data=target_file,
                    file_name="جميع_النماذج.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
    except Exception as e:
        st.error(f"حدث خطأ: {e}")