import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO
import copy

st.set_page_config(page_title="نظام مدرسة الإمام النووي", layout="centered")
st.title("📝 مدرسة الإمام النووي")
st.subheader("تصدير النماذج المكتملة في ملف واحد")

uploaded_excel = st.file_uploader("ارفع ملف الطلاب (Excel)", type=["xlsx"])

if uploaded_excel:
    try:
        df = pd.read_excel(uploaded_excel)
        cols = df.columns.tolist()
        
        col1, col2 = st.columns(2)
        with col1:
            name_col = st.selectbox("اختر عمود الأسماء:", cols)
        with col2:
            class_col = st.selectbox("اختر عمود الصفوف:", cols)
        
        df['display'] = df[name_col].astype(str) + " - " + df[class_col].astype(str)
        selected_display = st.multiselect("اختر الطلاب:", df['display'].tolist())
        reason = st.text_input("التاريخ (أو بيانات إضافية لرمز [T]):")

        if st.button("🚀 إنشاء وتحميل الملف الموحد"):
            if not selected_display:
                st.error("الرجاء اختيار طلاب")
            else:
                combined_doc = Document()
                selected_df = df[df['display'].isin(selected_display)]
                
                for index, (idx, row) in enumerate(selected_df.iterrows()):
                    # فتح القالب الأصلي لكل طالب
                    template = Document("template.docx")
                    
                    # دالة الاستبدال داخل النصوص والجداول
                    def replace_in_doc(doc):
                        # استبدال في الفقرات
                        for p in doc.paragraphs:
                            if '[A]' in p.text: p.text = p.text.replace('[A]', str(row[name_col]))
                            if '[B]' in p.text: p.text = p.text.replace('[B]', str(row[class_col]))
                            if '[T]' in p.text: p.text = p.text.replace('[T]', reason)
                        # استبدال في الجداول (ضروري لقالبك)
                        for table in doc.tables:
                            for r_obj in table.rows:
                                for cell in r_obj.cells:
                                    for paragraph in cell.paragraphs:
                                        if '[A]' in paragraph.text: paragraph.text = paragraph.text.replace('[A]', str(row[name_col]))
                                        if '[B]' in paragraph.text: paragraph.text = paragraph.text.replace('[B]', str(row[class_col]))
                                        if '[T]' in paragraph.text: paragraph.text = paragraph.text.replace('[T]', reason)

                    replace_in_doc(template)
                    
                    # نقل جميع محتويات القالب (جداول وفقرات) للمستند الموحد
                    for element in template.element.body:
                        combined_doc.element.body.append(element)
                    
                    # إضافة فاصل صفحات
                    if index < len(selected_df) - 1:
                        combined_doc.add_page_break()

                # حفظ وتنزيل
                target_file = BytesIO()
                combined_doc.save(target_file)
                target_file.seek(0)

                st.success(f"✅ تم دمج {len(selected_display)} نماذج بنجاح!")
                st.download_button(
                    label="📥 تحميل الملف الموحد",
                    data=target_file,
                    file_name="النماذج_النهائية.docx"
                )
    except Exception as e:
        st.error(f"حدث خطأ: {e}")