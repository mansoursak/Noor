import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

st.set_page_config(page_title="نظام مدرسة الإمام النووي", layout="centered")
st.title("📝 مدرسة الإمام النووي")
st.subheader("تصدير النماذج (الاسم والصف) في ملف واحد")

uploaded_excel = st.file_uploader("ارفع ملف الطلاب (Excel)", type=["xlsx"])

if uploaded_excel:
    try:
        df = pd.read_excel(uploaded_excel)
        
        # اختيار الأعمدة الصحيحة
        cols = df.columns.tolist()
        name_col = st.selectbox("اختر عمود الأسماء:", cols)
        class_col = st.selectbox("اختر عمود الصفوف:", cols)
        
        # دمج الاسم والصف للعرض في القائمة فقط
        df['display_name'] = df[name_col].astype(str) + " - " + df[class_col].astype(str)
        
        selected_display = st.multiselect("اختر الطلاب المراد تصديرهم:", df['display_name'].tolist())
        reason = st.text_input("سبب النموذج:")

        if st.button("🚀 إنشاء وتحميل الملف الموحد"):
            if not selected_display:
                st.error("الرجاء اختيار طالب واحد على الأقل")
            else:
                output_doc = Document()
                
                # تصفية البيانات المختارة فقط
                selected_df = df[df['display_name'].isin(selected_display)]
                
                for index, row in selected_df.iterrows():
                    template = Document("template.docx")
                    
                    # استبدال البيانات (تأكد أن القالب يحتوي على هذه الرموز)
                    for p in template.paragraphs:
                        if '[A]' in p.text: # رمز الاسم
                            p.text = p.text.replace('[A]', str(row[name_col]))
                        if '[C]' in p.text: # رمز الصف (أضف هذا الرمز في قالبك)
                            p.text = p.text.replace('[C]', str(row[class_col]))
                        if '[T]' in p.text: # رمز السبب
                            p.text = p.text.replace('[T]', reason)
                    
                    # دمج المحتوى
                    for element in template.element.body:
                        output_doc.element.body.append(element)
                    
                    if index < len(selected_df) - 1:
                        output_doc.add_page_break()

                target_file = BytesIO()
                output_doc.save(target_file)
                target_file.seek(0)

                st.success(f"✅ تم تجهيز نماذج ({len(selected_display)}) طلاب بنجاح!")
                st.download_button(
                    label="📥 تحميل الملف الموحد (Word)",
                    data=target_file,
                    file_name="النماذج_المكتملة.docx"
                )
    except Exception as e:
        st.error(f"حدث خطأ: {e}")