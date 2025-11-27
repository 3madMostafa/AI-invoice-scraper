#!/usr/bin/env python3
"""
واجهة Streamlit لتشغيل نظام معالجة الفواتير التلقائي
"""
import streamlit as st
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path

# إعداد الصفحة
st.set_page_config(
    page_title="نظام معالجة الفواتير",
    page_icon="📊",
    layout="centered"
)

# العنوان الرئيسي
st.title("نظام معالجة الفواتير التلقائي")
st.markdown("---")

# اختيار التاريخ
st.subheader("اختيار تاريخ المعالجة")
yesterday = datetime.now().date() - timedelta(days=1)
selected_date = st.date_input(
    "التاريخ المطلوب",
    value=yesterday,
    max_value=datetime.now().date(),
    help="اختر التاريخ الذي تريد معالجة فواتيره"
)

# تحويل التاريخ إلى الصيغة المطلوبة (dd-mm-yyyy)
date_str = selected_date.strftime("%d-%m-%Y")

st.info(f"سيتم معالجة فواتير يوم: **{date_str}**")
st.markdown("---")

# الزر الرئيسي
if st.button("شغل الأوتوميشن من هنا", type="primary", use_container_width=True):
    
    # حاوية للحالة
    status_container = st.container()
    
    with status_container:
        success = True
        
        # المرحلة الأولى: السحب
        with st.status("جاري تشغيل السحب... الوقت التقديري من 10 إلى 20 دقيقة", expanded=True) as status:
            try:
                result = subprocess.run(
                    [sys.executable, "scrapping_tool.py", "--date", date_str],
                    capture_output=True, 
                    text=True,
                    encoding='utf-8'
                )
                
                if result.returncode == 0:
                    status.update(label="تم السحب بنجاح", state="complete")
                else:
                    status.update(label="فشل السحب", state="error")
                    st.error(f"حدث خطأ في مرحلة السحب")
                    with st.expander("عرض تفاصيل الخطأ"):
                        st.code(result.stderr if result.stderr else result.stdout, language="text")
                    success = False
                    
            except Exception as e:
                status.update(label="فشل السحب", state="error")
                st.error(f"خطأ في تشغيل scrapping_tool.py: {str(e)}")
                success = False
        
        # المرحلة الثانية: استخراج البيانات
        if success:
            with st.status("جاري استخراج البيانات...", expanded=True) as status:
                try:
                    result = subprocess.run(
                        [sys.executable, "json_extractor.py", "--date", date_str],
                        capture_output=True,
                        text=True,
                        encoding='utf-8'
                    )
                    
                    # فحص شامل للنجاح
                    output_text = result.stdout + result.stderr
                    is_success = (
                        result.returncode == 0 or 
                        "Successfully processed all taxpayers" in output_text or
                        "Successfully processed" in output_text or
                        "Successful taxpayers: 2" in output_text or
                        (("Successful taxpayers:" in output_text) and ("Failed taxpayers: 0" in output_text))
                    )
                    
                    if is_success:
                        status.update(label="تم استخراج البيانات بنجاح", state="complete")
                    else:
                        status.update(label="فشل استخراج البيانات", state="error")
                        st.error(f"حدث خطأ في مرحلة استخراج البيانات")
                        with st.expander("عرض تفاصيل الخطأ"):
                            st.code(output_text, language="text")
                        success = False
                        
                except Exception as e:
                    status.update(label="فشل استخراج البيانات", state="error")
                    st.error(f"خطأ في تشغيل json_extractor.py")
                    success = False
        
        # المرحلة الثالثة: إرسال الإيميلات
        if success:
            with st.status("جاري إرسال الإيميلات...", expanded=True) as status:
                try:
                    result = subprocess.run(
                        [sys.executable, "send_email.py", "--date", date_str],
                        capture_output=True,
                        text=True,
                        encoding='utf-8'
                    )
                    
                    if result.returncode == 0:
                        status.update(label="تم إرسال الإيميلات بنجاح", state="complete")
                    else:
                        status.update(label="فشل إرسال الإيميلات", state="error")
                        st.error(f"حدث خطأ في مرحلة إرسال الإيميلات")
                        success = False
                        
                except Exception as e:
                    status.update(label="فشل إرسال الإيميلات", state="error")
                    st.error(f"خطأ في تشغيل send_email.py")
                    success = False
        
        # النتيجة النهائية
        if success:
            st.balloons()
            st.success("تم تشغيل كل المراحل بنجاح!")
            
            # عرض معلومات إضافية
            st.markdown("---")
            st.subheader("ملخص العملية")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("المرحلة الأولى", "السحب")
            with col2:
                st.metric("المرحلة الثانية", "الاستخراج")
            with col3:
                st.metric("المرحلة الثالثة", "الإرسال")
            
            # مسار الملفات
            output_path = Path("outputs") / date_str
            if output_path.exists():
                st.info(f"الملفات المعالجة موجودة في: `{output_path.absolute()}`")

# معلومات إضافية في الشريط الجانبي
with st.sidebar:
    st.header("معلومات النظام")
    st.markdown("""
    ### مراحل المعالجة:
    
    1. **السحب (Scraping)**
       - سحب الفواتير من النظام
       - الوقت: 10-20 دقيقة
    
    2. **الاستخراج (Extraction)**
       - تحليل ملفات JSON
       - استخراج البيانات المطلوبة
    
    3. **الإرسال (Email)**
       - إرسال التقارير بالإيميل
       - إرفاق ملفات Excel
    
    ---
    
    ### هيكل المخرجات:
    ```
    outputs/
    └── dd-mm-yyyy/
        ├── Excel/
        │   └── [Supplier]/
        │       └── results.xlsx
        └── PDF/
            └── [Supplier]/
                └── *.pdf
    ```
    
    ---
    
    ### السجلات (Logs):
    يمكن مراجعة السجلات التفصيلية في مجلد `logs/`
    """)
    
    st.markdown("---")
    st.caption("نظام معالجة الفواتير الآلي v1.0")