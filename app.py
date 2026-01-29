import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# إعدادات الصفحة
st.set_page_config(page_title="Excel to VCF Converter", layout="centered")

st.title("📱 Excel to VCF Converter")
st.write("حول ملفات الإكسيل لجهات اتصال (VCF) بسهولة وبدون تسطيب برامج.")

# 1. رفع الملف
uploaded_file = st.file_uploader("ارفع ملف الإكسيل هنا (XLSX)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # قراءة الملف
        df = pd.read_excel(uploaded_file)
        st.success("تم قراءة الملف بنجاح! ✅")
        st.dataframe(df.head(3)) # عرض أول 3 صفوف للتأكد

        # 2. اختيار الأعمدة (عشان البرنامج يفهم فين الاسم وفين الرقم)
        st.subheader("⚙️ ضبط البيانات")
        col1, col2 = st.columns(2)
        with col1:
            name_col = st.selectbox("اختر عمود 'الاسم'", df.columns)
        with col2:
            phone_col = st.selectbox("اختر عمود 'رقم الهاتف'", df.columns)

        # دالة التحويل (تم التعديل لتناسب الآيفون والأندرويد)
        def convert_to_vcf(dataframe, name_c, phone_c):
            vcf_data = ""
            for index, row in dataframe.iterrows():
                name = str(row[name_c]).strip()
                phone = str(row[phone_c]).strip()
                
                # تخطي الصفوف الفاضية
                if name == "nan" or phone == "nan" or name == "" or phone == "":
                    continue
                
                # الكود السحري لدعم الآيفون (N Field)
                vcf_data += "BEGIN:VCARD\n"
                vcf_data += "VERSION:3.0\n"
                vcf_data += f"N:;{name};;;\n"
                vcf_data += f"FN:{name}\n"
                vcf_data += f"TEL;TYPE=CELL:{phone}\n"
                vcf_data += "END:VCARD\n"
            return vcf_data

        # زر البدء
        if st.button("ابدأ التحويل 🔄"):
            vcf_result = convert_to_vcf(df, name_col, phone_col)
            
            # حفظ النتيجة في الـ Session State
            st.session_state['vcf_result'] = vcf_result
            st.session_state['file_ready'] = True

    except Exception as e:
        st.error(f"حدث خطأ أثناء قراءة الملف: {e}")

# 3. خيارات المخرجات (تظهر فقط بعد التحويل)
if st.session_state.get('file_ready'):
    st.divider()
    st.subheader("📂 الملف جاهز! اختر طريقة الاستلام:")
    
    col_dl, col_email = st.columns(2)
    
    # الخيار الأول: تحميل مباشر
    with col_dl:
        st.download_button(
            label="تحميل مباشر (Download) ⬇️",
            data=st.session_state['vcf_result'].encode('utf-8'),
            file_name="contacts.vcf",
            mime="text/vcard"
        )

    # الخيار الثاني: إرسال للإيميل
    with col_email:
        with st.form("email_form"):
            email_receiver = st.text_input("اكتب إيميلك هنا لاستلام الملف:")
            submit_email = st.form_submit_button("إرسال للإيميل 📧")
            
            if submit_email and email_receiver:
                try:
                    # إعدادات الإيميل المرسل (من Secrets)
                    sender_email = st.secrets["EMAIL_USER"]
                    sender_password = st.secrets["EMAIL_PASSWORD"]
                    
                    msg = MIMEMultipart()
                    msg['From'] = sender_email
                    msg['To'] = email_receiver
                    msg['Subject'] = "Your Converted VCF File is Ready! 📁"
                    
                    body = "مرحباً،\nمرفق ملف جهات الاتصال (VCF) الذي قمت بتحويله.\nيعمل الآن بكفاءة على iPhone و Android.\n\nتحياتنا."
                    msg.attach(MIMEText(body, 'plain'))
                    
                    # إرفاق الملف
                    attachment = MIMEBase('application', 'octet-stream')
                    attachment.set_payload(st.session_state['vcf_result'].encode('utf-8'))
                    encoders.encode_base64(attachment)
                    attachment.add_header('Content-Disposition', "attachment; filename=contacts.vcf")
                    msg.attach(attachment)
                    
                    # الاتصال بالسيرفر وإرسال الرسالة
                    server = smtplib.SMTP('smtp.gmail.com', 587)
                    server.starttls()
                    server.login(sender_email, sender_password)
                    text = msg.as_string()
                    server.sendmail(sender_email, email_receiver, text)
                    server.quit()
                    
                    st.success(f"تم إرسال الملف إلى {email_receiver} بنجاح! 🚀")
                    
                except Exception as e:
                    st.error(f"حدث خطأ في الإرسال. تأكد من إعدادات الإيميل في Secrets.\nالخطأ: {e}")
