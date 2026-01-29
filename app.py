# دالة التحويل (تم التعديل لتناسب الآيفون)
        def convert_to_vcf(dataframe, name_c, phone_c):
            vcf_data = ""
            for index, row in dataframe.iterrows():
                name = str(row[name_c]).strip()
                phone = str(row[phone_c]).strip()
                
                # تخطي الصفوف الفاضية
                if name == "nan" or phone == "nan" or name == "" or phone == "":
                    continue
                
                # إضافة السطر السحري للآيفون (N Field)
                # بنحط الاسم كله في خانة الاسم الأول عشان نضمن يظهر كامل
                # الفواصل الكتير ;;; دي ضرورية عشان ترتيب (العائلة;الأول;الأوسط)
                
                vcf_data += "BEGIN:VCARD\n"
                vcf_data += "VERSION:3.0\n"
                vcf_data += f"N:;{name};;;\n"
                vcf_data += f"FN:{name}\n"
                vcf_data += f"TEL;TYPE=CELL:{phone}\n"
                vcf_data += "END:VCARD\n"
            return vcf_data
