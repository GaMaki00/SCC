import streamlit as st
import pandas as pd
import pdfplumber
import io
import re 

st.set_page_config(page_title="ระบบเทียบคะแนน ปพ.", layout="wide")

st.title("📊 ระบบตรวจสอบเทียบเคียงคะแนน Excel vs PDF")
st.write("อัปโหลดไฟล์เพื่อตรวจสอบความถูกต้องของคะแนนและเกรด")

# 1. ส่วนการอัปโหลดไฟล์
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. เลือกไฟล์ Excel ", type=['xlsx'])
with col2:
    pdf_file = st.file_uploader("2. เลือกไฟล์ PDF เพื่อเทียบ", type=['pdf'])

if excel_file and pdf_file:
    if st.button("เริ่มการตรวจสอบ"):
        try:
            # อ่าน Excel
            xl = pd.ExcelFile(excel_file)
            df_excel = pd.read_excel(excel_file, sheet_name=xl.sheet_names[0], skiprows=2)
            df_excel = df_excel.iloc[:, [1, 3, 4, 17, 18]]
            df_excel.columns = ['ID', 'ชื่อ', 'นามสกุล', 'คะแนน_Excel', 'เกรด_Excel']
            df_excel['ID'] = df_excel['ID'].astype(str).str.strip().str.replace('.0', '', regex=False)
            df_excel['ชื่อ'] = df_excel['ชื่อ'].astype(str).str.strip()
            df_excel['นามสกุล'] = df_excel['นามสกุล'].astype(str).str.strip()
            
            # อ่าน PDF
            pdf_list = []
            with pdfplumber.open(pdf_file) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        for row in table:
                            if row[0] and str(row[0]).strip().isdigit():
                                pdf_list.append({
                                    'ID': str(row[1]).strip(), 
                                    'คะแนน_PDF': row[9], 
                                    'เกรด_PDF': row[10]
                                })
            
            df_pdf = pd.DataFrame(pdf_list)
            df_final = pd.merge(df_excel, df_pdf, on='ID', how='outer')
            df_final = df_final[['ID', 'ชื่อ', 'นามสกุล', 'คะแนน_Excel', 'คะแนน_PDF', 'เกรด_Excel', 'เกรด_PDF']]

            # Logic การตรวจสอบ (Highlight)
            def highlight_logic(row):
                def safe_float(val):
                    try: return float(str(val).strip().replace(',', ''))
                    except: return None
                
                s_ex, s_pdf = safe_float(row['คะแนน_Excel']), safe_float(row['คะแนน_PDF'])
                g_ex, g_pdf = str(row['เกรด_Excel']).strip(), str(row['เกรด_PDF']).strip()
                
                if s_ex is not None and s_pdf is not None:
                    bg = 'background-color: #C6EFCE' if (s_ex == s_pdf and g_ex == g_pdf) else 'background-color: #FFC7CE'
                else: 
                    bg = 'background-color: #FFEB9C'
                return [bg] * len(row)

            # แสดงผลบนหน้าเว็บ
            st.subheader("📋 ผลการตรวจสอบ")
            styled_df = df_final.style.apply(highlight_logic, axis=1)
            st.dataframe(styled_df, use_container_width=True)

            # 2. ส่วนการดาวน์โหลดไฟล์ (แทนที่ os.startfile)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                styled_df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์สรุป (Excel)",
                data=buffer.getvalue(),
                file_name="สรุปเทียบ_ผลตรวจสอบ.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")

# --- 3. ส่วนการตรวจสอบ "ร้อยละ" ท้ายตาราง ---
        st.divider()
        st.subheader("🎯 ตรวจสอบค่าสรุปท้ายตาราง (ร้อยละรวม)")

        try:
            # --- ดึงจาก Excel (หาแถวที่มีคำว่า "ร้อยละ") ---
            # อ่าน Excel แบบเต็มๆ ไม่ข้ามแถว เพื่อหาบรรทัดสรุป
            df_full = pd.read_excel(excel_file, sheet_name=xl.sheet_names[0])
            
            # ค้นหาแถวที่มีคำว่า "ร้อยละ" ในคอลัมน์ใดก็ได้
            mask = df_full.apply(lambda row: row.astype(str).str.contains('ร้อยละ').any(), axis=1)
            summary_row = df_full[mask]

            if not summary_row.empty:
                # ดึงค่าจากคอลัมน์ที่ 17 (R) ของแถวนั้น
                excel_total = summary_row.iloc[0, 17] 
                # ทำความสะอาดตัวเลข (ปัดทศนิยมให้เท่ากับ PDF)
                excel_total_val = f"{float(excel_total):.2f}"
            else:
                excel_total_val = "ไม่พบแถวร้อยละ"

            # --- ดึงจาก PDF (หาตัวเลขหลังคำว่า ผลการเรียนเฉลี่ยร้อยละ) ---
            import re
            pdf_total_val = "ไม่พบใน PDF"
            with pdfplumber.open(pdf_file) as pdf:
                # ตรวจสอบ 2 หน้าสุดท้าย (เผื่อข้อมูลสรุปอยู่หน้าแยก)
                pages_to_check = pdf.pages[-2:] if len(pdf.pages) > 1 else pdf.pages
                for page in pages_to_check:
                    text = page.extract_text()
                    if text:
                        # ใช้ Regex ดึงตัวเลขที่อยู่หลังคำว่า "ร้อยละ"
                        match = re.search(r"ร้อยละ\s*(\d+\.\d+)", text)
                        if match:
                            pdf_total_val = match.group(1)
                            break

            # --- แสดงผลเปรียบเทียบ ---
            col_ex, col_pdf, col_res = st.columns(3)
            
            with col_ex:
                st.metric("ร้อยละจาก Excel", excel_total_val)
            with col_pdf:
                st.metric("ร้อยละจาก PDF", pdf_total_val)
            with col_res:
                if excel_total_val == pdf_total_val:
                    st.success("✅ ค่าร้อยละตรงกัน")
                else:
                    st.error("❌ ค่าร้อยละไม่ตรงกัน!")
                    st.caption(f"ส่วนต่าง: {abs(float(excel_total_val) - float(pdf_total_val)):.2f}")

        except Exception as e:
            st.warning(f"ไม่สามารถตรวจสอบค่าร้อยละสรุปได้อัตโนมัติ: {e}")
