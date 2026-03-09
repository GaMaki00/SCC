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

# 1. ดึงจาก PDF (หาเลขหลังคำว่า ผลการเรียนเฉลี่ยร้อยละ)
with pdfplumber.open(pdf_file) as pdf:
    last_page_text = pdf.pages[-1].extract_text()
    # ใช้ Regex หาตัวเลขทศนิยมที่ตามหลังคำว่า "ร้อยละ"
    pdf_total_match = re.search(r"เฉลี่ยร้อยละ\s*(\d+\.\d+)", last_page_text)
    pdf_total_score = pdf_total_match.group(1) if pdf_total_match else "ไม่พบข้อมูล"

# 2. ดึงจาก Excel 
# สมมติว่าอยู่ในคอลัมน์ที่ 17 (R) และแถวที่เขียนว่า 'ร้อยละ'
# เราจะค้นหาแถวที่มีคำว่า 'ร้อยละ' ในคอลัมน์แรกๆ
row_index = df_excel_full[df_excel_full.iloc[:, 0].str.contains('ร้อยละ', na=False)].index
if not row_index.empty:
    excel_total_score = df_excel_full.iloc[row_index[0], 17] # 17 คือคอลัมน์คะแนนรวม
else:
    excel_total_score = "ไม่พบข้อมูล"

# --- ส่วนการแสดงผลบน Streamlit ---
st.divider() # ขีดเส้นคั่น
st.subheader("🎯 ตรวจสอบคะแนนเฉลี่ยร้อยละรวม")
c1, c2, c3 = st.columns(3)
c1.metric("ร้อยละใน Excel", excel_total_score)
c2.metric("ร้อยละใน PDF", pdf_total_score)

# เช็กว่าตรงกันไหม
if str(excel_total_score) == str(pdf_total_score):
    c3.success("✅ ตรงกัน")
else:
    c3.error("❌ ไม่ตรงกัน")
