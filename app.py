import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

st.set_page_config(page_title="ระบบตรวจคะแนน ปพ.", layout="wide")

st.title("📊 ระบบตรวจสอบคะแนน Excel vs PDF (แบบหลายห้อง)")
st.info("💡 วิธีใช้: อัปโหลดไฟล์ Excel ที่มีหลายห้องในชีตเดียว และไฟล์ PDF ที่เรียงลำดับเดียวกัน")

# 1. ส่วนการอัปโหลดไฟล์
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. เลือกไฟล์ Excel ", type=['xlsx'])
with col2:
    pdf_file = st.file_uploader("2. เลือกไฟล์ PDF เพื่อเทียบ", type=['pdf'])

if excel_file and pdf_file:
    if st.button("🚀 เริ่มประมวลผลทั้งหมด"):
        try:
            # --- เตรียมข้อมูล PDF ทั้งหมดไว้ก่อน ---
            pdf_list = []
            with pdfplumber.open(pdf_file) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        for row in table:
                            # กรองเฉพาะแถวที่มีเลขประจำตัว (สมมติอยู่คอลัมน์ index 1)
                            if row[1] and str(row[1]).strip().isdigit():
                                pdf_list.append({
                                    'ID': str(row[1]).strip(),
                                    'คะแนน_PDF': row[9],
                                    'เกรด_PDF': row[10]
                                })
            df_pdf_all = pd.DataFrame(pdf_list)

            # --- เตรียมข้อมูล Excel (หาพิกัดแต่ละห้อง) ---
            df_raw = pd.read_excel(excel_file, header=None)
            # หาแถวที่ขึ้นต้นด้วย ม.1/ (ปรับแก้คำค้นหาตรงนี้ให้ตรงกับไฟล์จริง)
            room_indices = df_raw[df_raw.iloc[:, 0].astype(str).str.contains('ม.1/', na=False)].index.tolist()
            room_indices.append(len(df_raw)) # จุดจบไฟล์

            # --- ลูปประมวลผลทีละห้อง ---
            for i in range(len(room_indices) - 1):
                start = room_indices[i]
                end = room_indices[i+1]
                
                # ตัดข้อมูลเฉพาะช่วงห้องนั้น
                df_room_chunk = df_raw.iloc[start:end].reset_index(drop=True)
                room_name = str(df_room_chunk.iloc[0, 0]).strip()

                with st.expander(f"📂 ผลการตรวจสอบ: {room_name}", expanded=True):
                    # 1. ดึงข้อมูลนักเรียนรายคน
                    # เริ่มอ่านจากบรรทัดที่ 3 หลังหัวห้อง (ปรับตัวเลขตามจริง)
                    df_students = df_room_chunk.iloc[3:].copy()
                    # กรองเฉพาะแถวที่คอลัมน์ "เลขประจำตัว" (index 1) เป็นตัวเลข
                    df_students = df_students[df_students.iloc[:, 1].astype(str).str.strip().str.isdigit()]
                    
                    # เลือกคอลัมน์: ID(1), ชื่อ(2), คะแนน(17), เกรด(18)
                    df_students = df_students.iloc[:, [1, 3, 4, 17, 18]]
                    df_students.columns = ['ID', 'ชื่อ', 'นามสกุล', 'คะแนน_Excel', 'เกรด_Excel']
                    df_students['ID'] = df_students['ID'].astype(str).str.strip().str.replace('.0', '', regex=False)

                    # 2. Merge กับ PDF เฉพาะ ID ที่อยู่ในห้องนี้
                    df_final = pd.merge(df_students, df_pdf_all, on='ID', how='left')

                    # Logic การใส่สี (Highlight)
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

                    # แสดงตารางรายคน
                    st.dataframe(df_final.style.apply(highlight_logic, axis=1), use_container_width=True)

                    # 3. เช็คค่าร้อยละท้ายตารางของห้องนี้
                    summary_row = df_room_chunk[df_room_chunk.apply(lambda x: x.astype(str).str.contains('ร้อยละ').any(), axis=1)]
                    if not summary_row.empty:
                        excel_total = summary_row.iloc[0, 17]
                        st.write(f"🎯 ค่าร้อยละรวมของ {room_name} (Excel): **{excel_total}**")
                    else:
                        st.warning(f"⚠️ ไม่พบแถวร้อยละสรุปของ {room_name}")

            st.success("✅ ตรวจสอบครบทุกห้องเรียบร้อยแล้ว!")

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")
