import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

st.set_page_config(page_title="ระบบตรวจคะแนน ปพ. ครูโอม", layout="wide")

st.title("📊 ระบบตรวจสอบคะแนน (เวอร์ชันแก้ไข Syntax)")

# 1. ส่วนการอัปโหลดไฟล์
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. เลือกไฟล์ Excel", type=['xlsx'])
with col2:
    pdf_file = st.file_uploader("2. เลือกไฟล์ PDF", type=['pdf'])

if excel_file and pdf_file:
    if st.button("🚀 เริ่มประมวลผล"):
        try:
            # --- 1. อ่านข้อมูลรายคนจาก PDF ทั้งหมด ---
            pdf_list = []
            all_pages_text = []
            with pdfplumber.open(pdf_file) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    all_pages_text.append(text if text else "")
                    
                    table = page.extract_table()
                    if table:
                        for row in table:
                            if len(row) > 1 and str(row[1]).strip().isdigit():
                                pdf_list.append({
                                    'ID': str(row[1]).strip(),
                                    'คะแนน_PDF': row[9],
                                    'เกรด_PDF': row[10]
                                })
            df_pdf_all = pd.DataFrame(pdf_list)

            # --- 2. เตรียมข้อมูล Excel (แยกห้อง) ---
            df_raw = pd.read_excel(excel_file, header=None)
            room_indices = df_raw[df_raw.iloc[:, 0].astype(str).str.contains('ม.1/', na=False)].index.tolist()
            room_indices.append(len(df_raw))

            # --- 3. ลูปประมวลผลทีละห้อง ---
            for i in range(len(room_indices) - 1):
                start = room_indices[i]
                end = room_indices[i+1]
                
                df_room_chunk = df_raw.iloc[start:end].reset_index(drop=True)
                room_name = str(df_room_chunk.iloc[0, 0]).strip()

                with st.expander(f"📂 ผลการตรวจสอบ: {room_name}", expanded=True):
                    # กรองข้อมูลนักเรียน
                    df_students = df_room_chunk.iloc[3:].copy()
                    df_students = df_students[df_students.iloc[:, 1].astype(str).str.strip().str.isdigit()]
                    df_students = df_students.iloc[:, [1, 3, 4, 17, 18]]
                    df_students.columns = ['ID', 'ชื่อ', 'นามสกุล', 'คะแนน_Excel', 'เกรด_Excel']
                    df_students['ID'] = df_students['ID'].astype(str).str.strip().str.replace('.0', '', regex=False)

                    # Merge กับ PDF
                    df_final = pd.merge(df_students, df_pdf_all, on='ID', how='left')

                    # แสดงตาราง
                    def highlight_logic(row):
                        def safe_float(val):
                            try: return float(str(val).strip().replace(',', ''))
                            except: return None
                        s_ex, s_pdf = safe_float(row['คะแนน_Excel']), safe_float(row['คะแนน_PDF'])
                        g_ex, g_pdf = str(row['เกรด_Excel']).strip(), str(row['เกรด_PDF']).strip()
                        if s_ex is not None and s_pdf is not None:
                            bg = 'background-color: #C6EFCE' if (s_ex == s_pdf and g_ex == g_pdf) else 'background-color: #FFC7CE'
                        else: bg = 'background-color: #FFEB9C'
                        return [bg] * len(row)

                    st.dataframe(df_final.style.apply(highlight_logic, axis=1), use_container_width=True)

                 st.markdown("---")
                    col_sum1, col_sum2, col_sum3 = st.columns(3)
                    
                    # 1. ดึงจาก Excel (แถวที่เขียนว่าร้อยละ)
                    summary_row = df_room_chunk[df_room_chunk.apply(lambda x: x.astype(str).str.contains('ร้อยละ').any(), axis=1)]
                    if not summary_row.empty:
                        try:
                            excel_val = float(summary_row.iloc[0, 17])
                            excel_total_str = f"{excel_val:.2f}"
                        except:
                            excel_total_str = str(summary_row.iloc[0, 17])
                    else:
                        excel_total_str = "N/A"
                    
                    # 2. คำนวณจาก PDF เอง (Mean of all scores in this room)
                    # แปลงคะแนน PDF ให้เป็นตัวเลข
                    pdf_scores_numeric = pd.to_numeric(df_final['คะแนน_PDF'].astype(str).str.replace(',', ''), errors='coerce')
                    pdf_scores_numeric = pdf_scores_numeric.dropna()
                    
                    if not pdf_scores_numeric.empty:
                        pdf_calc_avg = pdf_scores_numeric.mean()
                        pdf_total_str = f"{pdf_calc_avg:.2f}"
                    else:
                        pdf_total_str = "0.00"

                    with col_sum1:
                        st.metric(f"ร้อยละ {room_name} (Excel)", excel_total_str)
                    with col_sum2:
                        st.metric(f"ร้อยละ {room_name} (PDF - คำนวณให้ใหม่)", pdf_total_str)
                    with col_sum3:
                        # ตรวจสอบความถูกต้อง (ยอมให้ต่างกันได้เล็กน้อยจากทศนิยม 0.01)
                        try:
                            diff = abs(float(excel_total_str) - float(pdf_total_str))
                            if diff <= 0.01:
                                st.success("✅ ค่าเฉลี่ยตรงกัน")
                            else:
                                st.error(f"❌ ต่างกัน {diff:.2f}")
                        except:
                            st.warning("⚠️ ไม่สามารถเทียบค่าได้")

            st.success("✅ ตรวจสอบและคำนวณครบทุกห้องแล้วครับครูโอม!")

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
