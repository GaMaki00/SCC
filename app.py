import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

st.set_page_config(page_title="ระบบตรวจคะแนน ปพ. ครูโอม", layout="wide")

st.title("📊 ระบบตรวจสอบคะแนน (รองรับ PDF 2 หน้าต่อห้อง)")

# 1. ส่วนการอัปโหลดไฟล์
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. เลือกไฟล์ Excel", type=['xlsx'])
with col2:
    pdf_file = st.file_uploader("2. เลือกไฟล์ PDF", type=['pdf'])

if excel_file and pdf_file:
    if st.button("🚀 เริ่มประมวลผล"):
        try:
            # --- 1. อ่านข้อมูลรายคนจาก PDF ทั้งหมดเก็บไว้ก่อน ---
            pdf_list = []
            all_pages_text = [] # เก็บข้อความแยกตามหน้า
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

                  # --- แก้ไขเฉพาะส่วนการค้นหาร้อยละใน PDF (บรรทัดที่ประมาณ 90 เป็นต้นไป) ---

                    # --- ส่วนที่แก้ไข: ดึงร้อยละ PDF โดยนับ 2 หน้าต่อ 1 ห้อง ---
                    st.markdown("---")
                    col_sum1, col_sum2 = st.columns(2)
                    
                    # 1. หาใน Excel
                    summary_row = df_room_chunk[df_room_chunk.apply(lambda x: x.astype(str).str.contains('ร้อยละ').any(), axis=1)]
                    if not summary_row.empty:
                        # พยายามดึงจากคอลัมน์ที่ 17 (R) ถ้าเป็นตัวเลข
                        raw_val = summary_row.iloc[0, 17]
                        try:
                            excel_total = f"{float(raw_val):.2f}"
                        except:
                            excel_total = str(raw_val)
                    else:
                        excel_total = "N/A"
                    
                    # 2. หาใน PDF (เจาะจงหน้าของห้องนั้นๆ)
                    page_start = i * 2
                    page_end = page_start + 2
                    pdf_room_text = "\n".join(all_pages_text[page_start:page_end])
                    
                    # ปรับ Regex ให้ครอบคลุม "ผลการเรียนเฉลี่ยร้อยละ" และดึงตัวเลขทศนิยมที่ตามมา
                    # พยายามหาคำว่า "ผลการเรียนเฉลี่ยร้อยละ" หรือ "เฉลี่ยร้อยละ"
                    match = re.search(r"(?:ผลการเรียนเฉลี่ยร้อยละ|เฉลี่ยร้อยละ|ร้อยละ)\s*[:]*\s*(\d+\.\d+)", pdf_room_text)
                    
                    if match:
                        pdf_total = match.group(1)
                    else:
                        # ถ้ายังหาไม่เจอ ลองหาตัวเลขทศนิยมที่อยู่ใกล้ๆ คำว่า "ผลการเรียน"
                        match_backup = re.search(r"ผลการเรียนเฉลี่ยร้อยละ\s+(\d+\.\d+)", pdf_room_text)
                        pdf_total = match_backup.group(1) if match_backup else "ไม่พบข้อมูล"

                    with col_sum1:
                        st.metric(f"ร้อยละ {room_name} (Excel)", excel_total)
                    with col_sum2:
                        st.metric(f"ร้อยละ {room_name} (PDF)", pdf_total)
