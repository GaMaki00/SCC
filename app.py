import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

st.set_page_config(page_title="ระบบตรวจคะแนน ปพ. ครูโอม", layout="wide")

st.title("📊 ระบบตรวจสอบคะแนน (ดึงข้อมูลแบบละเอียด)")
st.info("💡 ระบบจะดึงข้อมูลเด็กทุกคนจาก PDF และนับจำนวนคนให้ตรวจสอบ")

# 1. ส่วนการอัปโหลดไฟล์
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. เลือกไฟล์ Excel", type=['xlsx'])
with col2:
    pdf_file = st.file_uploader("2. เลือกไฟล์ PDF", type=['pdf'])

if excel_file and pdf_file:
    if st.button("🚀 เริ่มประมวลผลและนับจำนวนคน"):
        try:
            # --- 1. อ่านข้อมูล PDF ทั้งหมด (เปลี่ยนเป็นวิธีสแกนบรรทัดต่อบรรทัด) ---
            pdf_data = []
            with pdfplumber.open(pdf_file) as pdf:
                for page in pdf.pages:
                    # ดึงข้อมูลจากตาราง (วิธีหลัก)
                    tables = page.extract_tables()
                    for table in tables:
                        for row in table:
                            # กรองแถวที่มีข้อมูลคะแนน (คอลัมน์ที่ 1=ID, 9=คะแนน, 10=เกรด)
                            if len(row) > 10:
                                s_id = str(row[1]).strip()
                                if s_id.isdigit() and len(s_id) >= 4:
                                    pdf_data.append({
                                        'ID': s_id,
                                        'คะแนน_PDF': row[9],
                                        'เกรด_PDF': row[10]
                                    })
                    
                    # วิธีสำรอง: ดึงข้อความดิบ (เผื่อตารางหน้า 2 มันแตก)
                    text = page.extract_text()
                    if text:
                        # หาแพทเทิร์น: เลขประจำตัว(5หลัก) + ช่องว่าง + ชื่อ + ตัวเลขคะแนน
                        # ส่วนนี้จะช่วยเก็บคนที่หลุดจากตาราง
                        lines = text.split('\n')
                        for line in lines:
                            # Regex หาเลขประจำตัว 5 หลัก และคะแนนช่วงท้ายบรรทัด
                            match = re.search(r'^(\d{5})\s+.*?\s+(\d+\.?\d*)\s+([0-4]\.?[0-5]*)', line)
                            if match:
                                s_id = match.group(1)
                                # เช็คว่าไม่ซ้ำกับที่ดึงจากตารางไปแล้ว
                                if not any(d['ID'] == s_id for d in pdf_data):
                                    pdf_data.append({
                                        'ID': s_id,
                                        'คะแนน_PDF': match.group(2),
                                        'เกรด_PDF': match.group(3)
                                    })
            
            df_pdf_all = pd.DataFrame(pdf_data).drop_duplicates(subset=['ID'])

            # --- 2. เตรียมข้อมูล Excel ---
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
                    df_students = df_room_chunk.iloc[3:].copy()
                    df_students = df_students[df_students.iloc[:, 1].astype(str).str.strip().str.isdigit()]
                    df_students = df_students.iloc[:, [1, 3, 4, 17, 18]]
                    df_students.columns = ['ID', 'ชื่อ', 'นามสกุล', 'คะแนน_Excel', 'เกรด_Excel']
                    df_students['ID'] = df_students['ID'].astype(str).str.strip().str.replace('.0', '', regex=False)

                    # Merge
                    df_final = pd.merge(df_students, df_pdf_all, on='ID', how='left')
                    
                    # นับจำนวนเด็ก
                    count_excel = len(df_final)
                    count_pdf = df_final['คะแนน_PDF'].notna().sum()

                    # แสดงตาราง
                    st.dataframe(df_final, use_container_width=True)

                    # สรุปท้ายห้อง
                    st.markdown("---")
                    c1, c2, c3 = st.columns(3)
                    
                    with c1:
                        st.metric("จำนวนนักเรียน (Excel / PDF)", f"{count_excel} / {count_pdf}")
                    
                    # คำนวณร้อยละ
                    pdf_scores = pd.to_numeric(df_final['คะแนน_PDF'].astype(str).str.replace(',', ''), errors='coerce').dropna()
                    pdf_avg = pdf_scores.mean() if not pdf_scores.empty else 0
                    
                    summary_row = df_room_chunk[df_room_chunk.apply(lambda x: x.astype(str).str.contains('ร้อยละ').any(), axis=1)]
                    excel_avg = float(summary_row.iloc[0, 17]) if not summary_row.empty else 0

                    with c2:
                        st.metric(f"ร้อยละ Excel", f"{excel_avg:.2f}")
                    with c3:
                        color = "normal" if abs(excel_avg - pdf_avg) <= 0.01 else "inverse"
                        st.metric(f"ร้อยละ PDF (คำนวณ)", f"{pdf_avg:.2f}", delta=f"{pdf_avg-excel_avg:.2f}", delta_color=color)

                    if count_excel != count_pdf:
                        st.warning(f"⚠️ จำนวนเด็กไม่เท่ากัน! ขาดไป {count_excel - count_pdf} คน (ลองเช็คหน้า 2 ของ PDF ห้องนี้)")

            st.success("✅ ประมวลผลเสร็จสิ้น")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
