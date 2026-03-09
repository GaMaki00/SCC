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
            pdf_data = []
            with pdfplumber.open(pdf_file) as pdf:
                for page in pdf.pages:
                    # ดึงคำออกมาพร้อมตำแหน่ง (X, Y)
                    words = page.extract_words()
                    
                    # เราจะหาเลขประจำตัว 5 หลักก่อน
                    for i, word in enumerate(words):
                        text = word['text'].strip()
                        
                        # ถ้าเจอเลข 5 หลัก (ID)
                        if text.isdigit() and len(text) == 5:
                            student_id = text
                            score_val = "0"
                            grade_val = "0"
                            
                            # สแกนคำที่อยู่ถัดไปในบรรทัดเดียวกัน (พิกัด Y ใกล้กัน)
                            # ปกติคะแนนรวมจะอยู่ห่างจาก ID ไปทางขวา (X) ประมาณหนึ่ง
                            current_y = word['top']
                            
                            row_words = [w for w in words if abs(w['top'] - current_y) < 3]
                            # เรียงจากซ้ายไปขวา
                            row_words.sort(key=lambda x: x['x0'])
                            
                            # จาก ID (ตัวที่ 1 ในตาราง) 
                            # คะแนนรวมมักจะเป็นตัวเลขลำดับที่ 9-10 ของบรรทัด
                            # เราจะใช้การกรองหาตัวเลขที่อยู่ในช่วงพิกัด X ของคอลัมน์ 'รวม' และ 'เกรด'
                            try:
                                # ค้นหาตัวเลขที่อยู่หลังชื่อ-นามสกุล
                                # ใน ปพ. โรงเรียนเรา คะแนนรวมจะอยู่คอลัมน์ท้ายๆ
                                # เราจะดึงตัวเลขทศนิยมที่อยู่ก่อนตัวเลขเกรด
                                numeric_words = [w['text'] for w in row_words if re.match(r'^\d+\.?\d*$', w['text'])]
                                
                                if len(numeric_words) >= 5:
                                    # คะแนนรวม (ช่องที่ 9 ของแถวข้อมูล)
                                    score_val = numeric_words[9] 
                                    # เกรด (ช่องที่ 10 ของแถวข้อมูล)
                                    grade_val = numeric_words[10]
                                    
                                    pdf_data.append({
                                        'ID': student_id,
                                        'คะแนน_PDF': score_val,
                                        'เกรด_PDF': grade_val
                                    })
                            except:
                                pass
            
            # ลบข้อมูลซ้ำและจัดการ DataFrame
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
