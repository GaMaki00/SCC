import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

st.set_page_config(page_title="ระบบตรวจคะแนน ปพ. ครูโอม", layout="wide")

st.title("📊 ระบบตรวจสอบคะแนน (ดึงข้อมูลละเอียด + โหลด Excel)")

# 1. ส่วนการอัปโหลดไฟล์
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. เลือกไฟล์ Excel", type=['xlsx'])
with col2:
    pdf_file = st.file_uploader("2. เลือกไฟล์ PDF", type=['pdf'])

if excel_file and pdf_file:
    if st.button("🚀 เริ่มประมวลผล"):
        try:
            # --- 1. อ่านข้อมูลจาก PDF (พิกัด X-Y) ---
            pdf_data = []
            with pdfplumber.open(pdf_file) as pdf:
                for page in pdf.pages:
                    words = page.extract_words()
                    lines = {}
                    for w in words:
                        y = round(w['top'], 0)
                        if y not in lines: lines[y] = []
                        lines[y].append(w)
                    
                    for y in sorted(lines.keys()):
                        row = sorted(lines[y], key=lambda x: x['x0'])
                        text_line = " ".join([w['text'] for w in row])
                        match_id = re.search(r'(\d{5})', text_line)
                        if match_id:
                            student_id = match_id.group(1)
                            nums = [w['text'] for w in row if re.match(r'^\d+\.?\d*$', w['text']) and w['x0'] > 250]
                            if len(nums) >= 2:
                                pdf_data.append({
                                    'ID': student_id,
                                    'คะแนน_PDF': nums[-3] if len(nums) >= 3 else nums[-2],
                                    'เกรด_PDF': nums[-2] if len(nums) >= 3 else nums[-1]
                                })
            df_pdf_all = pd.DataFrame(pdf_data).drop_duplicates(subset=['ID'])

            # --- 2. อ่าน Excel (ดึงคนแรกกลับมา) ---
            df_raw = pd.read_excel(excel_file, header=None)
            room_indices = df_raw[df_raw.iloc[:, 0].astype(str).str.contains('ม.1/', na=False)].index.tolist()
            room_indices.append(len(df_raw))

            for i in range(len(room_indices) - 1):
                start, end = room_indices[i], room_indices[i+1]
                df_room_chunk = df_raw.iloc[start:end].reset_index(drop=True)
                room_name = str(df_room_chunk.iloc[0, 0]).strip()

                with st.expander(f"📂 ผลการตรวจสอบ: {room_name}", expanded=True):
                    # iloc[2:] เพื่อดึงเลขที่ 1
                    df_students = df_room_chunk.iloc[2:].copy()
                    df_students = df_students[df_students.iloc[:, 1].astype(str).str.strip().str.isdigit()]
                    df_students = df_students.iloc[:, [1, 3, 4, 17, 18]]
                    df_students.columns = ['ID', 'ชื่อ', 'นามสกุล', 'คะแนน_Excel', 'เกรด_Excel']
                    df_students['ID'] = df_students['ID'].astype(str).str.strip().str.replace('.0', '', regex=False)

                    # Merge & Re-index
                    df_final = pd.merge(df_students, df_pdf_all, on='ID', how='left')
                    df_final.index = df_final.index + 1

                    # ฟังก์ชันไฮไลท์ในหน้าเว็บ
                    def apply_highlight(row):
                        try:
                            s_ex = float(str(row['คะแนน_Excel']).replace(',', ''))
                            s_pdf = float(str(row['คะแนน_PDF']).replace(',', ''))
                            bg = 'background-color: #C6EFCE' if s_ex == s_pdf else 'background-color: #FFC7CE'
                        except: bg = 'background-color: #FFEB9C'
                        return [bg] * len(row)

                    st.dataframe(df_final.style.apply(apply_highlight, axis=1), use_container_width=True)

                    # --- ส่วนสร้างไฟล์ Excel ให้ดาวน์โหลด ---
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_final.to_excel(writer, index=True, sheet_name='Summary')
                        workbook, worksheet = writer.book, writer.sheets['Summary']
                        # Format สี
                        green = workbook.add_format({'bg_color': '#C6EFCE', 'border': 1})
                        red   = workbook.add_format({'bg_color': '#FFC7CE', 'border': 1})
                        
                        for row_num in range(len(df_final)):
                            try:
                                if float(df_final.iloc[row_num]['คะแนน_Excel']) == float(df_final.iloc[row_num]['คะแนน_PDF']):
                                    fmt = green
                                else: fmt = red
                            except: fmt = None
                            worksheet.set_row(row_num + 1, None, fmt)
                    
                    st.download_button(
                        label=f"📥 ดาวน์โหลด Excel {room_name}",
                        data=output.getvalue(),
                        file_name=f"ตรวจคะแนน_{room_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"btn_{i}"
                    )

            st.success("✅ ประมวลผลสำเร็จ!")

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
            # --- ส่วนคำนวณร้อยละสรุป (เพิ่มกลับเข้ามาให้แล้วครับ) ---
                    st.markdown("---")
                    c1, c2, c3 = st.columns(3)
                    
                    # 1. จำนวนนักเรียน
                    count_excel = len(df_final)
                    count_pdf = df_final['คะแนน_PDF'].notna().sum()
                    with c1:
                        st.metric("จำนวนนักเรียน (Excel / PDF)", f"{count_excel} / {count_pdf}")

                    # 2. ร้อยละจาก Excel (ดึงจากแถวที่มีคำว่าร้อยละ)
                    summary_row = df_room_chunk[df_room_chunk.apply(lambda x: x.astype(str).str.contains('ร้อยละ').any(), axis=1)]
                    excel_avg = float(summary_row.iloc[0, 17]) if not summary_row.empty else 0
                    with c2:
                        st.metric(f"ร้อยละ Excel", f"{excel_avg:.2f}")

                    # 3. ร้อยละ PDF (คำนวณใหม่จากเด็กรายคน)
                    pdf_scores = pd.to_numeric(df_final['คะแนน_PDF'].astype(str).str.replace(',', ''), errors='coerce').dropna()
                    pdf_avg = pdf_scores.mean() if not pdf_scores.empty else 0
                    with c3:
                        # เช็คความต่าง ถ้าไม่เกิน 0.02 ให้ถือว่าปกติ (สีเขียว)
                        diff = pdf_avg - excel_avg
                        color = "normal" if abs(diff) <= 0.02 else "inverse"
                        st.metric(f"ร้อยละ PDF (คำนวณ)", f"{pdf_avg:.2f}", delta=f"{diff:.2f}", delta_color=color)

                    # --- ส่วนสร้างไฟล์ Excel ให้ดาวน์โหลด (มีสีไฮไลท์) ---
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df_final.to_excel(writer, index=True, sheet_name='Summary')
                        workbook = writer.book
                        worksheet = writer.sheets['Summary']
                        
                        # กำหนดสี
                        green = workbook.add_format({'bg_color': '#C6EFCE', 'border': 1})
                        red   = workbook.add_format({'bg_color': '#FFC7CE', 'border': 1})
                        
                        for row_num in range(len(df_final)):
                            try:
                                val_ex = float(df_final.iloc[row_num]['คะแนน_Excel'])
                                val_pdf = float(df_final.iloc[row_num]['คะแนน_PDF'])
                                fmt = green if val_ex == val_pdf else red
                            except:
                                fmt = None
                            worksheet.set_row(row_num + 1, None, fmt)
                    
                    st.download_button(
                        label=f"📥 ดาวน์โหลดไฟล์สรุป {room_name}",
                        data=output.getvalue(),
                        file_name=f"ตรวจคะแนน_{room_name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"btn_dl_{i}"
                    )
