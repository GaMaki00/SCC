import streamlit as st
import pandas as pd
import pdfplumber
import re
import io

st.set_page_config(page_title="ระบบตรวจคะแนน ปพ. ครูโอม", layout="wide")

st.title("📊 ระบบตรวจสอบคะแนน (พร้อมสรุปภาพรวมทุกห้อง)")

# 1. ส่วนการอัปโหลดไฟล์
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. เลือกไฟล์ Excel", type=['xlsx'])
with col2:
    pdf_file = st.file_uploader("2. เลือกไฟล์ PDF", type=['pdf'])

if excel_file and pdf_file:
    if st.button("🚀 เริ่มประมวลผลและสร้างสรุปภาพรวม"):
        try:
            # --- 1. อ่านข้อมูล PDF ทั้งหมด ---
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
                                    'เกรด_PDF': nums[-1] if len(nums) >= 1 else nums[-1]
                                })
            df_pdf_all = pd.DataFrame(pdf_data).drop_duplicates(subset=['ID'])

            # --- 2. เตรียมข้อมูล Excel และตัวแปรเก็บสรุป ---
            df_raw = pd.read_excel(excel_file, header=None)
            room_indices = df_raw[df_raw.iloc[:, 0].astype(str).str.contains('ม.1/', na=False)].index.tolist()
            room_indices.append(len(df_raw))
            
            summary_dashboard = [] # เก็บข้อมูลเพื่อทำ Dashboard บนสุด
            detailed_results = []  # เก็บผลลัพธ์รายห้องเพื่อแสดงด้านล่าง

            # --- 3. ประมวลผลล่วงหน้าเพื่อเอาข้อมูลมาทำ Dashboard ---
            for i in range(len(room_indices) - 1):
                start, end = room_indices[i], room_indices[i+1]
                df_room_chunk = df_raw.iloc[start:end].reset_index(drop=True)
                room_name = str(df_room_chunk.iloc[0, 0]).strip()

                # ดึงเด็กรายคน
                df_students = df_room_chunk.iloc[2:].copy()
                df_students = df_students[df_students.iloc[:, 1].astype(str).str.strip().str.isdigit()]
                df_students = df_students.iloc[:, [1, 3, 4, 17, 18]]
                df_students.columns = ['ID', 'ชื่อ', 'นามสกุล', 'คะแนน_Excel', 'เกรด_Excel']
                df_students['ID'] = df_students['ID'].astype(str).str.strip().str.replace('.0', '', regex=False)

                df_final = pd.merge(df_students, df_pdf_all, on='ID', how='left')
                df_final.index = df_final.index + 1

                # คำนวณค่าต่างๆ ของห้องนี้
                count_ex = len(df_final)
                count_pdf = df_final['คะแนน_PDF'].notna().sum()
                
                summary_row = df_room_chunk[df_room_chunk.apply(lambda x: x.astype(str).str.contains('ร้อยละ').any(), axis=1)]
                excel_avg = float(summary_row.iloc[0, 17]) if not summary_row.empty else 0
                
                pdf_scores = pd.to_numeric(df_final['คะแนน_PDF'].astype(str).str.replace(',', ''), errors='coerce').dropna()
                pdf_avg = pdf_scores.mean() if not pdf_scores.empty else 0
                
                status = "✅ ตรงเป๊ะ" if abs(excel_avg - pdf_avg) <= 0.00 else "❌ ตรวจสอบ"
                
                summary_dashboard.append({
                    "ห้อง": room_name,
                    "นร. (Excel/PDF)": f"{count_ex}/{count_pdf}",
                    "ร้อยละ Excel": f"{excel_avg:.2f}",
                    "ร้อยละ PDF": f"{excel_avg:.2f}",
                    "สถานะ": status
                })
                
                detailed_results.append({
                    "name": room_name,
                    "df": df_final,
                    "ex_avg": excel_avg,
                    "pdf_avg": pdf_avg,
                    "count_ex": count_ex,
                    "count_pdf": count_pdf,
                    "chunk": df_room_chunk
                })

            # --- 4. แสดงผล Dashboard บนสุด ---
            st.subheader("📌 สรุปภาพรวมทุกห้อง")
            df_dash = pd.DataFrame(summary_dashboard)
            # เริ่มต้นที่เลข 1 แทนเลข 0
            df_dash.index = df_dash.index + 1
            # ตกแต่งสีใน Dashboard
            def color_status(val):
                color = '#C6EFCE' if val == "✅ ตรงกัน" else '#FFC7CE'
                return f'background-color: {color}'
            
            st.table(df_dash.style.applymap(color_status, subset=['สถานะ']))

            # --- 5. แสดงผลรายละเอียดรายห้องด้านล่าง ---
            st.divider()
            for i, res in enumerate(detailed_results):
                with st.expander(f"📂 รายละเอียด: {res['name']} ({res['status'] if 'status' in res else ''})"):
                    # ไฮไลท์สีในตารางเว็บ
                    def apply_highlight(row):
                        try:
                            s_ex1 = float(str(row['คะแนน_Excel']).replace(',', '').strip())
                            s_ex2 = float(str(row['เกรด_Excel']).replace(',', '').strip())
                            
                            s_pdf1 = float(str(row['คะแนน_PDF']).replace(',', '').strip())
                            s_pdf2 = float(str(row['เกรด_PDF']).replace(',', '').strip())
                            if s_ex1 == s_pdf1 and s_ex2 == s_pdf2:
                                bg = 'background-color: #C6EFCE'
                            else:
                                bg = 'background-color: #FFC7CE'
                        except:
                            bg = 'background-color: #FFEB9C'
                        return [bg] * len(row)

                    st.dataframe(res['df'].style.apply(apply_highlight, axis=1), use_container_width=True)
                    st.markdown("---")
                    c1, c2, c3 = st.columns(3)
                    
                    with c1: 
                        st.metric("จำนวนนักเรียน", f"{res['count_ex']} / {res['count_pdf']}")
                    with c2: 
                        st.metric("ร้อยละ Excel", f"{res['ex_avg']:.2f}")
                    with c3: 
                        diff = res['pdf_avg'] - res['ex_avg']
                        
                        # ใช้สีแดงสำหรับตัวเลขหลัก
                        if abs(diff) > 0.01:
                            text_color = "#FF4B4B" # สีแดง
                            status_text = "ไม่ตรง"
                        else:
                            text_color = "#00D166" # สีเขียว
                            status_text = "ตรงกัน"

                        display_html = f"""
                            <div style="line-height: 1.2;">
                                <p style="font-size: 16px; margin-bottom: 0px; font-weight: bold;">ร้อยละ PDF (คำนวณ)</p>
                                <p style="font-size: 48px; font-weight: bold; color: {text_color}; margin: 5px 0px;">{res['pdf_avg']:.2f}</p>
                                <p style="color: {text_color}; font-size: 16px; font-weight: bold;">
                                    {'↑' if diff >= 0 else '↓'} {abs(diff):.2f} ({status_text})
                                </p>
                            </div>
                        """
                        st.markdown(display_html, unsafe_allow_html=True)
                    # ปุ่มดาวน์โหลด
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        res['df'].to_excel(writer, index=True, sheet_name='Summary')
                        workbook, worksheet = writer.book, writer.sheets['Summary']
                        green = workbook.add_format({'bg_color': '#C6EFCE', 'border': 1})
                        red   = workbook.add_format({'bg_color': '#FFC7CE', 'border': 1})
                        for row_num in range(len(res['df'])):
                            try:
                                if float(res['df'].iloc[row_num]['คะแนน_Excel']) == float(res['df'].iloc[row_num]['คะแนน_PDF']):
                                    fmt = green
                                else: fmt = red
                            except: fmt = None
                            worksheet.set_row(row_num + 1, None, fmt)
                    
                    st.download_button(
                        label=f"📥 ดาวน์โหลด Excel {res['name']}",
                        data=output.getvalue(),
                        file_name=f"ตรวจคะแนน_{res['name']}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"btn_dl_{i}"
                    )

            st.success("✅ ตรวจสอบครบถ้วนทุกห้องแล้วครับครูโอม!")

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
