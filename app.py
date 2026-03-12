import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import numpy as np

st.set_page_config(page_title="ระบบจัดการคะแนน - ครูโอม", layout="wide")

st.title("👨‍🏫 ระบบจัดการคะแนนอัตโนมัติ (ครูโอม)")

# 1. ส่วนการอัปโหลดไฟล์ (ไว้ด้านบนสุดเพื่อให้ทั้ง 2 Tabs ใช้ร่วมกันได้)
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("📂 1. เลือกไฟล์ Excel คะแนน", type=['xlsx'])
with col2:
    pdf_file = st.file_uploader("📄 2. เลือกไฟล์ PDF ปพ. (สำหรับตรวจสอบ)", type=['pdf'])

# สร้างหมวดหมู่ (Tabs)
tab1, tab2 = st.tabs(["🔍 ตรวจสอบคะแนน (Excel vs PDF)", "📈 วิเคราะห์สถิติรายห้อง (S.D./Mean)"])

if excel_file:
    # --- ส่วนกลาง: อ่านข้อมูล Excel รอไว้เลย ---
    df_raw = pd.read_excel(excel_file, header=None)
    room_indices = df_raw[df_raw.iloc[:, 0].astype(str).str.contains('ม.1/', na=False)].index.tolist()
    room_indices.append(len(df_raw))

    # --- TAB 1: ตรวจสอบคะแนน ---
    with tab1:
        if not pdf_file:
            st.warning("⚠️ กรุณาอัปโหลดไฟล์ PDF เพื่อเริ่มการตรวจสอบคะแนน")
        else:
            if st.button("🚀 เริ่มตรวจสอบคะแนนทั้งหมด"):
                try:
                    # 1. อ่าน PDF
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

                    # 2. ประมวลผลรายห้อง
                    detailed_results = []
                    summary_dashboard = []

                    for i in range(len(room_indices) - 1):
                        start, end = room_indices[i], room_indices[i+1]
                        df_room_chunk = df_raw.iloc[start:end].reset_index(drop=True)
                        room_name = str(df_room_chunk.iloc[0, 0]).strip()

                        df_students = df_room_chunk.iloc[2:].copy()
                        df_students = df_students[df_students.iloc[:, 1].astype(str).str.strip().str.isdigit()]
                        df_students = df_students.iloc[:, [1, 3, 4, 17, 18]]
                        df_students.columns = ['ID', 'ชื่อ', 'นามสกุล', 'คะแนน_Excel', 'เกรด_Excel']
                        df_students['ID'] = df_students['ID'].astype(str).str.strip().str.replace('.0', '', regex=False)

                        df_final = pd.merge(df_students, df_pdf_all, on='ID', how='left')
                        df_final.index = df_final.index + 1

                        count_ex = len(df_final)
                        count_pdf = df_final['คะแนน_PDF'].notna().sum()
                        summary_row = df_room_chunk[df_room_chunk.apply(lambda x: x.astype(str).str.contains('ร้อยละ').any(), axis=1)]
                        excel_avg = float(summary_row.iloc[0, 17]) if not summary_row.empty else 0
                        pdf_scores = pd.to_numeric(df_final['คะแนน_PDF'].astype(str).str.replace(',', ''), errors='coerce').dropna()
                        pdf_avg = pdf_scores.mean() if not pdf_scores.empty else 0

                        status = "✅ ตรงเป๊ะ" if abs(excel_avg - pdf_avg) <= 0.01 else "❌ ตรวจสอบ"
                        
                        summary_dashboard.append({
                            "ห้อง": room_name,
                            "นร. (Excel/PDF)": f"{count_ex}/{count_pdf}",
                            "ร้อยละ Excel": f"{excel_avg:.2f}",
                            "ร้อยละ PDF": f"{pdf_avg:.2f}",
                            "สถานะ": status
                        })
                        detailed_results.append({"name": room_name, "df": df_final, "ex_avg": excel_avg, "pdf_avg": pdf_avg, "count_ex": count_ex, "count_pdf": count_pdf})

                    # แสดงผล Dashboard
                    st.subheader("📌 สรุปภาพรวมทุกห้อง")
                    df_dash = pd.DataFrame(summary_dashboard)
                    df_dash.index = df_dash.index + 1
                    def color_status(val):
                        return 'background-color: #C6EFCE' if val == "✅ ตรงเป๊ะ" else 'background-color: #FFC7CE'
                    st.table(df_dash.style.applymap(color_status, subset=['สถานะ']))

                    # แสดงรายละเอียดรายห้อง (เหมือนโค้ดเดิมที่ครูใช้งานได้ดี)
                    for i, res in enumerate(detailed_results):
                        with st.expander(f"📂 รายละเอียด: {res['name']}"):
                            st.dataframe(res['df'], use_container_width=True)
                except Exception as e:
                    st.error(f"Error: {e}")

    # --- TAB 2: วิเคราะห์สถิติรายห้อง ---
    with tab2:
        st.subheader("📊 ข้อมูลทางสถิติแยกตามห้อง (อ้างอิงจาก Excel)")
        if st.button("📈 คำนวณสถิติ"):
            all_stats = []
            for i in range(len(room_indices) - 1):
                start, end = room_indices[i], room_indices[i+1]
                df_room = df_raw.iloc[start:end].reset_index(drop=True)
                room_name = str(df_room.iloc[0, 0]).strip()

                df_scores = df_room.iloc[2:].copy()
                df_scores = df_scores[df_scores.iloc[:, 1].astype(str).str.strip().str.isdigit()]
                scores = pd.to_numeric(df_scores.iloc[:, 17], errors='coerce').dropna()

                if not scores.empty:
                    all_stats.append({
                        "ห้อง": room_name,
                        "นร.": len(scores),
                        "Min": f"{scores.min():.2f}",
                        "Max": f"{scores.max():.2f}",
                        "Mean": f"{scores.mean():.2f}",
                        "S.D.": f"{scores.std():.2f}"
                    })
            
            df_stat_final = pd.DataFrame(all_stats)
            df_stat_final.index = df_stat_final.index + 1
            st.table(df_stat_final)
            
            # ปุ่มโหลด CSV สถิติ
            csv = df_stat_final.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 โหลดไฟล์สถิติ (CSV)", csv, "stat_summary.csv", "text/csv")
            # --- ส่วนคำนวณสรุปเป็นข้อความ (วางต่อจาก st.table) ---
            
            # 1. คำนวณภาพรวมทั้งสายชั้น
            total_students = df_stat_final['นร.'].astype(int).sum()
            
            # แปลงค่ากลับเป็นตัวเลขเพื่อคำนวณ
            df_calc = df_stat_final.copy()
            for col in ['Mean', 'S.D.', 'Max', 'Min']:
                df_calc[col] = pd.to_numeric(df_calc[col])
                
            total_mean = df_calc['Mean'].mean()
            total_sd = df_calc['S.D.'].mean()
            overall_max = df_calc['Max'].max()
            overall_min = df_calc['Min'].min()
            
            # 2. จัดอันดับห้องที่ได้คะแนนสูงสุด 3 อันดับ
            top_rooms = df_calc.sort_values(by='Mean', ascending=False).head(3)
            room_names = top_rooms['ห้อง'].str.extract(r'(ม\.1/\d+)')[0].tolist()
            
            # เพื่อกัน Error กรณีมีห้องไม่ถึง 3 ห้อง
            while len(room_names) < 3:
                room_names.append(" - ")

            # 3. แสดงผลเป็นข้อความสรุป
            st.divider()
            st.subheader("📝 บทสรุปสำหรับรายงาน")
            
            report_text = f"""
            จากตารางพบว่า นักเรียนที่เรียนรายวิชาเทคโนโลยี (วิทยาการคำนวณ) 
            รหัสวิชา ว21112 จำนวน **{total_students}** คน 
            ได้คะแนนเฉลี่ย **{total_mean:.2f}** ส่วนเบี่ยงเบนมาตรฐาน **{total_sd:.2f}** ค่าสูงสุด **{overall_max:.2f}** คะแนน 
            ค่าต่ำสุด **{overall_min:.2f}** คะแนน 
            โดยห้องที่ได้คะแนนเฉลี่ยสูงสุดได้แก่ **{room_names[0]}** และรองลงมาคือห้อง **{room_names[1]}**, และ **{room_names[2]}**
            """
            
            st.success(report_text)
            
            # เพิ่มปุ่มก๊อปปี้ข้อความ
            st.text_area("ก๊อปปี้ข้อความด้านล่างนี้ไปใช้:", value=report_text.replace("**", ""), height=150)
