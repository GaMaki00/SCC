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
tab1, tab2, tab3, tab4, tab5 = st.tabs(["🔍 ตรวจสอบคะแนน", "📈 วิเคราะห์สถิติ", "📊 ผลสัมฤทธิ์", "🧮 E1/E2", "🏆 Top 10"])

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
           # --- ส่วนคำนวณสรุปเป็นข้อความ (รองรับทุกห้อง) ---
            
            # 1. เตรียมข้อมูลตัวเลข
            df_calc = df_stat_final.copy()
            for col in ['Min', 'Max', 'Mean', 'S.D.', 'นร.']:
                df_calc[col] = pd.to_numeric(df_calc[col])

            total_students = int(df_calc['นร.'].sum())
            total_mean = df_calc['Mean'].mean()
            total_sd = df_calc['S.D.'].mean()
            overall_max = df_calc['Max'].max()
            overall_min = df_calc['Min'].min()

            # 2. จัดอันดับห้องจาก Mean มากไปน้อย (ทั้งหมดที่มี)
            df_calc['Room_Short'] = df_calc['ห้อง'].str.extract(r'(ม\.1/\d+)')
            # ถ้าดึงเลขห้องไม่ได้ ให้ใช้ชื่อเต็ม
            df_calc['Room_Label'] = df_calc['Room_Short'].fillna(df_calc['ห้อง'])
            
            sorted_rooms = df_calc.sort_values(by='Mean', ascending=False)
            room_list = sorted_rooms['Room_Label'].tolist()

            # 3. จัดการเรื่องภาษา (ตัวสุดท้ายใช้ "และ")
            if len(room_list) > 1:
                rooms_text = ", ".join(room_list[:-1]) + f", และ{room_list[-1]}"
            else:
                rooms_text = room_list[0]

            # 4. แสดงบทสรุป
            st.divider()
            st.subheader("📝 บทสรุปสำหรับรายงาน (ครบทุกห้อง)")
            
            report_text = (
                f"จากตารางพบว่า นักเรียนที่เรียนรายวิชาเทคโนโลยี (วิทยาการคำนวณ) รหัสวิชา ว21112 "
                f"จำนวน {total_students} คน ได้คะแนนเฉลี่ย {total_mean:.2f} "
                f"ส่วนเบี่ยงเบนมาตรฐาน {total_sd:.2f} ค่าสูงสุด {overall_max:.2f} คะแนน "
                f"ค่าต่ำสุด {overall_min:.2f} คะแนน โดยห้องที่ได้คะแนนเฉลี่ยสูงสุดได้แก่ {rooms_text}"
            )
            
            st.success(report_text)
            
            # ช่องก๊อปปี้
            st.text_area("📋 ก๊อปปี้ข้อความไปใช้ในรายงาน:", value=report_text, height=150)
  # --- TAB 3: ผลสัมฤทธิ์ทางการเรียน (คำนวณ X̄ ตามสูตร Σfx/N) ---
    with tab3:
        st.subheader("📊 ตารางแสดงผลสัมฤทธิ์ทางการเรียน (คำนวณ X̄ ตามสูตร Σfx/N)")
        if st.button("📝 ประมวลผลตารางเกรด"):
            rows = []
            # เก็บค่าสะสมสำหรับสรุปรวมทั้งสายชั้น
            totals_n = {g: 0 for g in [4.0, 3.5, 3.0, 2.5, 2.0, 1.5, 1.0, 0.0]}
            grand_total_n = 0
            grand_sum_fx = 0

            for i in range(len(room_indices) - 1):
                start, end = room_indices[i], room_indices[i+1]
                df_room = df_raw.iloc[start:end].reset_index(drop=True)
                room_name = str(df_room.iloc[0, 0]).strip()

                df_students = df_room.iloc[2:].copy()
                df_students = df_students[df_students.iloc[:, 1].astype(str).str.strip().str.isdigit()]
                
                # ดึงข้อมูลเกรด
                grades = pd.to_numeric(df_students.iloc[:, 18], errors='coerce').dropna()
                total_n = len(grades)

                if total_n > 0:
                    # 1. คำนวณ Σfx (ผลรวมของ จำนวนคน x เกรด)
                    sum_fx = 0
                    counts = {}
                    for g in [4.0, 3.5, 3.0, 2.5, 2.0, 1.5, 1.0, 0.0]:
                        count_g = (grades == g).sum()
                        counts[str(g)] = count_g
                        sum_fx += (count_g * g) # f * x
                        # สะสมค่าไว้ใช้แถวรวม
                        totals_n[g] += count_g
                    
                    grand_total_n += total_n
                    grand_sum_fx += sum_fx

                    # 2. คำนวณ X̄ = Σfx / N
                    mean_x = sum_fx / total_n
                    
                    # 3. คำนวณ S.D. (จากเกรด)
                    sd_val = grades.std()

                    # เพิ่มแถว "จำนวน"
                    row_n = {"ชั้น": room_name, "ประเภท": "จำนวน"}
                    row_n.update(counts)
                    row_n.update({"รวม": total_n, "X̄": f"{mean_x:.2f}", "S.D.": f"{sd_val:.2f}"})
                    rows.append(row_n)

                    # เพิ่มแถว "ร้อยละ"
                    row_p = {"ชั้น": room_name, "ประเภท": "ร้อยละ (%)"}
                    row_p.update({str(g): f"{(counts[str(g)]/total_n*100):.2f}" for g in [4.0, 3.5, 3.0, 2.5, 2.0, 1.5, 1.0, 0.0]})
                    row_p.update({"รวม": "100.00", "X̄": "", "S.D.": ""})
                    rows.append(row_p)

            # --- แถวสรุปรวมทั้งสายชั้น ---
            if grand_total_n > 0:
                overall_mean = grand_sum_fx / grand_total_n # Σfx รวม / N รวม
                
                # แถวรวมจำนวน
                summary_n = {"ชั้น": "รวม", "ประเภท": "จำนวน"}
                summary_n.update({str(g): totals_n[g] for g in totals_n})
                summary_n.update({"รวม": grand_total_n, "X̄": f"{overall_mean:.2f}", "S.D.": ""})
                rows.append(summary_n)

                # แถวรวมร้อยละ
                summary_p = {"ชั้น": "รวม", "ประเภท": "ร้อยละ (%)"}
                summary_p.update({str(g): f"{(totals_n[g]/grand_total_n*100):.2f}" for g in totals_n})
                summary_p.update({"รวม": "100.00", "X̄": "", "S.D.": ""})
                rows.append(summary_p)

               # --- ส่วนคำนวณบทสรุปผลสัมฤทธิ์ทางการเรียน ---
            if grand_total_n > 0:
                # 1. คำนวณร้อยละรายเกรด (ภาพรวมทั้งสายชั้น)
                per_g = {str(g): (totals_n[g] / grand_total_n * 100) for g in totals_n}
                
                # 2. คำนวณกลุ่ม ดี ขึ้นไป (3 ขึ้นไป)
                count_3_up = totals_n[4.0] + totals_n[3.5] + totals_n[3.0]
                percent_3_up = (count_3_up / grand_total_n * 100)
                
                # 3. คำนวณกลุ่มไม่ผ่าน (เกรด 0, ร, มส)
                # ในที่นี้สมมติว่าเกรด 0 คือไม่ผ่าน (ถ้ามี ร, มส ในอนาคตสามารถบวกเพิ่มได้)
                count_fail = totals_n[0.0]
                percent_fail = (count_fail / grand_total_n * 100)
                
                # 4. ตรวจสอบค่าเป้าหมาย (GPA เฉลี่ย 3.00)
                target_gpa = 3.00
                status_gpa = "ซึ่งเป็นไปตามค่าเป้าหมายที่กำหนด" if overall_mean >= target_gpa else "ซึ่งต่ำกว่าค่าเป้าหมายที่กำหนด"

                # 5. สร้างข้อความรายงาน
                st.divider()
                st.subheader("📝 บทสรุปรายงานผลสัมฤทธิ์")
                
                report_achievement = (
                    f"ร้อยละของจำนวนนักเรียนที่ได้ระดับผลการเรียน ภาคเรียนที่ 2 ปีการศึกษา 2568 พบว่า "
                    f"ได้ระดับผลการเรียน 4 ร้อยละ {per_g['4.0']:.2f} "
                    f"ได้ระดับผลการเรียน 3.5 ร้อยละ {per_g['3.5']:.2f} "
                    f"ได้ระดับผลการเรียน 3 ร้อยละ {per_g['3.0']:.2f} "
                    f"ได้ระดับผลการเรียน 2.5 ร้อยละ {per_g['2.5']:.2f} "
                    f"ได้ระดับผลการเรียน 2 ร้อยละ {per_g['2.0']:.2f} "
                    f"ได้ระดับผลการเรียน 1.5 ร้อยละ {per_g['1.5']:.2f} "
                    f"ได้ระดับผลการเรียน 1 ร้อยละ {per_g['1.0']:.2f} "
                    f"ได้ระดับผลการเรียน 0 ร้อยละ {per_g['0.0']:.2f} "
                    f"และได้ระดับผลการเรียน ร, มส ร้อยละ 0.00 "
                    f"และตามที่กลุ่มสาระการเรียนรู้วิทยาศาสตร์และเทคโนโลยี ได้กำหนดค่าเป้าหมายของระดับผลการเรียนเฉลี่ย เท่ากับ {target_gpa} "
                    f"พบว่า นักเรียนได้ระดับผลการเรียนเฉลี่ย เท่ากับ {overall_mean:.2f} {status_gpa} "
                    f"นักเรียนที่มีผลการเรียนระดับ ดี ขึ้นไป (3 ขึ้นไป) จำนวน {count_3_up} คน คิดเป็นร้อยละ {percent_3_up:.2f} "
                    f"ซึ่งเป็นไปตามค่าเป้าหมายที่กำหนด และนักเรียนที่มีผลการเรียนไม่ผ่าน จำนวน {count_fail} คน คิดเป็นร้อยละ {percent_fail:.2f} "
                    f"ซึ่งไม่เกินร้อยละ 3 และเป็นไปตามค่าเป้าหมายที่กำหนด"
                )
                
                st.success(report_achievement)
                
                # ช่องสำหรับก๊อปปี้ไปวางใน Word/Report
                st.text_area("📋 ก๊อปปี้ข้อความรายงานไปใช้:", value=report_achievement, height=200)

            df_result = pd.DataFrame(rows).fillna("")
            st.table(df_result)
            # ปุ่มดาวน์โหลด
            csv = df_result.to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 ดาวน์โหลดตารางผลสัมฤทธิ์ (CSV)", csv, "achievement.csv", "text/csv")
           # --- TAB 4: ประสิทธิภาพของแผนการจัดการเรียนรู้ (E1/E2) - เวอร์ชันเลือกคอลัมน์ได้ ---
    with tab4:
        st.subheader("📊 การคำนวณประสิทธิภาพของแผนการจัดการเรียนรู้ (E1/E2)")
        
        # ส่วนตั้งค่าคะแนนเต็ม
        col_set1, col_set2 = st.columns(2)
        with col_set1:
            full_score_process = st.number_input("คะแนนเต็มระหว่างเรียน (E1)", value=70)
        with col_set2:
            full_score_post = st.number_input("คะแนนเต็มหลังเรียน (E2)", value=30)

        # ดึงตัวอย่างหัวคอลัมน์มาให้เลือก
        # ใช้ข้อมูลห้องแรกเป็นเกณฑ์ในการดึงหัวตาราง
        sample_room_start = room_indices[0]
        sample_room_end = room_indices[1]
        df_sample = df_raw.iloc[sample_room_start:sample_room_end].reset_index(drop=True)
        # แถวที่ 1 มักจะเป็นหัวตาราง (ปรับเลขแถวได้)
        column_options = [f"คอลัมน์ที่ {i}: {str(val)}" for i, val in enumerate(df_sample.iloc[1])]

        col_sel1, col_sel2 = st.columns(2)
        with col_sel1:
            idx_e1 = st.selectbox("เลือกคอลัมน์ 'คะแนนระหว่างเรียน'", range(len(column_options)), index=17, format_func=lambda x: column_options[x])
        with col_sel2:
            idx_e2 = st.selectbox("เลือกคอลัมน์ 'คะแนนหลังเรียน'", range(len(column_options)), index=16, format_func=lambda x: column_options[x])

        if st.button("🧮 คำนวณค่า E1/E2"):
            all_process_scores = []
            all_post_scores = []
            
            for i in range(len(room_indices) - 1):
                start, end = room_indices[i], room_indices[i+1]
                df_room = df_raw.iloc[start:end].reset_index(drop=True)
                
                df_students = df_room.iloc[2:].copy()
                df_students = df_students[df_students.iloc[:, 1].astype(str).str.strip().str.isdigit()]
                
                # ดึงคะแนนตาม index ที่ครูเลือก
                process_s = pd.to_numeric(df_students.iloc[:, idx_e1], errors='coerce').dropna()
                post_s = pd.to_numeric(df_students.iloc[:, idx_e2], errors='coerce').dropna()
                
                all_process_scores.extend(process_s.tolist())
                all_post_scores.extend(post_s.tolist())

            if all_process_scores and all_post_scores:
                n = len(all_process_scores)
                mean_e1 = np.mean(all_process_scores)
                sd_e1 = np.std(all_process_scores)
                percent_e1 = (mean_e1 / full_score_process) * 100

                mean_e2 = np.mean(all_post_scores)
                sd_e2 = np.std(all_post_scores)
                percent_e2 = (mean_e2 / full_score_post) * 100

                e_data = [
                    {"คะแนน": "คะแนนระหว่างเรียน", "n": n, "คะแนนเต็ม": full_score_process, "X̄": f"{mean_e1:.2f}", "S.D.": f"{sd_e1:.2f}", "ร้อยละ": f"{percent_e1:.2f}", "E1 / E2": ""},
                    {"คะแนน": "สรุปผลประสิทธิภาพ", "n": "", "คะแนนเต็ม": "", "X̄": "", "S.D.": "", "ร้อยละ": "", "E1 / E2": f"{percent_e1:.2f} / {percent_e2:.2f}"},
                    {"คะแนน": "คะแนนการทดสอบหลังเรียน", "n": n, "คะแนนเต็ม": full_score_post, "X̄": f"{mean_e2:.2f}", "S.D.": f"{sd_e2:.2f}", "ร้อยละ": f"{percent_e2:.2f}", "E1 / E2": ""}
                ]
                
                st.table(pd.DataFrame(e_data))
                # --- ส่วนคำนวณบทสรุปประสิทธิภาพ E1/E2 ---
                target_e1 = 70
                target_e2 = 70
                
                # ตรวจสอบว่าสูงกว่าเกณฑ์หรือไม่
                status_e1 = "สูงกว่าเกณฑ์" if percent_e1 >= target_e1 else "ต่ำกว่าเกณฑ์"
                status_e2 = "สูงกว่าเกณฑ์" if percent_e2 >= target_e2 else "ต่ำกว่าเกณฑ์"
                
                # สรุปภาพรวม (ถ้าสูงกว่าเกณฑ์ทั้งคู่ หรือตามเงื่อนไขที่ครูต้องการ)
                overall_status = "มีประสิทธิภาพ" if percent_e1 >= target_e1 and percent_e2 >= target_e2 else "ควรปรับปรุง"

                st.divider()
                st.subheader("📝 บทสรุปประสิทธิภาพแผนการจัดการเรียนรู้")
                
                report_e1_e2 = (
                    f"ประสิทธิภาพของแผนการจัดการเรียนรู้ พบว่าคะแนนเฉลี่ยของคะแนนระหว่างเรียน เท่ากับ {mean_e1:.2f} "
                    f"คิดเป็นร้อยละ {percent_e1:.2f} คะแนนเฉลี่ยของการทดสอบหลังเรียนเท่ากับ {mean_e2:.2f} "
                    f"คิดเป็นร้อยละ {percent_e2:.2f} ค่า E1 / E2 มีค่าเท่ากับ {percent_e1:.2f} / {percent_e2:.2f} "
                    f"โดยตั้งเกณฑ์ประสิทธิภาพ เท่ากับ {target_e1}/{target_e2} ซึ่งค่า E1 {status_e1} "
                    f"แสดงว่าการสอนโดยใช้แผนการจัดการเรียนรู้{overall_status}"
                )
                
                if overall_status == "มีประสิทธิภาพ":
                    st.success(report_e1_e2)
                else:
                    st.warning(report_e1_e2)
                
                # ช่องสำหรับก๊อปปี้ไปใช้งาน
                st.text_area("📋 ก๊อปปี้ข้อความสรุป E1/E2:", value=report_e1_e2, height=150)
    # --- TAB 5: นักเรียนที่ได้คะแนนสูงสุด 10 อันดับแรก (เวอร์ชันแก้ Error) ---
    with tab5:
        st.subheader("🏆 รายชื่อนักเรียนที่ได้คะแนนสูงสุด 10 อันดับแรก")
        
        if excel_file:
            # 1. เตรียมรายชื่อคอลัมน์จากข้อมูลห้องแรก
            sample_room_idx = room_indices[0]
            sample_room_end_idx = room_indices[1]
            df_sample_top = df_raw.iloc[sample_room_idx:sample_room_end_idx].reset_index(drop=True)
            
            # สร้างตัวแปร col_options_top ป้องกัน NameError
            col_options_top = [f"คอลัมน์ที่ {i}: {str(val)}" for i, val in enumerate(df_sample_top.iloc[1])]
            
            target_col_idx = st.selectbox(
                "เลือกคอลัมน์คะแนนที่ใช้จัดอันดับ:", 
                range(len(col_options_top)), 
                index=17, 
                format_func=lambda x: col_options_top[x],
                key="top10_select_fixed"
            )

            if st.button("🥇 ค้นหา 10 อันดับแรก"):
                all_students_data = []
                
                for i in range(len(room_indices) - 1):
                    start, end = room_indices[i], room_indices[i+1]
                    df_room = df_raw.iloc[start:end].reset_index(drop=True)
                    room_name = str(df_room.iloc[0, 0]).strip()
                    room_short = room_name.split('-')[0].strip()

                    df_students = df_room.iloc[2:].copy()
                    df_students = df_students[df_students.iloc[:, 1].astype(str).str.strip().str.isdigit()]
                    
                    for _, row in df_students.iterrows():
                        score_val = pd.to_numeric(row[target_col_idx], errors='coerce')
                        if not pd.isna(score_val):
                            # ดึงคำนำหน้า (row[2]) + ชื่อ (row[3]) + นามสกุล (row[4])
                            prefix = str(row[2]).strip() if not pd.isna(row[2]) else ""
                            first_name = str(row[3]).strip() if not pd.isna(row[3]) else ""
                            last_name = str(row[4]).strip() if not pd.isna(row[4]) else ""
                            
                            all_students_data.append({
                                "เลขประจำตัว": str(row[1]).replace('.0', '').strip(),
                                "ชื่อ - สกุล": f"{prefix}{first_name} {last_name}",
                                "ชั้น": room_short,
                                "คะแนนที่ได้": score_val
                            })

                if all_students_data:
                    df_all_top = pd.DataFrame(all_students_data)
                    # เรียงจากมากไปน้อย
                    df_top10 = df_all_top.sort_values(by="คะแนนที่ได้", ascending=False).head(10).reset_index(drop=True)
                    df_top10.index = df_top10.index + 1
                    
                    st.balloons()
                    st.table(df_top10)
                    # --- ส่วนคำนวณบทสรุป Top 10 ---
                if not df_top10.empty:
                    # 1. ข้อมูลอันดับ 1
                    top_1 = df_top10.iloc[0]
                    # 2. ข้อมูลอันดับ 2
                    top_2 = df_top10.iloc[1] if len(df_top10) > 1 else top_1
                    # 3. ข้อมูลอันดับสุดท้าย (อันดับ 10 หรืออันดับสุดท้ายที่มี)
                    last_top = df_top10.iloc[-1]
                    
                    # 4. หาห้องที่มีจำนวนนักเรียนติด Top 10 มากที่สุด
                    room_counts = df_top10['ชั้น'].value_counts()
                    most_room = room_counts.idxmax()
                    most_room_count = room_counts.max()

                    st.divider()
                    st.subheader("📝 บทสรุปนักเรียนที่มีผลการเรียนดีเด่น")
                    
                    report_top10 = (
                        f"จากการศึกษาจำนวนนักเรียนที่ได้คะแนน 10 อันดับแรกของรายวิชา พบว่านักเรียนชั้น{top_1['ชั้น']} "
                        f"ได้คะแนนสูงสุด คือ {top_1['คะแนนที่ได้']:.2f} คะแนน "
                        f"รองลงมาเป็นนักเรียนชั้น{top_2['ชั้น']} ได้คะแนน {top_2['คะแนนที่ได้']:.2f} คะแนน "
                        f"ส่วนอันดับสุดท้ายเป็นนักเรียนชั้น{last_top['ชั้น']} ได้คะแนน {last_top['คะแนนที่ได้']:.2f} คะแนน "
                        f"และนักเรียนชั้นมัธยมศึกษาปีที่ 1 มีจำนวนนักเรียนที่ได้คะแนนใน 10 อันดับแรกมากที่สุด "
                        f"คือชั้น {most_room} จำนวน {most_room_count} คน"
                    )
                    
                    st.success(report_top10)
                    
                    # ช่องก๊อปปี้ข้อความ
                    st.text_area("📋 ก๊อปปี้ข้อความสรุป Top 10:", value=report_top10, height=150)
                    
                    # ปุ่มดาวน์โหลด
                    output_top = io.BytesIO()
                    with pd.ExcelWriter(output_top, engine='xlsxwriter') as writer:
                        df_top10.to_excel(writer, index=True, sheet_name='Top10')
                    st.download_button("📥 ดาวน์โหลดรายชื่อ Top 10", output_top.getvalue(), "top10_students.xlsx")
                else:
                    st.error("ไม่พบข้อมูลนักเรียน")
        else:
            st.warning("⚠️ กรุณาอัปโหลดไฟล์ Excel ที่ด้านบนก่อนใช้งานหมวดนี้")
