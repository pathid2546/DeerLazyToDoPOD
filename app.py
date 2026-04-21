import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime
import streamlit.components.v1 as components

# ตั้งค่าหน้าจอแบบตัวแม่ Wide Layout
st.set_page_config(page_title="POD BNN - Fabulous & Radical All", layout="wide")

# --- ท่าไม้ตาย: ฉีดความแรดระดับตำนาน (Top Level Injection) ---
def inject_radical_fabulousness():
    # ใช้ HTML + JS แบบจัดเต็ม บังคับพุ่งทะลุจอ
    components.html(
        """
        <script src="https://cdn.jsdelivr.net/npm/canvas-confetti@1.5.1/dist/confetti.browser.min.js"></script>
        <div id="radical-overlay" style="
            position: fixed;
            top: 0; left: 0;
            width: 100vw; height: 100vh;
            background-color: rgba(255, 105, 180, 0.05); /* เคลือบชมพูอ่อน */
            z-index: 999999;
            display: flex;
            justify-content: center;
            align-items: center;
            pointer-events: none;
        ">
            <div style="text-align: center;">
                <h1 id="radical-text" style="
                    font-family: 'Tahoma', sans-serif;
                    font-size: 180px; /* ใหญ่ยักษ์ตะโกน */
                    color: #ff4bad; /* ชมพูสะท้อนแสง */
                    text-shadow: 
                        0 0 10px #fff,
                        0 0 20px #ff4bad,
                        0 0 30px #ff4bad,
                        0 0 40px #ff4bad,
                        5px 5px 0px #fff;
                    transform: scale(0);
                    transition: transform 0.6s cubic-bezier(0.175, 0.885, 0.32, 1.275);
                    margin: 0;
                ">แรดมากกกก</h1>
                <div style="
                    font-size: 50px; 
                    color: #ff4bad; 
                    font-weight: bold; 
                    background: white; 
                    display: inline-block; 
                    padding: 10px 40px; 
                    border-radius: 50px; 
                    border: 4px solid #ff4bad; 
                    margin-top: 20px;
                    box-shadow: 5px 5px 15px rgba(255, 75, 173, 0.4);
                ">
                    💅🏻✨ สวย สับ แรด ระดับตัวแม่! 👑💥
                </div>
            </div>
        </div>

        <script>
            // แสดงข้อความแรดมากกกก
            const text = document.getElementById('radical-text');
            setTimeout(() => { text.style.transform = 'scale(1.1)'; }, 100);

            // พ่นอีโมจิแบบบ้าคลั่ง
            var duration = 5 * 1000;
            var animationEnd = Date.now() + duration;
            var defaults = { startVelocity: 45, spread: 360, ticks: 100, zIndex: 1000000 };

            function randomInRange(min, max) { return Math.random() * (max - min) + min; }

            // เลือกอีโมจิแบบแรดๆ
            var nail = confetti.shapeFromText({ text: '💅🏻', scalar: 6 });
            var sparkle = confetti.shapeFromText({ text: '✨', scalar: 5 });
            var crown = confetti.shapeFromText({ text: '👑', scalar: 6 });
            var burst = confetti.shapeFromText({ text: '💥', scalar: 7 });
            var heart = confetti.shapeFromText({ text: '💖', scalar: 6 });

            var interval = setInterval(function() {
                var timeLeft = animationEnd - Date.now();
                if (timeLeft <= 0) { return clearInterval(interval); }

                var particleCount = 60 * (timeLeft / duration);
                
                // พ่นจากซ้ายแบบสับๆ
                confetti(Object.assign({}, defaults, { 
                    particleCount, 
                    origin: { x: randomInRange(0, 0.2), y: Math.random() - 0.2 },
                    shapes: [nail, sparkle, crown, burst, heart],
                    colors: ['#ff4bad', '#ffecf5', '#ff1493', '#ffd700']
                }));
                // พ่นจากขวาแบบแรดๆ
                confetti(Object.assign({}, defaults, { 
                    particleCount, 
                    origin: { x: randomInRange(0.8, 1), y: Math.random() - 0.2 },
                    shapes: [nail, sparkle, crown, burst, heart],
                    colors: ['#ff4bad', '#ffecf5', '#ff1493', '#ffd700']
                }));
            }, 200);
        </script>
        """,
        height=0, # กำหนด height=0 เพื่อไม่ให้กินพื้นที่ UI
    )

# --- ส่วนประมวลผล (เหมือนเดิมเป๊ะ) ---
def process_excel_to_buffer(uploaded_file):
    raw_df = pd.read_excel(uploaded_file, header=None)
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in row.values)
    header_row_raw = raw_df.iloc[header_row_index].astype(str).str.strip().tolist()
    code_row_raw = raw_df.iloc[header_row_index + 1].tolist()
    branch_to_code = {name: str(code) if pd.notna(code) else "" 
                      for name, code in zip(header_row_raw, code_row_raw) if name and name != 'nan'}
    df = pd.read_excel(uploaded_file, header=header_row_index).iloc[1:].reset_index(drop=True)
    df.columns = df.columns.str.strip()
    df_melted = df.melt(id_vars=['Item No.', 'Description', 'UNIT'], var_name='Branch', value_name='Qty')
    df_melted['Qty'] = pd.to_numeric(df_melted['Qty'], errors='coerce')
    final_list = df_melted.dropna(subset=['Item No.', 'Qty']).query('Qty > 0')
    current_date = datetime.now().strftime('%d/%m/%Y')
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for branch_name, branch_data in final_list.groupby('Branch', sort=False):
            store_code = branch_to_code.get(str(branch_name).strip(), "")
            sheet_name = str(branch_name)[:30].replace('/', '-').replace(':', '')
            items_df = pd.DataFrame({
                'No': range(1, len(branch_data) + 1),
                'Code': branch_data['Item No.'],
                'Name': branch_data['Description'],
                'Unit': branch_data['UNIT'],
                'ORD': branch_data['Qty'],
                'MBL': "", 'BNN': ""
            })
            items_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=10)
            apply_styles_to_sheet(writer.sheets[sheet_name], branch_name, store_code, current_date)
        summary_all = final_list.groupby(['Item No.', 'Description', 'UNIT'], sort=False)['Qty'].sum().reset_index()
        summary_all.insert(0, 'No', range(1, len(summary_all) + 1))
        summary_all.to_excel(writer, sheet_name="Summary_All", index=False, header=False, startrow=10)
        ws_sum = writer.sheets["Summary_All"]
        last_row = 10 + len(summary_all) + 1
        ws_sum.cell(row=last_row, column=3, value="Grand Total (ยอดรวมทั้งหมด)").font = Font(name='Cordia New', bold=True, size=14)
        qty_cell = ws_sum.cell(row=last_row, column=5, value=summary_all['Qty'].sum())
        qty_cell.font = Font(name='Cordia New', bold=True, size=14)
        qty_cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        qty_cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='double'))
        apply_styles_to_sheet(ws_sum, "สรุปยอดรวมทุกรายการ", "ALL", current_date, is_summary=True)
    return output.getvalue()

def apply_styles_to_sheet(ws, branch_name, store_code, current_date, is_summary=False):
    font_name = 'Cordia New'
    f_title, f_header, f_data = Font(name=font_name, bold=True, size=20), Font(name=font_name, bold=True, size=14), Font(name=font_name, size=14)
    f_white = Font(name=font_name, bold=True, color="FFFFFF", size=14)
    fill_green, fill_light = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid"), PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    ws.merge_cells('A1:G1')
    ws['A1'] = "ใบสรุปยอดรวม" if is_summary else "ใบส่งสินค้าชั่วคราว"
    ws['A1'].font, ws['A1'].alignment = f_title, Alignment(horizontal='center')
    ws['A2'], ws['A3'], ws['A4'] = "บริษัท โมบาย โลจิสติกส์ จำกัด", "278 หมู่ที่ 9 จ.สมุทรปราการ", "โทร. 02-337-1200"
    ws['A2'].font, ws['A3'].font, ws['A4'].font = f_header, f_data, f_data
    ws['G2'], ws['G4'] = f"Date: {current_date}", f"Delivery Date: {current_date}"
    ws['G2'].alignment = ws['G4'].alignment = Alignment(horizontal='right')
    ws['G2'].font = ws['G4'].font = f_header
    ws['A7'], ws['C7'] = f"Code: {store_code}", f"Name: {branch_name}"
    ws['A7'].font = ws['C7'].font = f_header
    headers = ['No', 'Code', 'Name', 'Unit', 'TOTAL' if is_summary else 'ORDER', 'MBL', 'BNN']
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=10, column=i, value=h); c.font, c.fill, c.border, c.alignment = f_white, fill_green, border, Alignment(horizontal='center')
    for r_idx, row in enumerate(ws.iter_rows(min_row=11, max_row=ws.max_row, min_col=1, max_col=7), 1):
        if ws.cell(row=row[0].row, column=3).value == "Grand Total (ยอดรวมทั้งหมด)": continue
        for cell in row:
            cell.border, cell.font = border, f_data
            if r_idx % 2 == 0: cell.fill = fill_light
            if cell.column in [1, 4, 5, 6, 7]: cell.alignment = Alignment(horizontal='center')
    ws.page_setup.paperSize = 9
    widths = {'A': 6, 'B': 16, 'C': 35, 'D': 10, 'E': 12, 'F': 10, 'G': 10}
    for col, w in widths.items(): ws.column_dimensions[col].width = w

# --- UI Interface จริตแรดมากกกก ---
# ใช้ Markdown ฉีด CSS เพื่อซ่อนลูกโป่ง Streamlit และปรับสีปุ่ม
st.markdown("""
    <style>
    /* ซ่อนลูกโป่ง Streamlit default */
    .stBalloons { display: none !important; }
    /* ปรับสีปุ่ม Process ให้เป็นชมพูแรด */
    div.stButton > button:first-child {
        background-color: #ff4bad;
        color: white;
        border: 2px solid #ff4bad;
        border-radius: 50px;
        font-weight: bold;
        transition: all 0.3s;
    }
    div.stButton > button:first-child:hover {
        background-color: white;
        color: #ff4bad;
        border: 2px solid #ff4bad;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: center; color: #ff4bad; font-size: 60px;'>✨💅🏻 คลังแรดระดับตัวแม่ 💅🏻✨</h1>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload ไฟล์ Excel ที่แม่ต้องจัดการ (แบบแรดๆ)", type="xlsx")

if uploaded_file:
    if st.button("🚀 เนรมิตไฟล์เดี๋ยวนี้! (แรดพุ่งนะแม่) 🚀", use_container_width=True):
        with st.spinner('กำลังเนรมิตความสวยแบบแรดๆ... แป๊บนึงนะแม่...'):
            st.session_state.out_bytes = process_excel_to_buffer(uploaded_file)
            st.session_state.fab_radical = True

    if st.session_state.get('fab_radical'):
        # บังคับพ่นความแรดระดับตำนาน ทะลุจอ!
        inject_radical_fabulousness()
        
        st.success("เรียบร้อยค่ะแม่! แรดสมใจมั้ยคะ? 💅🏻👑💥")
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ความแรด 📥",
            data=st.session_state.out_bytes,
            file_name=f"POD_Radical_{datetime.now().strftime('%H%M')}.xlsx",
            use_container_width=True
        )