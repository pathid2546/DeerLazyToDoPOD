import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime
import streamlit.components.v1 as components

st.set_page_config(page_title="POD BNN - ตัวแม่เล็บเจล", layout="wide")

# --- ฟังก์ชันฉีดจริต "ตัวแม่" พุ่งทะลุจอ ---
def inject_fabulous_nails():
    # ใช้ HTML + JS ฉีดเข้า Component บังคับรันเล็บเจลพุ่ง
    components.html(
        """
        <script src="https://cdn.jsdelivr.net/npm/canvas-confetti@1.5.1/dist/confetti.browser.min.js"></script>
        <script>
            function fireFabulous() {
                var duration = 5 * 1000;
                var animationEnd = Date.now() + duration;
                var defaults = { startVelocity: 45, spread: 360, ticks: 100, zIndex: 999999 };

                // เลือกเฉพาะ Emoji ที่สื่อถึงความตัวแม่และเล็บเจล
                var nail1 = confetti.shapeFromText({ text: '💅🏻', scalar: 7 });
                var nail2 = confetti.shapeFromText({ text: '💅🏼', scalar: 7 });
                var star = confetti.shapeFromText({ text: '✨', scalar: 5 });
                var crown = confetti.shapeFromText({ text: '👑', scalar: 6 });

                var interval = setInterval(function() {
                    var timeLeft = animationEnd - Date.now();
                    if (timeLeft <= 0) return clearInterval(interval);

                    var particleCount = 50 * (timeLeft / duration);
                    
                    // พ่นจากฝั่งซ้ายแบบสับๆ
                    confetti(Object.assign({}, defaults, { 
                        particleCount, 
                        origin: { x: 0, y: 0.7 },
                        angle: 60,
                        shapes: [nail1, nail2, star, crown],
                        colors: ['#ff4bad', '#ffecf5']
                    }));
                    // พ่นจากฝั่งขวาแบบตัวแม่
                    confetti(Object.assign({}, defaults, { 
                        particleCount, 
                        origin: { x: 1, y: 0.7 },
                        angle: 120,
                        shapes: [nail1, nail2, star, crown],
                        colors: ['#ff4bad', '#ffecf5']
                    }));
                }, 250);
            }
            
            // สั่งรันทันที
            fireFabulous();
        </script>
        <div style="
            position: fixed; 
            top: 50%; left: 50%; 
            transform: translate(-50%, -50%);
            z-index: 1000000;
            text-align: center;
            pointer-events: none;
            width: 100%;
        ">
            <h1 style="
                font-family: 'Tahoma', sans-serif;
                font-size: 130px;
                color: #ff4bad;
                text-shadow: 10px 10px 0px #fff, 15px 15px 30px rgba(255, 75, 173, 0.7);
                animation: zoomPulse 0.8s ease-in-out infinite alternate;
                margin: 0;
            ">เริ่ดเลยหล่ะ</h1>
            <p style="font-size: 40px; color: #ff4bad; font-weight: bold; background: white; display: inline-block; padding: 5px 20px; border-radius: 50px;">💅🏻 สวย สับ ปัง! ✨</p>
        </div>
        <style>
            @keyframes zoomPulse {
                from { transform: scale(0.9); opacity: 0.8; }
                to { transform: scale(1.1); opacity: 1; }
            }
        </style>
        """,
        height=600, # ขยายพื้นที่ให้ตัวหนังสือและเล็บเจลพุ่งได้เต็มที่
    )

# --- ส่วนประมวลผล Excel (คงเดิมเพื่อความแม่นยำ) ---
def process_excel_to_buffer(uploaded_file):
    raw_df = pd.read_excel(uploaded_file, header=None)
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in row.values)
    df = pd.read_excel(uploaded_file, header=header_row_index).iloc[1:].reset_index(drop=True)
    df.columns = df.columns.str.strip()
    
    # ดึงรหัสสาขาจากหัวตาราง
    header_raw = raw_df.iloc[header_row_index].astype(str).str.strip().tolist()
    code_raw = raw_df.iloc[header_row_index + 1].tolist()
    branch_map = {n: str(c) if pd.notna(c) else "" for n, c in zip(header_raw, code_raw) if n and n != 'nan'}

    df_melted = df.melt(id_vars=['Item No.', 'Description', 'UNIT'], var_name='Branch', value_name='Qty')
    df_melted['Qty'] = pd.to_numeric(df_melted['Qty'], errors='coerce')
    final_list = df_melted.dropna(subset=['Item No.', 'Qty']).query('Qty > 0')
    
    current_date = datetime.now().strftime('%d/%m/%Y')
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # แยกสาขา
        for b_name, b_data in final_list.groupby('Branch', sort=False):
            sheet_name = str(b_name)[:30].replace('/', '-')
            items_df = pd.DataFrame({'No': range(1, len(b_data)+1), 'Code': b_data['Item No.'], 'Name': b_data['Description'], 'Unit': b_data['UNIT'], 'ORD': b_data['Qty'], 'MBL': '', 'BNN': ''})
            items_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=10)
            apply_styles(writer.sheets[sheet_name], b_name, branch_map.get(str(b_name).strip(), ""), current_date)
        
        # สรุปจบ
        summary = final_list.groupby(['Item No.', 'Description', 'UNIT'], sort=False)['Qty'].sum().reset_index()
        summary.insert(0, 'No', range(1, len(summary)+1))
        summary.to_excel(writer, sheet_name="Summary_All", index=False, header=False, startrow=10)
        ws_sum = writer.sheets["Summary_All"]
        last_row = 10 + len(summary) + 1
        ws_sum.cell(row=last_row, column=3, value="Grand Total").font = Font(name='Cordia New', bold=True, size=14)
        ws_sum.cell(row=last_row, column=5, value=summary['Qty'].sum()).font = Font(name='Cordia New', bold=True, size=14)
        apply_styles(ws_sum, "สรุปยอดรวม", "ALL", current_date, True)
    
    return output.getvalue()

def apply_styles(ws, b_name, b_code, date, is_sum=False):
    f_data = Font(name='Cordia New', size=14)
    f_white = Font(name='Cordia New', bold=True, color="FFFFFF", size=14)
    fill_green = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    ws.merge_cells('A1:G1')
    ws['A1'] = "ใบสรุปยอดรวม" if is_sum else "ใบส่งสินค้าชั่วคราว"
    ws['A1'].font = Font(name='Cordia New', bold=True, size=20)
    ws['A1'].alignment = Alignment(horizontal='center')
    ws['A7'], ws['C7'] = f"Code: {b_code}", f"Name: {b_name}"
    
    headers = ['No', 'Code', 'Name', 'Unit', 'TOTAL' if is_sum else 'ORDER', 'MBL', 'BNN']
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=10, column=i, value=h)
        c.font, c.fill, c.border, c.alignment = f_white, fill_green, border, Alignment(horizontal='center')
    
    for row in ws.iter_rows(min_row=11, max_row=ws.max_row, min_col=1, max_col=7):
        for cell in row: cell.border, cell.font = border, f_data
        
    for col, w in zip('ABCDEFG', [6, 16, 35, 10, 12, 10, 10]): ws.column_dimensions[col].width = w

# --- UI ตัวแม่ ---
st.markdown("<h1 style='text-align: center; color: #ff4bad; font-size: 60px;'>✨💅🏻 คลังตัวแม่เล็บเจล 💅🏻✨</h1>", unsafe_allow_html=True)

file = st.file_uploader("ส่งไฟล์มาให้แม่จัดการค่ะ", type="xlsx")

if file:
    if st.button("💖 เนรมิตความเริ่ด เดี๋ยวนี้! 💖", use_container_width=True):
        with st.spinner('เล็บเจลกำลังแห้ง... รอแป๊บนึงนะแม่...'):
            st.session_state.out_bytes = process_excel_to_buffer(file)
            st.session_state.fab = True

    if st.session_state.get('fab'):
        # ฉีดจริตเล็บเจลพุ่ง
        inject_fabulous_nails()
        
        st.markdown("<div style='text-align: center; border: 2px solid #ff4bad; padding: 15px; border-radius: 15px; background: #fff0f5;'><h3>💅🏻 เล็บเจลพุ่งมั้ยคะคุณนาย? ไฟล์พร้อมแล้วค่ะ!</h3></div>", unsafe_allow_html=True)
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ความเริ่ดระดับตัวแม่ 📥",
            data=st.session_state.out_bytes,
            file_name="POD_Fabulous_Nails.xlsx",
            use_container_width=True
        )