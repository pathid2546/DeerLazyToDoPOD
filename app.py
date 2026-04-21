import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime
import streamlit.components.v1 as components

st.set_page_config(page_title="POD BNN - คุณนายเล็บเจล", layout="wide")

# --- ท่าไม้ตาย: ฉีดความเริ่ดเข้ากระแสเลือด (Inject JS to Top Level) ---
def inject_fabulous_nails():
    # ใช้ท่า components.html แบบกำหนดให้มันไปงอกที่หน้าจอหลัก
    components.html(
        """
        <script src="https://cdn.jsdelivr.net/npm/canvas-confetti@1.5.1/dist/confetti.browser.min.js"></script>
        <script>
            // ฟังก์ชันพ่นเล็บเจลแบบสับ
            function fireNails() {
                var duration = 5 * 1000;
                var animationEnd = Date.now() + duration;
                var defaults = { startVelocity: 45, spread: 360, ticks: 100, zIndex: 999999 };

                var nail1 = confetti.shapeFromText({ text: '💅🏻', scalar: 6 });
                var nail2 = confetti.shapeFromText({ text: '💅🏼', scalar: 6 });
                var star = confetti.shapeFromText({ text: '✨', scalar: 5 });

                var interval = setInterval(function() {
                    var timeLeft = animationEnd - Date.now();
                    if (timeLeft <= 0) return clearInterval(interval);

                    var particleCount = 40 * (timeLeft / duration);
                    // พ่นจากซ้าย
                    confetti(Object.assign({}, defaults, { 
                        particleCount, 
                        origin: { x: 0, y: 0.7 },
                        angle: 60,
                        shapes: [nail1, nail2, star]
                    }));
                    // พ่นจากขวา
                    confetti(Object.assign({}, defaults, { 
                        particleCount, 
                        origin: { x: 1, y: 0.7 },
                        angle: 120,
                        shapes: [nail1, nail2, star]
                    }));
                }, 250);
            }
            
            // รันทันทีที่ Component โหลด
            fireNails();
        </script>
        <div style="
            position: fixed; 
            top: 50%; left: 50%; 
            transform: translate(-50%, -50%);
            z-index: 1000000;
            text-align: center;
            pointer-events: none;
        ">
            <h1 style="
                font-family: 'Tahoma', sans-serif;
                font-size: 150px;
                color: #ff4bad;
                text-shadow: 8px 8px 0px #fff, 12px 12px 20px rgba(255, 75, 173, 0.6);
                animation: zoomInOut 1s ease-in-out infinite alternate;
            ">เริ่ดเลยหล่ะ</h1>
        </div>
        <style>
            @keyframes zoomInOut {
                from { transform: scale(0.8); }
                to { transform: scale(1.1); }
            }
        </style>
        """,
        height=500, # จองที่ให้ตัวหนังสือเด้ง
    )

# --- ส่วนประมวลผล (เหมือนเดิมเป๊ะเพื่อความถูกต้องของงาน) ---
def process_excel_to_buffer(uploaded_file):
    raw_df = pd.read_excel(uploaded_file, header=None)
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in row.values)
    header_row_raw = raw_df.iloc[header_row_index].astype(str).str.strip().tolist()
    code_row_raw = raw_df.iloc[header_row_index + 1].tolist()
    branch_to_code = {name: str(code) if pd.notna(code) else "" for name, code in zip(header_row_raw, code_row_raw) if name and name != 'nan'}
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
            items_df = pd.DataFrame({'No': range(1, len(branch_data) + 1), 'Code': branch_data['Item No.'], 'Name': branch_data['Description'], 'Unit': branch_data['UNIT'], 'ORD': branch_data['Qty'], 'MBL': "", 'BNN': ""})
            items_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=10)
            apply_styles_to_sheet(writer.sheets[sheet_name], branch_name, store_code, current_date)
        summary_all = final_list.groupby(['Item No.', 'Description', 'UNIT'], sort=False)['Qty'].sum().reset_index()
        summary_all.insert(0, 'No', range(1, len(summary_all) + 1))
        summary_all.to_excel(writer, sheet_name="Summary_All", index=False, header=False, startrow=10)
        ws_sum = writer.sheets["Summary_All"]
        last_row = 10 + len(summary_all) + 1
        ws_sum.cell(row=last_row, column=3, value="Grand Total").font = Font(name='Cordia New', bold=True, size=14)
        ws_sum.cell(row=last_row, column=5, value=summary_all['Qty'].sum()).font = Font(name='Cordia New', bold=True, size=14)
        apply_styles_to_sheet(ws_sum, "สรุปยอดรวม", "ALL", current_date, is_summary=True)
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
    ws['G2'], ws['G4'] = f"Date: {current_date}", f"Delivery Date: {current_date}"
    ws['G2'].alignment = ws['G4'].alignment = Alignment(horizontal='right')
    ws['A7'], ws['C7'] = f"Code: {store_code}", f"Name: {branch_name}"
    ws['A7'].font = ws['C7'].font = f_header
    headers = ['No', 'Code', 'Name', 'Unit', 'TOTAL' if is_summary else 'ORDER', 'MBL', 'BNN']
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=10, column=i, value=h); c.font, c.fill, c.border, c.alignment = f_white, fill_green, border, Alignment(horizontal='center')
    for row in ws.iter_rows(min_row=11, max_row=ws.max_row, min_col=1, max_col=7):
        for cell in row: cell.border, cell.font = border, f_data
    widths = {'A': 6, 'B': 16, 'C': 35, 'D': 10, 'E': 12, 'F': 10, 'G': 10}
    for col, w in widths.items(): ws.column_dimensions[col].width = w

# --- จริตหน้าจอ (UI) ---
st.markdown("<h1 style='text-align: center; color: #ff4bad; font-size: 50px;'>✨💅🏻 คลังคุณนายเล็บเจล 💅🏻✨</h1>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("ส่งไฟล์มาให้แม่เนรมิตค่ะ", type="xlsx")

if uploaded_file:
    if st.button("💖 เนรมิตไฟล์แบบสับ! 💖", use_container_width=True):
        with st.spinner('เล็บเจลกำลังอบไฟ... แป๊บนึงนะแม่...'):
            st.session_state.excel_bytes = process_excel_to_buffer(uploaded_file)
            st.session_state.fabulous = True

    if st.session_state.get('fabulous'):
        # บังคับฉีดความเริ่ด
        inject_fabulous_nails()
        
        st.markdown("<div style='text-align: center; padding: 20px;'><h3>เสร็จแล้วค่ะคุณนาย! เล็บเจลพุ่งมั้ยคะ?</h3></div>", unsafe_allow_html=True)
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ความเริ่ด (กดแล้วลูกโป่งมาซ้ำอีกรอบ!)",
            data=st.session_state.excel_bytes,
            file_name=f"POD_Fabulous.xlsx",
            use_container_width=True,
            on_click=lambda: st.balloons()
        )