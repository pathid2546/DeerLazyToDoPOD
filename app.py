import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime
import streamlit.components.v1 as components

st.set_page_config(page_title="POD BNN - จริตตัวแม่เล็บเจล", layout="wide")

# --- ท่าไม้ตายฉีดเล็บเจลพุ่ง (Full Screen Fabulousness) ---
def inject_fabulous_nails():
    # ใช้ HTML + JS แบบจัดเต็ม บังคับให้พุ่งออกมาจากหน้าจอ
    components.html(
        """
        <div id="fabulous-wrapper" style="position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; pointer-events: none; z-index: 9999;">
            <script src="https://cdn.jsdelivr.net/npm/canvas-confetti@1.5.1/dist/confetti.browser.min.js"></script>
            <script>
                function fireNails() {
                    var end = Date.now() + (4 * 1000); // พุ่งยาวๆ 4 วินาที
                    var nail = confetti.shapeFromText({ text: '💅🏻', scalar: 6 });
                    var sparkle = confetti.shapeFromText({ text: '✨', scalar: 5 });

                    (function frame() {
                        confetti({
                            particleCount: 10,
                            angle: 60,
                            spread: 55,
                            origin: { x: 0, y: 0.8 },
                            shapes: [nail, sparkle],
                            colors: ['#ff4bad', '#ffecf5']
                        });
                        confetti({
                            particleCount: 10,
                            angle: 120,
                            spread: 55,
                            origin: { x: 1, y: 0.8 },
                            shapes: [nail, sparkle],
                            colors: ['#ff4bad', '#ffecf5']
                        });

                        if (Date.now() < end) {
                            requestAnimationFrame(frame);
                        }
                    }());
                }
                fireNails();
            </script>
            <div style="
                position: absolute;
                top: 40%;
                left: 50%;
                transform: translate(-50%, -50%);
                text-align: center;
                width: 100%;
            ">
                <h1 style="
                    font-family: 'Tahoma', sans-serif;
                    font-size: 150px;
                    color: #ff4bad;
                    text-shadow: 10px 10px 0px #fff, 15px 15px 40px rgba(255, 75, 173, 0.8);
                    animation: zoomIn 0.5s cubic-bezier(0.175, 0.885, 0.32, 1.275) forwards;
                    margin: 0;
                ">เริ่ดเลยหล่ะ</h1>
                <div style="font-size: 40px; color: #ff4bad; font-weight: bold; background: white; display: inline-block; padding: 10px 30px; border-radius: 50px; border: 3px solid #ff4bad; margin-top: 20px;">
                    💅🏻 สวย สับ ระดับตัวแม่! ✨
                </div>
            </div>
        </div>
        <style>
            @keyframes zoomIn {
                from { opacity: 0; transform: scale(0.5); }
                to { opacity: 1; transform: scale(1.1); }
            }
        </style>
        """,
        height=600, # จองที่ให้ความเริ่ดได้แสดงตัว
    )

# --- ส่วนประมวลผล Excel (เหมือนเดิมเป๊ะ) ---
def process_excel_to_buffer(uploaded_file):
    raw_df = pd.read_excel(uploaded_file, header=None)
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in row.values)
    
    header_row_raw = raw_df.iloc[header_row_index].astype(str).str.strip().tolist()
    code_row_raw = raw_df.iloc[header_row_index + 1].tolist()
    branch_to_code = {name: str(code) if pd.notna(code) else "" 
                      for name, code in zip(header_row_raw, code_row_raw) if name and name != 'nan'}

    df = pd.read_excel(uploaded_file, header=header_row_index)
    df = df.iloc[1:].reset_index(drop=True)
    df.columns = df.columns.str.strip()

    id_cols = ['Item No.', 'Description', 'UNIT']
    df_melted = df.melt(id_vars=id_cols, var_name='Branch', value_name='Qty')
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
    ws['A2'], ws['A3'], ws['A4'] = "บริษัท โมบาย โลจิสติกส์ จำกัด", "278 หมู่ที่ 9 จ.สมุทรปราการ 10540", "โทร. 02-337-1200"
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

# --- จริตหน้าจอ (UI) ---
st.markdown("<h1 style='text-align: center; color: #ff4bad; font-size: 60px;'>✨💅🏻 คลังคุณนายเล็บเจล 💅🏻✨</h1>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("ส่งไฟล์มาให้แม่จัดการค่ะ", type="xlsx")

if uploaded_file:
    if st.button("💖 เริ่มความเริ่ด (Process Data) 💖", use_container_width=True):
        with st.spinner('เล็บเจลกำลังแห้ง... แป๊บนึงนะแม่...'):
            st.session_state.excel_bytes = process_excel_to_buffer(uploaded_file)
            st.session_state.fabulous = True

    if st.session_state.get('fabulous'):
        # บังคับพ่นเล็บเจล 💅🏻 พุ่งกระจาย
        inject_fabulous_nails()
        
        st.markdown("<div style='text-align: center; padding: 20px;'><h3>เสร็จแล้วค่ะแม่! เล็บเจลพุ่งมั้ยคะ?</h3></div>", unsafe_allow_html=True)
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ความเริ่ด 📥",
            data=st.session_state.excel_bytes,
            file_name=f"POD_Fabulous_{datetime.now().strftime('%H%M')}.xlsx",
            use_container_width=True
        )