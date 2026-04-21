import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime
import streamlit.components.v1 as components

st.set_page_config(page_title="POD BNN - Fabulous Edition", layout="wide")

# --- ฟังก์ชันเอฟเฟกต์ 💅🏻 แบบบังคับพุ่ง ---
def show_fabulous_popup():
    # เราใช้ Components.html เพื่อรัน JS แยกส่วน บังคับให้แสดงผล
    components.html(
        """
        <script src="https://cdn.jsdelivr.net/npm/canvas-confetti@1.5.1/dist/confetti.browser.min.js"></script>
        <div id="fabulous-container" style="display:flex; justify-content:center; align-items:center; height:100vh;">
            <h1 id="text" style="
                font-family: 'Tahoma', sans-serif; 
                font-size: 100px; 
                color: #ff4bad; 
                text-shadow: 5px 5px 15px rgba(0,0,0,0.3);
                opacity: 0;
                transition: all 0.5s ease-out;
                transform: scale(0.5);
            ">เริ่ดเลยหล่ะ</h1>
        </div>
        <script>
            const text = document.getElementById('text');
            setTimeout(() => {
                text.style.opacity = '1';
                text.style.transform = 'scale(1.2)';
            }, 100);

            var end = Date.now() + (3 * 1000);
            var scalar = 4;
            var nail = confetti.shapeFromText({ text: '💅🏻', scalar });

            (function frame() {
                confetti({
                    particleCount: 10,
                    angle: 60,
                    spread: 70,
                    origin: { x: 0, y: 0.8 },
                    shapes: [nail],
                    scalar
                });
                confetti({
                    particleCount: 10,
                    angle: 120,
                    spread: 70,
                    origin: { x: 1, y: 0.8 },
                    shapes: [nail],
                    scalar
                });
                if (Date.now() < end) {
                    requestAnimationFrame(frame);
                }
            }());
        </script>
        """,
        height=300, # กำหนดความสูงให้มองเห็นตัวหนังสือ
    )

# --- ฟังก์ชันประมวลผล Excel (เหมือนเดิม) ---
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

        # หน้าสรุป
        summary_all = final_list.groupby(['Item No.', 'Description', 'UNIT'], sort=False)['Qty'].sum().reset_index()
        summary_all.insert(0, 'No', range(1, len(summary_all) + 1))
        summary_all['MBL'] = ""; summary_all['BNN'] = ""
        sheet_name_sum = "Summary_All"
        summary_all.to_excel(writer, sheet_name=sheet_name_sum, index=False, header=False, startrow=10)
        
        ws_sum = writer.sheets[sheet_name_sum]
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
    ws['A1'] = "ใบสรุปรายการเบิกสินค้า" if is_summary else "ใบส่งสินค้าชั่วคราว"
    ws['A1'].font, ws['A1'].alignment = f_title, Alignment(horizontal='center')
    ws['A2'], ws['A3'], ws['A4'] = "บริษัท โมบาย โลจิสติกส์ จำกัด", "278 หมู่ที่ 9 ตำบลบางโฉลง อ.บางพลี จ.สมุทรปราการ 10540", "โทร. 02-337-1200"
    ws['A2'].font, ws['A3'].font, ws['A4'].font = f_header, f_data, f_data
    ws['G2'], ws['G4'] = f"Date: {current_date}", f"Delivery Date: {current_date}"
    ws['G2'].alignment = ws['G4'].alignment = Alignment(horizontal='right')
    ws['G2'].font = ws['G4'].font = f_header
    ws['A7'], ws['C7'] = f"Code: {store_code}", f"Name: {branch_name}"
    ws['A7'].font = ws['C7'].font = f_header

    ws.merge_cells('E9:G9')
    ws['E9'] = "Total Qty" if is_summary else "Qty"
    ws['E9'].font, ws['E9'].fill, ws['E9'].border, ws['E9'].alignment = f_white, fill_green, border, Alignment(horizontal='center')
    
    for i, h in enumerate(['No', 'Product Code', 'Product Name', 'Unit', 'TOTAL' if is_summary else 'ORDER', 'MBL', 'BNN'], 1):
        c = ws.cell(row=10, column=i, value=h); c.font, c.fill, c.border, c.alignment = f_white, fill_green, border, Alignment(horizontal='center')

    for r_idx, row in enumerate(ws.iter_rows(min_row=11, max_row=ws.max_row, min_col=1, max_col=7), 1):
        if ws.cell(row=row[0].row, column=3).value == "Grand Total (ยอดรวมทั้งหมด)": continue
        for cell in row:
            cell.border, cell.font = border, f_data
            if r_idx % 2 == 0: cell.fill = fill_light
            if cell.column in [1, 4, 5, 6, 7]: cell.alignment = Alignment(horizontal='center')

    if not is_summary:
        curr_row = ws.max_row + 2
        for i, label in enumerate(["ผู้รับสินค้า:", "ผู้ส่งสินค้า:", "ทะเบียนรถ:", "คลังสินค้า:"]):
            ws.cell(row=curr_row + i, column=1, value=f"{label} ...........................................").font = f_header
        for i, label in enumerate(["ตะกร้าใหญ่", "ตะกร้าเล็ก"], 1):
            ws.cell(row=curr_row + i, column=5, value=label).font, ws.cell(row=curr_row + i, column=5).border = f_header, border
            ws.cell(row=curr_row + i, column=6).border = ws.cell(row=curr_row + i, column=7).border = border

    ws.page_setup.paperSize = 9
    widths = {'A': 6, 'B': 16, 'C': 35, 'D': 10, 'E': 12, 'F': 10, 'G': 10}
    for col, w in widths.items(): ws.column_dimensions[col].width = w

# --- UI ---
st.markdown("<h1 style='text-align: center; color: #2E7D32;'>💅🏻 POD FABULOUS GENERATOR 💅🏻</h1>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload ไฟล์ Excel ที่นี่", type="xlsx")

if uploaded_file:
    if st.button("✨ เริ่มความเริ่ด (Process Data) ✨", use_container_width=True):
        with st.spinner('ความสวยกำลังเดินทางมา...'):
            st.session_state.excel_bytes = process_excel_to_buffer(uploaded_file)
            st.session_state.done = True

    if st.session_state.get('done'):
        # แสดงเอฟเฟกต์ 💅🏻 พุ่งกระจาย
        show_fabulous_popup()
        
        st.download_button(
            label="💅🏻 ดาวน์โหลดไฟล์ที่เริ่ดที่สุดในโลก 💅🏻",
            data=st.session_state.excel_bytes,
            file_name=f"POD_BNN_Fabulous.xlsx",
            use_container_width=True,
            on_click=lambda: st.balloons()
        )