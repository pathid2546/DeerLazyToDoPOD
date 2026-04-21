import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime
import streamlit.components.v1 as components

# --- ตั้งค่าหน้าจอแบบหรูหรา ---
st.set_page_config(page_title="POD BNN - Luxury Edition", layout="wide")

# CSS สำหรับจริตตัวแม่แบบสุภาพ (Rose Gold Theme)
st.markdown("""
    <style>
    .main { background-color: #ffffff; }
    .stButton>button {
        background-color: #d4a373;
        color: white;
        border-radius: 20px;
        border: none;
        padding: 10px 25px;
        font-weight: bold;
        transition: 0.3s;
    }
    .stButton>button:hover {
        background-color: #bc8a5f;
        color: #fff;
        transform: translateY(-2px);
    }
    h1 { color: #bc8a5f; font-family: 'Tahoma', sans-serif; }
    .success-text { color: #2e7d32; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

# --- ฟังก์ชันเอฟเฟกต์ "คุณนายเล็บเจล" (สุภาพแต่เริ่ด) ---
def show_elegant_celebration():
    components.html(
        """
        <script src="https://cdn.jsdelivr.net/npm/canvas-confetti@1.5.1/dist/confetti.browser.min.js"></script>
        <script>
            var duration = 3 * 1000;
            var end = Date.now() + duration;

            (function frame() {
                confetti({
                    particleCount: 2,
                    angle: 60,
                    spread: 55,
                    origin: { x: 0 },
                    colors: ['#d4a373', '#ffccd5', '#ffffff'],
                    shapes: [confetti.shapeFromText({ text: '💅🏻', scalar: 3 })]
                });
                confetti({
                    particleCount: 2,
                    angle: 120,
                    spread: 55,
                    origin: { x: 1 },
                    colors: ['#d4a373', '#ffccd5', '#ffffff'],
                    shapes: [confetti.shapeFromText({ text: '✨', scalar: 3 })]
                });

                if (Date.now() < end) {
                    requestAnimationFrame(frame);
                }
            }());
        </script>
        """,
        height=0,
    )

# --- หัวใจหลัก: การประมวลผลข้อมูล (Logic เดิมเป๊ะ) ---
def process_excel_to_buffer(uploaded_file):
    raw_df = pd.read_excel(uploaded_file, header=None)
    # หาแถว Item No.
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in row.values)
    
    # ดึงข้อมูลสาขาและรหัสสาขา
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
        # 1. หน้าแยกสาขา
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

        # 2. หน้าสรุปยอดรวม (Summary All)
        summary_all = final_list.groupby(['Item No.', 'Description', 'UNIT'], sort=False)['Qty'].sum().reset_index()
        summary_all.insert(0, 'No', range(1, len(summary_all) + 1))
        summary_all['MBL'] = ""; summary_all['BNN'] = ""
        
        sheet_name_sum = "Summary_All"
        summary_all.to_excel(writer, sheet_name=sheet_name_sum, index=False, header=False, startrow=10)
        ws_sum = writer.sheets[sheet_name_sum]
        
        # Grand Total
        last_row = 10 + len(summary_all) + 1
        ws_sum.cell(row=last_row, column=3, value="Grand Total (ยอดรวมทั้งหมด)").font = Font(name='Cordia New', bold=True, size=14)
        ws_sum.cell(row=last_row, column=3).alignment = Alignment(horizontal='right')
        
        qty_total = ws_sum.cell(row=last_row, column=5, value=summary_all['Qty'].sum())
        qty_total.font = Font(name='Cordia New', bold=True, size=14)
        qty_total.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        qty_total.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='double'))
        qty_total.alignment = Alignment(horizontal='center')

        apply_styles_to_sheet(ws_sum, "สรุปยอดรวมทุกรายการ", "ALL", current_date, is_summary=True)

    return output.getvalue()

def apply_styles_to_sheet(ws, branch_name, store_code, current_date, is_summary=False):
    font_name = 'Cordia New'
    f_title = Font(name=font_name, bold=True, size=20)
    f_header = Font(name=font_name, bold=True, size=14)
    f_data = Font(name=font_name, size=14)
    f_white = Font(name=font_name, bold=True, color="FFFFFF", size=14)
    fill_main = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
    fill_stripe = PatternFill(start_color="F9F9F9", end_color="F9F9F9", fill_type="solid")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Header Section
    ws.merge_cells('A1:G1')
    ws['A1'] = "ใบสรุปรายการเบิกสินค้า" if is_summary else "ใบส่งสินค้าชั่วคราว"
    ws['A1'].font = f_title; ws['A1'].alignment = Alignment(horizontal='center')
    ws['A2'] = "บริษัท โมบาย โลจิสติกส์ จำกัด"; ws['A2'].font = f_header
    ws['G2'] = f"Date: {current_date}"; ws['G2'].alignment = Alignment(horizontal='right'); ws['G2'].font = f_header
    ws['A7'] = f"Code: {store_code}"; ws['C7'] = f"Name: {branch_name}"
    ws['A7'].font = ws['C7'].font = f_header

    # Table Header
    ws.merge_cells('E9:G9')
    ws['E9'] = "Total Qty" if is_summary else "Qty"
    ws['E9'].font, ws['E9'].fill, ws['E9'].border = f_white, fill_main, border
    ws['E9'].alignment = Alignment(horizontal='center')
    
    col_headers = ['No', 'Product Code', 'Product Name', 'Unit', 'TOTAL' if is_summary else 'ORDER', 'MBL', 'BNN']
    for i, text in enumerate(col_headers, 1):
        cell = ws.cell(row=10, column=i, value=text)
        cell.font, cell.fill, cell.border = f_white, fill_main, border
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Data Body
    for r_idx, row in enumerate(ws.iter_rows(min_row=11, max_row=ws.max_row, min_col=1, max_col=7), 1):
        if ws.cell(row=row[0].row, column=3).value == "Grand Total (ยอดรวมทั้งหมด)": continue
        for cell in row:
            cell.border = border; cell.font = f_data
            if r_idx % 2 == 0: cell.fill = fill_stripe
            if cell.column in [1, 4, 5, 6, 7]: cell.alignment = Alignment(horizontal='center')

    # Footer
    if not is_summary:
        curr_row = ws.max_row + 2
        labels = ["ผู้รับสินค้า:", "ผู้ส่งสินค้า:", "ทะเบียนรถ:"]
        for i, label in enumerate(labels):
            ws.cell(row=curr_row + i, column=1, value=f"{label} ....................................").font = f_header

    # Page Setup
    ws.page_setup.paperSize = 9
    for col, width in zip('ABCDEFG', [6, 16, 35, 10, 12, 10, 10]):
        ws.column_dimensions[col].width = width

# --- UI ส่วนหน้าบ้าน ---
st.markdown("<h1 style='text-align: center;'>✨ POD BNN Luxury Management ✨</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: #666;'>จัดการข้อมูลแบบตัวแม่ ข้อมูลเป๊ะ งานเริ่ด สบายตา</p>", unsafe_allow_html=True)

file = st.file_uploader("กรุณาเลือกไฟล์ Excel ที่ต้องการเนรมิตค่ะแม่", type="xlsx")

if file:
    # เพิ่ม Button จริตคุณนาย
    if st.button("✨ เริ่มการเนรมิตข้อมูล ✨", use_container_width=True):
        with st.spinner('กรุณารอสักครู่ค่ะแม่ ข้อมูลกำลังถูกจัดระเบียบให้สวยงาม...'):
            try:
                st.session_state.result_bytes = process_excel_to_buffer(file)
                st.session_state.is_done = True
            except Exception as e:
                st.error(f"อุ๊ย! เกิดข้อผิดพลาดนิดหน่อยค่ะแม่: {e}")

    if st.session_state.get('is_done'):
        # เอฟเฟกต์ความเริ่ด
        show_elegant_celebration()
        
        st.markdown("<p class='success-text' style='text-align: center;'>💅🏻 เริ่ดเลยหล่ะค่ะ! ข้อมูลตรวจสอบเรียบร้อย พร้อมใช้งานแล้วนะคะ</p>", unsafe_allow_html=True)
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ที่เนรมิตเสร็จแล้ว (Download)",
            data=st.session_state.result_bytes,
            file_name=f"POD_Elegant_{datetime.now().strftime('%H%M')}.xlsx",
            use_container_width=True
        )