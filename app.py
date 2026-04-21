import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

st.set_page_config(page_title="Delivery Formatter Pro", layout="wide")

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
            
            ws = writer.sheets[sheet_name]
            
            # --- Styles (Cordia New) ---
            font_name = 'Cordia New'
            f_title = Font(name=font_name, bold=True, size=18)
            f_header = Font(name=font_name, bold=True, size=14)
            f_data = Font(name=font_name, size=12)
            f_white = Font(name=font_name, bold=True, color="FFFFFF", size=12)
            
            fill_green = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            fill_light = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid") # สีเว้นบรรทัด
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            # --- Header ---
            ws.merge_cells('A1:G1')
            ws['A1'] = "ใบส่งสินค้าชั่วคราว"; ws['A1'].font = f_title; ws['A1'].alignment = Alignment(horizontal='center')
            ws['A2'] = "บริษัท โมบาย โลจิสติกส์ จำกัด"; ws['A2'].font = f_header
            ws['A3'] = "278 หมู่ที่ 9 ตำบลบางโฉลง อ.บางพลี จ.สมุทรปราการ 10540"; ws['A3'].font = f_data
            ws['A4'] = "โทร. 02-337-1200 แฟกซ์. 02-337-1201"; ws['A4'].font = f_data
            
            ws['G2'] = f"Date: {current_date}"; ws['G2'].alignment = Alignment(horizontal='right'); ws['G2'].font = f_header
            ws['G3'] = "Zone: "; ws['G3'].alignment = Alignment(horizontal='right'); ws['G3'].font = f_header
            ws['G4'] = f"Delivery Date: {current_date}"; ws['G4'].alignment = Alignment(horizontal='right'); ws['G4'].font = f_header
            
            ws['A6'] = "Customer Name"; ws['A6'].font = f_header
            ws['A7'] = f"Store Code: {store_code}"; ws['A7'].font = f_header
            ws['C7'] = f"Store Name: {branch_name}"; ws['C7'].font = f_header

            # --- Table Header ---
            ws.merge_cells('E9:G9')
            ws['E9'] = "Qty"; ws['E9'].font = f_white; ws['E9'].fill = fill_green; ws['E9'].border = border; ws['E9'].alignment = Alignment(horizontal='center')
            
            headers = ['No', 'Product Code', 'Product Name', 'Unit', 'ORDER', 'MBL', 'BNN']
            for i, h in enumerate(headers, 1):
                cell = ws.cell(row=10, column=i, value=h)
                cell.font, cell.fill, cell.border = f_white, fill_green, border
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # --- Table Body (มีสีเว้นบรรทัด) ---
            for r_idx, row in enumerate(ws.iter_rows(min_row=11, max_row=ws.max_row, min_col=1, max_col=7), 1):
                for cell in row:
                    cell.border = border
                    cell.font = f_data
                    if r_idx % 2 == 0: cell.fill = fill_light # สลับสีบรรทัด
                    if cell.column in [1, 4, 5, 6, 7]: cell.alignment = Alignment(horizontal='center')

            # --- Footer (แยกบรรทัด + ตารางสรุปใหม่) ---
            start_f = ws.max_row + 2
            
            # ฝั่งซ้าย
            labels = ["ผู้รับสินค้า:", "ผู้ส่งสินค้า:", "ทะเบียนรถ:", "คลังสินค้า:"]
            for i, label in enumerate(labels):
                c = ws.cell(row=start_f + i, column=1, value=f"{label} .......................................................")
                c.font = f_header
            
            # ฝั่งขวา (ตารางสรุปเปลี่ยนชื่อตามสั่ง)
            summary_headers = ["", "MBL", "BNN"]
            for i, h in enumerate(summary_headers):
                c = ws.cell(row=start_f, column=5+i, value=h)
                c.font, c.fill, c.border = f_white, fill_green, border
                c.alignment = Alignment(horizontal='center')
            
            # เปลี่ยนคำ: ตะกร้าใหญ่ / ตะกร้าเล็ก
            summary_rows = ["ตะกร้าใหญ่", "ตะกร้าเล็ก"]
            for r_idx, label in enumerate(summary_rows, 1):
                c_label = ws.cell(row=start_f + r_idx, column=5, value=label)
                c_label.font, c_label.border = f_header, border
                ws.cell(row=start_f + r_idx, column=6).border = border
                ws.cell(row=start_f + r_idx, column=7).border = border

            # Print Settings
            ws.page_setup.paperSize = 9 # A4
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 1
            ws.print_options.horizontalCentered = True
            
            widths = {'A': 6, 'B': 15, 'C': 35, 'D': 10, 'E': 12, 'F': 10, 'G': 10}
            for col, w in widths.items(): ws.column_dimensions[col].width = w

    return output.getvalue()

# UI
st.title("🚚 อีเดียร์ขี้เกียจทำงาน")
uploaded_file = st.file_uploader("Upload Excel File", type="xlsx")

if uploaded_file:
    with st.spinner('เดียร์กำลังนั่งเกมดูหนังแต่ไม่ทำงานอยู่ รอแปป...'):
        excel_bytes = process_excel_to_buffer(uploaded_file)
        st.success("✅ เริ่ดเลยหล่ะะะะะ!")
        st.download_button(label="📥 Download Excel ที่ปรับปรุงแล้ว", data=excel_bytes, file_name=f"Delivery_Final_{datetime.now().strftime('%H%M')}.xlsx")