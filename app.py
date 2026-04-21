import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime
import io

# --- ฟังก์ชันจัดการ Excel ---
def process_excel(uploaded_file):
    current_date = datetime.now().strftime('%d/%m/%Y')
    raw_df = pd.read_excel(uploaded_file, header=None)
    customer_name = str(raw_df.iloc[1, 0])
    
    # หาตำแหน่ง Header
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in row.values)
    
    # Mapping Store Code
    header_row_raw = raw_df.iloc[header_row_index].astype(str).str.strip().tolist()
    code_row_raw = raw_df.iloc[header_row_index + 1].tolist()
    branch_to_code = {name: str(code) if pd.notna(code) else "" 
                      for name, code in zip(header_row_raw, code_row_raw) if name and name != 'nan'}

    # เตรียมข้อมูลสินค้า
    df = pd.read_excel(uploaded_file, header=header_row_index)
    df['original_index'] = df.index
    df = df.iloc[1:].reset_index(drop=True)
    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
    df.columns = df.columns.str.strip()

    id_cols = ['Item No.', 'Description', 'UNIT', 'original_index']
    df_melted = df.melt(id_vars=id_cols, var_name='Branch', value_name='Qty')
    df_melted['Qty'] = pd.to_numeric(df_melted['Qty'], errors='coerce')
    final_list = df_melted.dropna(subset=['Item No.', 'Qty'])
    final_list = final_list[final_list['Qty'] > 0]

    # สร้าง Excel ใน Memory (BytesIO)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for branch_name, branch_data in final_list.groupby('Branch', sort=False):
            store_code = branch_to_code.get(str(branch_name).strip(), "")
            branch_sorted = branch_data.sort_values(by='original_index')
            sheet_name = str(branch_name)[:30].replace('/', '-').replace(':', '')
            
            data_to_write = pd.DataFrame({
                'No': range(1, len(branch_sorted) + 1),
                'Code': branch_sorted['Item No.'],
                'Name': branch_sorted['Description'],
                'Unit': branch_sorted['UNIT'],
                'ORD': branch_sorted['Qty'],
                'MBL': "", 'BNN': ""
            })
            data_to_write.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=10)
            
            ws = writer.sheets[sheet_name]
            
            # Styles
            f_norm = Font(name='Sarabun', size=10); f_bold = Font(name='Sarabun', bold=True, size=10)
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            f_white = Font(name='Sarabun', bold=True, color="FFFFFF", size=10)

            # Header & Table Construction (เหมือนเดิมที่คุณต้องการเป๊ะๆ)
            ws['A1'] = "ใบส่งสินค้าชั่วคราว"; ws['A1'].font = Font(bold=True, size=16); ws.merge_cells('A1:G1'); ws['A1'].alignment = Alignment(horizontal='center')
            ws['A2'] = "บริษัท โมบาย โลจิสติกส์ จำกัด"; ws.merge_cells('A2:D2')
            ws['G2'] = f"Date: {current_date}"; ws['G3'] = "Zone: "; ws['G4'] = f"Delivery: {current_date}"
            ws['A6'] = f"Customer: {customer_name}"; ws['A7'] = f"Store Code: {store_code}"; ws['C7'] = f"Store Name: {branch_name}"
            
            # Merge Header Qty / ORDER / MBL / BNN
            ws.merge_cells('E9:G9'); ws['E9'] = "Qty"
            sub_h = ['No', 'Product Code', 'Product Name', 'Unit', 'ORDER', 'MBL', 'BNN']
            for i, h in enumerate(sub_h, 1):
                cell = ws.cell(row=10, column=i, value=h)
                cell.fill, cell.font, cell.border = fill, f_white, border
                cell.alignment = Alignment(horizontal='center')
            
            # Footer (ไม่มีหมายเหตุ)
            f_row = ws.max_row + 2
            ws.cell(row=f_row, column=1, value="ลงชื่อ ......................................... ผู้ส่งสินค้า")
            ws.cell(row=f_row+3, column=1, value="ลงชื่อ ......................................... ผู้รับสินค้า")
            ws.cell(row=f_row, column=5, value="ลงชื่อ ......................................... ผู้ตรวจสอบ")

    return output.getvalue()

# --- หน้าจอ Streamlit UI ---
st.set_page_config(page_title="Delivery Note Generator", layout="centered")
st.title("🚚 Delivery Note Generator")
st.write("อัปโหลดไฟล์ Excel เพื่อจัดฟอร์แมตใบส่งสินค้า (A4)")

uploaded_file = st.file_uploader("เลือกไฟล์ Excel (.xlsx)", type="xlsx")

if uploaded_file:
    with st.spinner('กำลังประมวลผล...'):
        processed_data = process_excel(uploaded_file)
        st.success("จัดฟอร์แมตเสร็จเรียบร้อย!")
        
        st.download_button(
            label="📥 ดาวน์โหลดไฟล์ Excel (จัดฟอร์แมตแล้ว)",
            data=processed_data,
            file_name=f"Delivery_Note_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )