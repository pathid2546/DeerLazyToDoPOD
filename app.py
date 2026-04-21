import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

def process_excel_to_buffer(uploaded_file):
    # อ่านข้อมูล
    raw_df = pd.read_excel(uploaded_file, header=None)
    customer_name = str(raw_df.iloc[1, 0])
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in row.values)
    
    # Mapping Store Code
    header_row_raw = raw_df.iloc[header_row_index].astype(str).str.strip().tolist()
    code_row_raw = raw_df.iloc[header_row_index + 1].tolist()
    branch_to_code = {name: str(code) if pd.notna(code) else "" 
                      for name, code in zip(header_row_raw, code_row_raw) if name and name != 'nan'}

    # เตรียมข้อมูลสินค้า
    df = pd.read_excel(uploaded_file, header=header_row_index)
    df = df.iloc[1:].reset_index(drop=True)
    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
    df.columns = df.columns.str.strip()

    id_cols = ['Item No.', 'Description', 'UNIT']
    df_melted = df.melt(id_vars=id_cols, var_name='Branch', value_name='Qty')
    df_melted['Qty'] = pd.to_numeric(df_melted['Qty'], errors='coerce')
    final_list = df_melted.dropna(subset=['Item No.', 'Qty'])
    final_list = final_list[final_list['Qty'] > 0]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for branch_name, branch_data in final_list.groupby('Branch', sort=False):
            store_code = branch_to_code.get(str(branch_name).strip(), "")
            sheet_name = str(branch_name)[:30].replace('/', '-').replace(':', '')
            
            data_to_write = pd.DataFrame({
                'No': range(1, len(branch_data) + 1),
                'Code': branch_data['Item No.'],
                'Name': branch_data['Description'],
                'Unit': branch_data['UNIT'],
                'ORD': branch_data['Qty'],
                'MBL': "", 'BNN': ""
            })
            data_to_write.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=10)
            
            ws = writer.sheets[sheet_name]
            f_white = Font(name='Sarabun', bold=True, color="FFFFFF", size=10)
            f_norm = Font(name='Sarabun', size=10)
            fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            # --- Qty Center ---
            ws.merge_cells(start_row=9, start_column=5, end_row=9, end_column=7)
            cell_qty = ws.cell(row=9, column=5, value="Qty")
            cell_qty.font = f_white; cell_qty.fill = fill; cell_qty.border = border
            cell_qty.alignment = Alignment(horizontal='center', vertical='center')

            # หัวตารางแถวที่ 10
            sub_h = ['ORDER', 'MBL', 'BNN']
            for i, h in enumerate(sub_h, 5):
                c = ws.cell(row=10, column=i, value=h)
                c.font = f_white; c.fill = fill; c.border = border; c.alignment = Alignment(horizontal='center')

            # Merge แนวตั้งคอลัมน์ 1-4
            main_h = ['No', 'Product Code', 'Product Name', 'Unit']
            for i, h in enumerate(main_h, 1):
                ws.merge_cells(start_row=9, start_column=i, end_row=10, end_column=i)
                c = ws.cell(row=9, column=i, value=h)
                c.font = f_white; c.fill = fill; c.border = border; c.alignment = Alignment(horizontal='center', vertical='center')

            # --- ตั้งค่า A4 และแนวตั้ง (แบบที่ Server ยอมรับ) ---
            ws.page_setup.paperSize = 9 # ใส่เลข 9 ตรงๆ สำหรับ A4
            ws.page_setup.orientation = 'portrait' # ใส่เป็น Text ตามที่ Error แจ้ง
            
            ws.page_margins.left = 0.5
            ws.page_margins.right = 0.5

            # ลายเซ็น
            f_row = ws.max_row + 2
            ws.cell(row=f_row, column=1, value="ลงชื่อ ......................................... ผู้ส่งสินค้า").font = f_norm
            ws.cell(row=f_row, column=5, value="ลงชื่อ ......................................... ผู้ตรวจสอบ").font = f_norm

            # ปรับความกว้างคอลัมน์
            ws.column_dimensions['C'].width = 35

    return output.getvalue()

# UI Streamlit
st.title("🚚 Delivery Generator (Fix Orientation Value)")
file = st.file_uploader("Upload Excel", type="xlsx")

if file:
    try:
        excel_bytes = process_excel_to_buffer(file)
        st.success("สำเร็จ! แก้ไขค่า Orientation เป็น 'portrait' เรียบร้อย")
        st.download_button("📥 Download Excel", excel_bytes, "delivery_note.xlsx")
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")