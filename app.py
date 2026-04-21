import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

def process_excel_to_buffer(uploaded_file):
    # 1. อ่านข้อมูลเหมือนเดิม
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
            
            # ข้อมูลสินค้า
            data_to_write = pd.DataFrame({
                'No': range(1, len(branch_data) + 1),
                'Code': branch_data['Item No.'],
                'Name': branch_data['Description'],
                'Unit': branch_data['UNIT'],
                'ORD': branch_data['Qty'],
                'MBL': "",
                'BNN': ""
            })
            data_to_write.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=10)
            
            ws = writer.sheets[sheet_name]
            f_white = Font(name='Sarabun', bold=True, color="FFFFFF", size=10)
            fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            # --- แก้จุดที่ Qty ไม่ตรงกลาง ---
            # 1. Merge คอลัมน์ E, F, G แถวที่ 9
            ws.merge_cells(start_row=9, start_column=5, end_row=9, end_column=7)
            cell_qty = ws.cell(row=9, column=5, value="Qty")
            cell_qty.font = f_white
            cell_qty.fill = fill
            cell_qty.alignment = Alignment(horizontal='center', vertical='center') # จัดกึ่งกลางเป๊ะ
            cell_qty.border = border
            
            # เติมเส้นขอบให้เซลล์ที่ถูก Merge (F9, G9)
            ws.cell(row=9, column=6).border = border
            ws.cell(row=9, column=7).border = border

            # 2. หัวตารางชั้นที่ 2 (ORDER, MBL, BNN)
            sub_headers = ['ORDER', 'MBL', 'BNN']
            for i, sh in enumerate(sub_headers, 5):
                cell = ws.cell(row=10, column=i, value=sh)
                cell.font, cell.fill, cell.border = f_white, fill, border
                cell.alignment = Alignment(horizontal='center')

            # 3. Merge แนวตั้งสำหรับ No, Code, Name, Unit
            main_headers = ['No', 'Product Code', 'Product Name', 'Unit']
            for i, h in enumerate(main_headers, 1):
                ws.merge_cells(start_row=9, start_column=i, end_row=10, end_column=i)
                cell = ws.cell(row=9, column=i, value=h)
                cell.font, cell.fill, cell.border = f_white, fill, border
                cell.alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(row=10, column=i).border = border

            # จัดการความกว้างและ PDF Layout
            ws.page_setup.paperSize = ws.page_setup.PAPERSIZE_A4
            ws.column_dimensions['C'].width = 35

    return output.getvalue()

# --- Streamlit UI ---
st.title("Delivery Formatter (A4 & Center Qty)")
file = st.file_uploader("อัปโหลดไฟล์ Excel", type="xlsx")

if file:
    excel_bytes = process_excel_to_buffer(file)
    st.success("จัดรูปแบบ Qty ตรงกลางเรียบร้อยแล้ว!")
    
    st.download_button(
        label="📥 ดาวน์โหลดไฟล์ Excel (A4 Ready)",
        data=excel_bytes,
        file_name="delivery_note_formatted.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.warning("⚠️ สำหรับ PDF: เนื่องจากข้อจำกัดของระบบ Cloud แนะนำให้โหลดไฟล์ Excel แล้วกด 'Save as PDF' ในคอมพิวเตอร์ จะได้หน้าที่สวยและตรงที่สุดครับ")