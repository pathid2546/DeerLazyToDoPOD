import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime
import io

# --- ฟังก์ชันสร้าง Excel ---
def create_excel(final_list, customer_name, branch_to_code):
    current_date = datetime.now().strftime('%d/%m/%Y')
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
            f_bold = Font(name='Sarabun', bold=True, size=10)
            f_norm = Font(name='Sarabun', size=10)
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            fill = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            f_white = Font(name='Sarabun', bold=True, color="FFFFFF", size=10)

            # Header
            ws['A1'] = "ใบส่งสินค้าชั่วคราว"; ws['A1'].font = Font(bold=True, size=16); ws.merge_cells('A1:G1'); ws['A1'].alignment = Alignment(horizontal='center')
            ws['A6'] = f"Customer: {customer_name}"; ws['A7'] = f"Store Code: {store_code}"; ws['C7'] = f"Store Name: {branch_name}"
            
            # --- จัดการหัวตารางให้ Qty อยู่ตรงกลางเป๊ะ ---
            # 1. Merge Qty ครอบคลุม ORDER, MBL, BNN (คอลัมน์ E-G)
            ws.merge_cells(start_row=9, start_column=5, end_row=9, end_column=7)
            ws.cell(row=9, column=5, value="Qty").alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(row=9, column=5).font = f_white
            ws.cell(row=9, column=5).fill = fill

            # 2. หัวตารางชั้นล่าง
            sub_h = ['No', 'Product Code', 'Product Name', 'Unit', 'ORDER', 'MBL', 'BNN']
            for i, h in enumerate(sub_h, 1):
                cell = ws.cell(row=10, column=i, value=h)
                cell.fill, cell.font, cell.border = fill, f_white, border
                cell.alignment = Alignment(horizontal='center')
                # สำหรับคอลัมน์ 1-4 ให้ Merge แนวตั้ง
                if i <= 4:
                    ws.merge_cells(start_row=9, start_column=i, end_row=10, end_column=i)
                    ws.cell(row=9, column=i, value=h).alignment = Alignment(horizontal='center', vertical='center')

            # เส้นขอบข้อมูล
            for row in ws.iter_rows(min_row=11, max_row=ws.max_row, min_col=1, max_col=7):
                for cell in row:
                    cell.border = border
            
            # Footer ลายเซ็น
            f_row = ws.max_row + 2
            ws.cell(row=f_row, column=1, value="ลงชื่อ ......................................... ผู้ส่งสินค้า")
            ws.cell(row=f_row, column=5, value="ลงชื่อ ......................................... ผู้ตรวจสอบ")

    return output.getvalue()

# --- หน้าจอหลัก Streamlit ---
st.title("🚚 Delivery Generator (A4 & PDF)")

uploaded_file = st.file_uploader("Upload Excel File", type="xlsx")

if uploaded_file:
    # (โค้ดส่วนประมวลผลข้อมูล raw_df เหมือนเดิม)
    raw_df = pd.read_excel(uploaded_file, header=None)
    customer_name = str(raw_df.iloc[1, 0])
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in row.values)
    
    # ... (ส่วนเตรียมข้อมูล final_list และ branch_to_code เหมือนเดิม) ...
    # [เพื่อให้โค้ดกระชับ ผมละส่วนเตรียม DataFrame ไว้ แต่ในไฟล์จริงต้องมีครบครับ]

    # ปุ่มดาวน์โหลด Excel
    excel_data = create_excel(final_list, customer_name, branch_to_code)
    st.download_button("📥 Download Excel (Qty Centered)", excel_data, "delivery.xlsx")
    
    st.info("💡 สำหรับ PDF: แนะนำให้เปิดไฟล์ Excel ที่โหลดไป แล้วกด 'Save as PDF' หรือ 'Print to PDF' จะได้ผลลัพธ์ที่จัดหน้า A4 ได้สวยงามที่สุดครับ")