import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from datetime import datetime

# ตั้งค่าหน้าเว็บ Streamlit
st.set_page_config(page_title="Delivery Generator", layout="wide")

def process_excel_to_buffer(uploaded_file):
    # 1. อ่านข้อมูล
    raw_df = pd.read_excel(uploaded_file, header=None)
    customer_name = str(raw_df.iloc[1, 0])
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in row.values)
    
    header_row_raw = raw_df.iloc[header_row_index].astype(str).str.strip().tolist()
    code_row_raw = raw_df.iloc[header_row_index + 1].tolist()
    branch_to_code = {name: str(code) if pd.notna(code) else "" 
                      for name, code in zip(header_row_raw, code_row_raw) if name and name != 'nan'}

    df = pd.read_excel(uploaded_file, header=header_row_index)
    df = df.iloc[1:].reset_index(drop=True)
    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]
    df.columns = df.columns.str.strip()

    id_cols = ['Item No.', 'Description', 'UNIT']
    df_melted = df.melt(id_vars=id_cols, var_name='Branch', value_name='Qty')
    df_melted['Qty'] = pd.to_numeric(df_melted['Qty'], errors='coerce')
    final_list = df_melted.dropna(subset=['Item No.', 'Qty'])
    final_list = final_list[final_list['Qty'] > 0]

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
            
            # --- Styles (ใช้ Cordia New) ---
            f_title = Font(name='Cordia New', bold=True, size=20)
            f_bold = Font(name='Cordia New', bold=True, size=14)
            f_norm = Font(name='Cordia New', size=14)
            f_white = Font(name='Cordia New', bold=True, color="FFFFFF", size=14)
            fill_green = PatternFill(start_color="2E7D32", end_color="2E7D32", fill_type="solid")
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            # --- Header ---
            ws.merge_cells('A1:G1')
            ws['A1'] = "ใบส่งสินค้าชั่วคราว"; ws['A1'].font = f_title; ws['A1'].alignment = Alignment(horizontal='center')
            ws['A2'] = "บริษัท โมบาย โลจิสติกส์ จำกัด"; ws['A2'].font = f_bold
            ws['A3'] = "278 หมู่ที่ 9 ตำบลบางโฉลง อ.บางพลี จ.สมุทรปราการ 10540"; ws['A3'].font = f_norm
            ws['A4'] = "โทร. 02-337-1200 แฟกซ์. 02-337-1201"; ws['A4'].font = f_norm
            
            ws['G2'] = f"Date: {current_date}"; ws['G2'].alignment = Alignment(horizontal='right'); ws['G2'].font = f_bold
            ws['G3'] = "Zone: "; ws['G3'].alignment = Alignment(horizontal='right'); ws['G3'].font = f_bold
            ws['G4'] = f"Delivery Date: {current_date}"; ws['G4'].alignment = Alignment(horizontal='right'); ws['G4'].font = f_bold
            
            ws['A6'] = f"Customer Name: {customer_name}"; ws['A6'].font = f_bold
            ws['A7'] = f"Store Code: {store_code}"; ws['A7'].font = f_bold
            ws['C7'] = f"Store Name: {branch_name}"; ws['C7'].font = f_bold

            # --- Table Header ---
            ws.merge_cells('E9:G9')
            ws['E9'] = "Qty"; ws['E9'].font = f_white; ws['E9'].fill = fill_green; ws['E9'].border = border
            ws['E9'].alignment = Alignment(horizontal='center', vertical='center')
            
            headers = ['No', 'Product Code', 'Product Name', 'Unit', 'ORDER', 'MBL', 'BNN']
            for i, h in enumerate(headers, 1):
                cell = ws.cell(row=10, column=i, value=h)
                cell.font, cell.fill, cell.border = f_white, fill_green, border
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                if i <= 4:
                    ws.merge_cells(start_row=9, start_column=i, end_row=10, end_column=i)
                    cell_main = ws.cell(row=9, column=i, value=h)
                    cell_main.font, cell_main.fill, cell_main.border = f_white, fill_green, border
                    cell_main.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            # --- Data Rows ---
            for row in ws.iter_rows(min_row=11, max_row=ws.max_row, min_col=1, max_col=7):
                for cell in row:
                    cell.border = border; cell.font = f_norm
                    cell.alignment = Alignment(vertical='center', wrap_text=True)
                    if cell.column in [1, 4, 5, 6, 7]:
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

            # --- Print Settings (Center & Fit) ---
            ws.page_setup.paperSize = 9 
            ws.page_setup.orientation = 'portrait'
            ws.page_setup.fitToWidth = 1 
            ws.page_setup.fitToHeight = 0
            ws.print_options.horizontalCentered = True
            
            # Margins & Widths
            ws.page_margins.left = 0.25; ws.page_margins.right = 0.25
            widths = {'A': 5, 'B': 14, 'C': 35, 'D': 10, 'E': 8, 'F': 8, 'G': 8}
            for col, w in widths.items(): ws.column_dimensions[col].width = w

    return output.getvalue(), final_list

# --- UI Layout ---
st.title("🚚 Delivery Formatter")
st.subheader("จัดรูปแบบใบส่งสินค้า A4 (Cordia New)")

uploaded_file = st.file_uploader("เลือกไฟล์ Excel (TestPOD)", type="xlsx")

if uploaded_file:
    with st.spinner('กำลังประมวลผลข้อมูล...'):
        try:
            excel_bytes, preview_data = process_excel_to_buffer(uploaded_file)
            
            # ส่วน Preview ข้อมูล
            st.success("✅ ประมวลผลสำเร็จ!")
            
            with st.expander("🔍 ดูตัวอย่างข้อมูลที่ประมวลผล (Preview)", expanded=True):
                # แสดงแค่บางคอลัมน์ให้ดูง่าย
                st.dataframe(preview_data[['Branch', 'Item No.', 'Description', 'Qty']], 
                             use_container_width=True, hide_index=True)
            
            # ปุ่มดาวน์โหลด
            col1, col2 = st.columns([1, 4])
            with col1:
                st.download_button(
                    label="📥 ดาวน์โหลดไฟล์ Excel",
                    data=excel_bytes,
                    file_name=f"Delivery_{datetime.now().strftime('%d%m%Y_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with col2:
                st.info("💡 ไฟล์ที่ดาวน์โหลดจะถูกตั้งค่า Font เป็น Cordia New และจัดกึ่งกลางหน้ากระดาษ A4 พร้อมสั่งพิมพ์ได้ทันที")

        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {str(e)}")
            st.info("โปรดตรวจสอบว่าไฟล์ Excel มีคอลัมน์ 'Item No.' และรูปแบบถูกต้องตามต้นฉบับ")