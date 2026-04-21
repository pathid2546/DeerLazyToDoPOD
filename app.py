import streamlit as st
import pandas as pd
import io
from fpdf import FPDF
from datetime import datetime

# --- PDF Class ---
class DeliveryPDF(FPDF):
    def __init__(self, footer_img_stream=None):
        super().__init__()
        self.footer_img_stream = footer_img_stream
        # พยายามโหลด Font (ต้องมีไฟล์ ttf ในเครื่อง/Github)
        try:
            self.add_font('THSarabun', '', 'THSarabunNew.ttf', uni=True)
            self.add_font('THSarabun', 'B', 'THSarabunNew_Bold.ttf', uni=True)
            self.font_name = 'THSarabun'
        except:
            self.font_name = 'Arial' # สำรองกรณีไม่มี Font

    def header_info(self, customer_name, store_code, store_name, current_date):
        self.set_font(self.font_name, 'B', 20)
        self.cell(0, 10, 'ใบส่งสินค้าชั่วคราว', ln=True, align='C')
        
        self.set_font(self.font_name, 'B', 14)
        self.cell(120, 7, 'บริษัท โมบาย โลจิสติกส์ จำกัด')
        self.set_x(-60)
        self.cell(0, 7, f'Date: {current_date}', align='R', ln=True)
        
        self.set_font(self.font_name, '', 12)
        self.cell(120, 5, '278 หมู่ที่ 9 ตำบลบางโฉลง อ.บางพลี จ.สมุทรปราการ 10540')
        self.set_x(-60)
        self.cell(0, 5, 'Zone:', align='R', ln=True)
        self.cell(120, 5, 'โทร. 02-337-1200 แฟกซ์. 02-337-1201')
        self.set_x(-60)
        self.cell(0, 5, f'Delivery Date: {current_date}', align='R', ln=True)
        
        self.ln(5)
        self.set_font(self.font_name, 'B', 14)
        self.cell(0, 7, f'Customer Name {customer_name}', ln=True)
        self.cell(60, 7, f'Store Code: {store_code}')
        self.cell(0, 7, f'Store Name: {store_name}', ln=True)
        self.ln(2)

    def draw_table(self, df, customer_name, store_code, store_name, current_date):
        cols = [10, 35, 75, 25, 15, 15, 15]
        headers = ['No', 'Product Code', 'Product Name', 'Unit', 'ORD', 'MBL', 'BNN']
        
        # Table Header
        self.set_fill_color(46, 125, 50)
        self.set_text_color(255, 255, 255)
        self.set_font(self.font_name, 'B', 12)
        
        start_x = self.get_x()
        for i in range(4): self.cell(cols[i], 12, headers[i], border=1, align='C', fill=True)
        self.cell(sum(cols[4:]), 6, 'Qty', border=1, align='C', fill=True, ln=True)
        self.set_x(start_x + sum(cols[:4]))
        for i in range(4, 7): self.cell(cols[i], 6, headers[i], border=1, align='C', fill=True)
        self.ln(6)
        
        self.set_text_color(0, 0, 0)
        self.set_font(self.font_name, '', 12)
        for _, row in df.iterrows():
            if self.get_y() > 220: # ถ้าใกล้หมดหน้า
                self.add_page()
                self.header_info(customer_name, store_code, store_name, current_date)
                # วาดหัวตารางใหม่ในหน้าถัดไป... (ย่อเพื่อความสั้น)

            self.cell(cols[0], 7, str(row['No']), border=1, align='C')
            self.cell(cols[1], 7, str(row['Code']), border=1)
            self.cell(cols[2], 7, str(row['Name']), border=1)
            self.cell(cols[3], 7, str(row['Unit']), border=1, align='C')
            self.cell(cols[4], 7, str(row['ORD']), border=1, align='C')
            self.cell(cols[5], 7, '', border=1)
            self.cell(cols[6], 7, '', border=1, ln=True)

    def footer(self):
        if self.footer_img_stream:
            # ใช้รูปภาพจากที่อัปโหลด
            self.image(self.footer_img_stream, x=10, y=245, w=190)

# --- Processing Function ---
def process_to_pdf(uploaded_file, footer_img_file):
    raw_df = pd.read_excel(uploaded_file, header=None)
    header_row_index = next(i for i, row in raw_df.iterrows() if 'Item No.' in row.values)
    
    # ดึงข้อมูลรหัสร้านค้า
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
    final_list = df_melted.dropna(subset=['Item No.', 'Qty'])
    final_list = final_list[final_list['Qty'] > 0]

    # สร้าง PDF
    pdf = DeliveryPDF(footer_img_stream=footer_img_file)
    current_date = datetime.now().strftime('%d/%m/%Y')

    for branch_name, branch_data in final_list.groupby('Branch', sort=False):
        store_code = branch_to_code.get(str(branch_name).strip(), "")
        pdf.add_page()
        pdf.header_info("", store_code, branch_name, current_date)
        
        items_df = pd.DataFrame({
            'No': range(1, len(branch_data) + 1),
            'Code': branch_data['Item No.'],
            'Name': branch_data['Description'],
            'Unit': branch_data['UNIT'],
            'ORD': branch_data['Qty']
        })
        
        # ส่งตัวแปรให้ครบเพื่อแก้ Error
        pdf.draw_table(items_df, "", store_code, branch_name, current_date)

    return pdf.output(dest='S').encode('latin1')

# --- Streamlit UI ---
st.title("🚚 Delivery PDF Generator")

excel_input = st.file_uploader("1. อัปโหลดไฟล์ Excel", type="xlsx")
# ให้ผู้ใช้อัปโหลดรูปเอง จะได้ไม่เกิด Error "หาไฟล์ไม่เจอ"
img_input = st.file_uploader("2. อัปโหลดรูปลายเซ็น (Footer)", type=["png", "jpg"])

if excel_input and img_input:
    if st.button("สร้าง PDF"):
        try:
            # อ่านไบนารีของรูป
            img_byte_arr = io.BytesIO(img_input.read())
            pdf_out = process_to_pdf(excel_input, img_byte_arr)
            
            st.download_button(
                label="📥 ดาวน์โหลด PDF",
                data=pdf_out,
                file_name="delivery_note.pdf",
                mime="application/pdf"
            )
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
            st.warning("อย่าลืมอัปโหลดไฟล์ THSarabunNew.ttf ขึ้น Github ด้วยนะครับ!")