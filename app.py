import streamlit as st
import pandas as pd
import io
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, PatternFill

def parse_number(val):
    """Fungsi tangguh untuk mengubah berbagai jenis string format Indonesia menjadi angka murni (Float)"""
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    
    s = str(val).strip()
    if not s or s == '-':
        return 0.0
    
    # Bersihkan dari karakter aneh, spasi, atau simbol mata uang (sisakan angka, koma, titik, minus)
    s = re.sub(r'[^\d.,-]', '', s)
    
    # Logika Format Indonesia (Titik = Ribuan, Koma = Desimal)
    if ',' in s:
        s = s.replace('.', '')   # Hapus titik pemisah ribuan
        s = s.replace(',', '.')  # Ubah koma desimal menjadi titik agar bisa dihitung Python
    else:
        # Jika tidak ada koma tapi ada titik pemisah ribuan (misal: "1.500")
        if s.count('.') > 1 or re.search(r'\.\d{3}$', s):
            s = s.replace('.', '')
            
    try:
        return float(s)
    except ValueError:
        return 0.0

def clean_single_file(file):
    """Fungsi untuk membersihkan 1 file mentah"""
    # 1. Hapus 4 baris atas
    df = pd.read_excel(file, skiprows=4)
    
    # 2. Hapus 4 baris paling bawah
    df = df.iloc[:-4]
    
    # 3. Hapus baris yang semua kolomnya kosong
    df = df.dropna(how='all')
    
    # 4. Filter Kode Item hanya TWP/MWP/MBB/TBB/TBM
    df['Kode Item'] = df['Kode Item'].astype(str)
    keyword_pattern = 'TWP|MWP|MBB|TBB|TBM'
    df = df[df['Kode Item'].str.contains(keyword_pattern, case=True, na=False)]
    
    # 5. Forward fill kolom tertentu
    kolom_ffill = ['Gudang', 'Kode Barang Jadi', 'Nama Barang Jadi']
    for col in kolom_ffill:
        if col in df.columns:
            df[col] = df[col].ffill()
    
    # 6. Paksa konversi tipe data ke angka
    kolom_angka = ['Target Qty', 'Qty', 'Total', 'Total Harga', 'Harga', 'Berat Netto'] 
    for col in kolom_angka:
        if col in df.columns:
            df[col] = df[col].apply(parse_number)
            
    return df

def generate_excel_report(df_gabungan):
    """Fungsi untuk memecah sheet, membuat rekap detail, dan sheet summary"""
    summary_data = []
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active) # Hapus sheet default

    # ==========================================
    # FASE SPLIT SHEET & REKAP DETAIL
    # ==========================================
    grouped_df = df_gabungan.groupby('Nama Barang Jadi')

    for nama_bj, group_data in grouped_df:
        # 1. Penamaan Sheet: Tetap logika awal (Ambil penggalan setelah "-")
        nama_bj_str = str(nama_bj)
        if "-" in nama_bj_str:
            sheet_name = nama_bj_str.split("-", 1)[1].strip()
        else:
            sheet_name = nama_bj_str.strip()
            
        sheet_name = re.sub(r'[\\/*?:\[\]]', '_', sheet_name) # Bersihkan karakter terlarang
        sheet_name = sheet_name[:31] # Potong maksimal 31 karakter
        
        ws = wb.create_sheet(title=sheet_name)
        
        # Tulis tabel master ke sheet
        for r in dataframe_to_rows(group_data, index=False, header=True):
            ws.append(r)
            
        # Bold header master
        for cell in ws[1]:
            cell.font = Font(bold=True)
            
        # Tentukan posisi Tabel Rekap Detail
        max_col_master = ws.max_column
        start_col_rekap = max_col_master + 3
        
        # Header Tabel Rekap Detail
        rekap_headers = [
            "Kode Barang Jadi", "Nama Barang Jadi", "Target Qty", "Target Unit",
            "Kode Item", "Nama Item", "Unit", "Qty", "Total Harga"
        ]
        for i, header in enumerate(rekap_headers):
            cell = ws.cell(row=1, column=start_col_rekap + i, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

        kolom_total = 'Total Harga' if 'Total Harga' in group_data.columns else 'Total'
        
        # Isi Tabel Rekap Detail
        current_row = 2
        sub_grouped = group_data.groupby('Kode Barang Jadi')
        
        for kode_bj, sub_df in sub_grouped:
            nama_bj_val = sub_df['Nama Barang Jadi'].iloc[0]
            target_qty_sum = sub_df['Target Qty'].sum()
            target_unit_val = sub_df['Target Unit'].iloc[0]
            
            item_grouped = sub_df.groupby(['Kode Item', 'Nama Item', 'Unit']).agg({
                'Qty': 'sum',
                kolom_total: 'sum' 
            }).reset_index()
            
            jumlah_item = len(item_grouped)
            start_merge_row = current_row
            
            for _, item_row in item_grouped.iterrows():
                ws.cell(row=current_row, column=start_col_rekap, value=kode_bj)
                ws.cell(row=current_row, column=start_col_rekap + 1, value=nama_bj_val)
                ws.cell(row=current_row, column=start_col_rekap + 2, value=target_qty_sum)
                ws.cell(row=current_row, column=start_col_rekap + 3, value=target_unit_val)
                
                ws.cell(row=current_row, column=start_col_rekap + 4, value=item_row['Kode Item'])
                ws.cell(row=current_row, column=start_col_rekap + 5, value=item_row['Nama Item'])
                ws.cell(row=current_row, column=start_col_rekap + 6, value=item_row['Unit'])
                ws.cell(row=current_row, column=start_col_rekap + 7, value=item_row['Qty'])
                ws.cell(row=current_row, column=start_col_rekap + 8, value=item_row[kolom_total])
                
                current_row += 1
                
            if jumlah_item > 1:
                end_merge_row = current_row - 1
                for col_offset in range(4):
                    ws.merge_cells(start_row=start_merge_row, start_column=start_col_rekap + col_offset,
                                   end_row=end_merge_row, end_column=start_col_rekap + col_offset)
                    ws.cell(row=start_merge_row, column=start_col_rekap + col_offset).alignment = Alignment(vertical='center', horizontal='center')

        # ==========================================
        # KUMPULKAN DATA UNTUK SUMMARY
        # ==========================================
        # Mengelompokkan MURNI berdasarkan 'Nama Barang Jadi' saja agar datanya unik
        sum_group = group_data.groupby(['Gudang', 'Nama Barang Jadi', 'Target Unit']).agg({'Target Qty': 'sum'}).reset_index()
        for _, row in sum_group.iterrows():
            nama_asli = str(row['Nama Barang Jadi'])
            
            # REVISI: Logika penggalan untuk kolom MERK
            if "B. " in nama_asli:
                merk_display = nama_asli.split("B. ", 1)[1].strip()
            elif "-" in nama_asli:
                merk_display = nama_asli.split("-", 1)[1].strip()
            else:
                merk_display = nama_asli.strip()

            # Menyimpan data termasuk nama "sheet_name" aslinya secara rahasia
            summary_data.append({
                'DIVISI': row['Gudang'],
                'MERK': merk_display,
                'SATUAN TARGET': row['Target Unit'],
                'QTY TARGET': row['Target Qty'],
                'SATUAN': row['Target Unit'],
                'QTY GRAM': row['Target Qty'],
                'Sheet_Link': sheet_name # Variabel belakang layar untuk pengikat Hyperlink
            })

    # ==========================================
    # PEMBUATAN SHEET SUMMARY
    # ==========================================
    ws_summary = wb.create_sheet(title="Summary", index=0)
    
    headers = ['DIVISI', 'MERK', 'SATUAN TARGET', 'QTY TARGET', 'SATUAN', 'QTY GRAM']
    
    # Tulis header Summary
    for c_idx, col_name in enumerate(headers, 1):
        cell = ws_summary.cell(row=1, column=c_idx, value=col_name)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        
    # Tulis baris data Summary beserta eksekusi Hyperlink
    for r_idx, row in enumerate(summary_data, 2):
        ws_summary.cell(row=r_idx, column=1, value=row['DIVISI'])
        
        # Tulis MERK dan sisipkan fungsi Hyperlink internal Excel
        cell_merk = ws_summary.cell(row=r_idx, column=2, value=row['MERK'])
        cell_merk.hyperlink = f"#'{row['Sheet_Link']}'!A1" # Mengikat ke Sheet yang tepat sel A1
        cell_merk.font = Font(color="0563C1", underline="single") # Format biru bergaris bawah (Standar Hyperlink)
        
        ws_summary.cell(row=r_idx, column=3, value=row['SATUAN TARGET'])
        ws_summary.cell(row=r_idx, column=4, value=row['QTY TARGET'])
        ws_summary.cell(row=r_idx, column=5, value=row['SATUAN'])
        ws_summary.cell(row=r_idx, column=6, value=row['QTY GRAM'])

    # ==========================================
    # PEMBUATAN SHEET MASTER DATA
    # ==========================================
    ws_master = wb.create_sheet(title="master data", index=1)
    
    for c_idx, col_name in enumerate(df_gabungan.columns, 1):
        cell = ws_master.cell(row=1, column=c_idx, value=col_name)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        
    for r_idx, row in enumerate(dataframe_to_rows(df_gabungan, index=False, header=False), 2):
        for c_idx, value in enumerate(row, 1):
            ws_master.cell(row=r_idx, column=c_idx, value=value)

    wb.save(output)
    return output.getvalue()


# ==========================================
# UI STREAMLIT
# ==========================================
st.set_page_config(page_title="Pengelolaan Data Laporan", layout="wide")
st.title("Aplikasi Pengelolaan Data Laporan")

uploaded_files = st.file_uploader("Upload File Excel Mentah Anda (Bisa lebih dari 1)", type=['xlsx'], accept_multiple_files=True)

if uploaded_files:
    st.info(f"Memproses {len(uploaded_files)} file... Mohon tunggu.")
    
    try:
        list_df_bersih = []
        for file in uploaded_files:
            df_bersih = clean_single_file(file)
            list_df_bersih.append(df_bersih)
            
        df_gabungan = pd.concat(list_df_bersih, ignore_index=True)
        processed_excel = generate_excel_report(df_gabungan)
        
        st.success("Seluruh data berhasil digabungkan, diolah, dan laporan siap diunduh!")
        
        st.download_button(
            label="📥 Download Laporan Akhir (Excel)",
            data=processed_excel,
            file_name="Laporan_Gabungan_Rekap_Detail_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
        st.warning("Pastikan format kolom di semua file seragam dan sesuai dengan sistem.")
