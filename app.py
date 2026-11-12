import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, PatternFill

def process_excel(file):
    # ==========================================
    # 1. FASE DATA CLEANING
    # ==========================================
    # Hapus 4 baris atas
    df = pd.read_excel(file, skiprows=4)
    
    # Hapus 4 baris paling bawah
    df = df.iloc[:-4]
    
    # Hapus baris yang semua kolomnya kosong
    df = df.dropna(how='all')
    
    # Filter Kode Item hanya TWP/MWP/MBB/TBB/TBM
    df['Kode Item'] = df['Kode Item'].astype(str)
    keyword_pattern = 'TWP|MWP|MBB|TBB|TBM'
    df = df[df['Kode Item'].str.contains(keyword_pattern, case=True, na=False)]
    
    # Forward fill kolom tertentu
    kolom_ffill = ['Gudang', 'Kode Barang Jadi', 'Nama Barang Jadi']
    df[kolom_ffill] = df[kolom_ffill].ffill()

    # Siapkan list untuk menampung data Sheet Summary nanti
    summary_data = []

    # Siapkan file Excel virtual (openpyxl)
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active) # Hapus sheet default

    # ==========================================
    # 2. FASE SPLIT SHEET & REKAP DETAIL
    # ==========================================
    # Kelompokkan berdasarkan Nama Barang Jadi
    grouped_df = df.groupby('Nama Barang Jadi')

    for nama_bj, group_data in grouped_df:
        # Penamaan Sheet: Ambil penggalan setelah "-"
        nama_bj_str = str(nama_bj)
        if "-" in nama_bj_str:
            sheet_name = nama_bj_str.split("-", 1)[1].strip()
        else:
            sheet_name = nama_bj_str.strip()
            
        # Pastikan nama sheet tidak lebih dari 31 karakter (batas maksimal Excel)
        sheet_name = sheet_name[:31]
        ws = wb.create_sheet(title=sheet_name)
        
        # Tulis tabel master ke sheet (mulai dari ujung kiri)
        for r in dataframe_to_rows(group_data, index=False, header=True):
            ws.append(r)
            
        # Percantik header master
        for cell in ws[1]:
            cell.font = Font(bold=True)
            
        # Tentukan posisi Tabel Rekap Detail (2 kolom kosong setelah tabel master)
        max_col_master = ws.max_column
        start_col_rekap = max_col_master + 3
        
        # Tulis Header Tabel Rekap Detail
        rekap_headers = [
            "Kode Barang Jadi", "Nama Barang Jadi", "Target Qty", "Target Unit",
            "Kode Item", "Nama Item", "Unit", "Qty", "Total Harga"
        ]
        for i, header in enumerate(rekap_headers):
            cell = ws.cell(row=1, column=start_col_rekap + i, value=header)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

        # Isi Tabel Rekap Detail
        current_row = 2
        # Group lagi berdasarkan Kode Barang Jadi di dalam sheet ini
        sub_grouped = group_data.groupby('Kode Barang Jadi')
        
        for kode_bj, sub_df in sub_grouped:
            # Nilai Level 1 (Barang Jadi)
            nama_bj_val = sub_df['Nama Barang Jadi'].iloc[0]
            target_qty_sum = sub_df['Target Qty'].sum()
            target_unit_val = sub_df['Target Unit'].iloc[0]
            
            # Nilai Level 2 (Item) -> Group berdasarkan Item unik
            item_grouped = sub_df.groupby(['Kode Item', 'Nama Item', 'Unit']).agg({
                'Qty': 'sum',
                'Total': 'sum' # Sesuaikan jika nama kolom aslinya bukan 'Total'
            }).reset_index()
            
            jumlah_item = len(item_grouped)
            start_merge_row = current_row
            
            for _, item_row in item_grouped.iterrows():
                # Tulis Level 1
                ws.cell(row=current_row, column=start_col_rekap, value=kode_bj)
                ws.cell(row=current_row, column=start_col_rekap + 1, value=nama_bj_val)
                ws.cell(row=current_row, column=start_col_rekap + 2, value=target_qty_sum)
                ws.cell(row=current_row, column=start_col_rekap + 3, value=target_unit_val)
                
                # Tulis Level 2
                ws.cell(row=current_row, column=start_col_rekap + 4, value=item_row['Kode Item'])
                ws.cell(row=current_row, column=start_col_rekap + 5, value=item_row['Nama Item'])
                ws.cell(row=current_row, column=start_col_rekap + 6, value=item_row['Unit'])
                ws.cell(row=current_row, column=start_col_rekap + 7, value=item_row['Qty'])
                ws.cell(row=current_row, column=start_col_rekap + 8, value=item_row['Total'])
                
                current_row += 1
                
            # Merge cell untuk Level 1 jika ada lebih dari 1 item di bawahnya
            if jumlah_item > 1:
                end_merge_row = current_row - 1
                for col_offset in range(4): # Menggabungkan 4 kolom pertama
                    ws.merge_cells(start_row=start_merge_row, start_column=start_col_rekap + col_offset,
                                   end_row=end_merge_row, end_column=start_col_rekap + col_offset)
                    # Posisikan teks di tengah setelah di-merge
                    ws.cell(row=start_merge_row, column=start_col_rekap + col_offset).alignment = Alignment(vertical='center', horizontal='center')

        # ==========================================
        # 3. KUMPULKAN DATA UNTUK SUMMARY
        # ==========================================
        # Tarik data dari master per Barang Jadi
        sum_group = group_data.groupby(['Gudang', 'Kode Barang Jadi', 'Nama Barang Jadi', 'Target Unit']).agg({'Target Qty': 'sum'}).reset_index()
        for _, row in sum_group.iterrows():
            summary_data.append({
                'DIVISI': row['Gudang'],
                'MERK': f"{row['Kode Barang Jadi']} {row['Nama Barang Jadi']}",
                'SATUAN TARGET': row['Target Unit'],
                'QTY TARGET': row['Target Qty'],
                'SATUAN': row['Target Unit'],
                'QTY GRAM': row['Target Qty']
            })

    # ==========================================
    # 4. PEMBUATAN SHEET SUMMARY
    # ==========================================
    ws_summary = wb.create_sheet(title="Summary", index=0) # Posisi paling depan
    df_summary = pd.DataFrame(summary_data)
    
    # Tulis header
    for c_idx, col_name in enumerate(df_summary.columns, 1):
        cell = ws_summary.cell(row=1, column=c_idx, value=col_name)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        
    # Tulis baris data
    for r_idx, row in enumerate(dataframe_to_rows(df_summary, index=False, header=False), 2):
        for c_idx, value in enumerate(row, 1):
            ws_summary.cell(row=r_idx, column=c_idx, value=value)

    # Simpan ke virtual bytes
    wb.save(output)
    return output.getvalue()

# ==========================================
# UI STREAMLIT
# ==========================================
st.set_page_config(page_title="Pengelolaan Data", layout="wide")
st.title("Aplikasi Pengelolaan Data Laporan")

uploaded_file = st.file_uploader("Upload File Excel Mentah Anda", type=['xlsx'])

if uploaded_file is not None:
    st.info("Sedang memproses data... Mohon tunggu.")
    
    try:
        # Jalankan fungsi utama
        processed_excel = process_excel(uploaded_file)
        
        st.success("Data berhasil diolah dan laporan siap diunduh!")
        
        # Tombol Download
        st.download_button(
            label="📥 Download Laporan Akhir (Excel)",
            data=processed_excel,
            file_name="Laporan_Rekap_Detail_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses data: {e}")
        st.warning("Pastikan nama kolom di file Excel persis sama dengan yang Anda sebutkan (seperti 'Target Qty', 'Target Unit', 'Total', dll).")
