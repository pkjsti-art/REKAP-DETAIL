import streamlit as st
import pandas as pd
import io
import re
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Alignment, Font, PatternFill

def parse_number(val):
    """Fungsi tangguh mengubah berbagai string format Indonesia menjadi angka murni (Float)"""
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    
    s = str(val).strip()
    if not s or s == '-':
        return 0.0
    
    # Bersihkan simbol (sisakan angka, koma, titik, minus)
    s = re.sub(r'[^\d.,-]', '', s)
    
    # Logika Titik/Koma
    if ',' in s:
        s = s.replace('.', '')
        s = s.replace(',', '.')
    else:
        if s.count('.') > 1 or re.search(r'\.\d{3}$', s):
            s = s.replace('.', '')
            
    try:
        return float(s)
    except ValueError:
        return 0.0

def clean_single_file(file):
    """Fungsi membersihkan 1 file mentah (Tahap 1 - Standar)"""
    df = pd.read_excel(file, skiprows=4)
    df = df.iloc[:-4]
    df = df.dropna(how='all')
    
    kolom_ffill = ['Gudang', 'Kode Barang Jadi', 'Nama Barang Jadi']
    for col in kolom_ffill:
        if col in df.columns:
            df[col] = df[col].ffill()
    
    kolom_angka = ['Target Qty', 'Qty', 'Total', 'Total Harga', 'Harga', 'Berat Netto'] 
    for col in kolom_angka:
        if col in df.columns:
            df[col] = df[col].apply(parse_number)
            
    return df

def generate_excel_report(df_gabungan):
    """Fungsi Utama Pembentukan Semua Aturan Logika Excel"""
    output = io.BytesIO()
    wb = Workbook()
    wb.remove(wb.active) # Hapus sheet default

    # ==========================================
    # 1. FILTERING DATA BERDASARKAN ATURAN GUDANG
    # ==========================================
    df_gabungan['Kode Item'] = df_gabungan['Kode Item'].astype(str)
    masks = []
    
    if 'Gudang' in df_gabungan.columns:
        for gudang, group in df_gabungan.groupby('Gudang', dropna=False):
            g = str(gudang).upper()
            if "DYEING" in g or "LAB" in g or "RND" in g:
                # Dyeing & Lab: Hanya TWP, MWP, TBM
                m = group['Kode Item'].str.contains('TWP|MWP|TBM', case=True, na=False)
            elif "500" in g or "5000" in g:
                # 500 & 5000 MRT: Hanya TWP
                m = group['Kode Item'].str.contains('TWP', case=True, na=False)
            else:
                # Softcone & Lainnya: Standar
                m = group['Kode Item'].str.contains('TWP|MWP|MBB|TBB|TBM', case=True, na=False)
            masks.append(group[m])
            
    df_filtered = pd.concat(masks, ignore_index=True) if masks else pd.DataFrame(columns=df_gabungan.columns)

    # ==========================================
    # 2. PENAMBAHAN KOLOM STANDAR (GR) & KONVERSI
    # ==========================================
    # Sisipkan Qty Target Standar (GR)
    if 'Target Unit' in df_filtered.columns and 'Target Qty' in df_filtered.columns:
        idx = df_filtered.columns.get_loc('Target Unit') + 1
        df_filtered.insert(idx, 'Qty Target Standar (GR)', df_filtered['Target Qty'])
    else:
        df_filtered['Qty Target Standar (GR)'] = df_filtered.get('Target Qty', 0)
        
    # Sisipkan Qty BB Standar (GR)
    if 'Qty' in df_filtered.columns and 'Unit' in df_filtered.columns:
        idx = df_filtered.columns.get_loc('Qty') + 1
        df_filtered.insert(idx, 'Qty BB Standar (GR)', df_filtered['Qty'])
    else:
        df_filtered['Qty BB Standar (GR)'] = df_filtered.get('Qty', 0)

    # Eksekusi Konversi (Jika KG dikali 1000) KHUSUS DYEING & LAB/RND
    mask_dyeing_lab = df_filtered['Gudang'].astype(str).str.upper().str.contains('DYEING|LAB|RND', regex=True)
    mask_kg_target = df_filtered['Target Unit'].astype(str).str.upper() == 'KG'
    mask_kg_unit = df_filtered['Unit'].astype(str).str.upper() == 'KG'

    df_filtered.loc[mask_dyeing_lab & mask_kg_target, 'Qty Target Standar (GR)'] = df_filtered['Target Qty'] * 1000
    df_filtered.loc[mask_dyeing_lab & mask_kg_unit, 'Qty BB Standar (GR)'] = df_filtered['Qty'] * 1000

    # Buat Sheet Summary Terlebih Dahulu (Agar di depan)
    ws_summary = wb.create_sheet(title="Summary", index=0)
    summary_dict = {}

    kolom_total = 'Total Harga' if 'Total Harga' in df_filtered.columns else 'Total'

    # ==========================================
    # 3. LOOPING PER GUDANG (MEMBUAT MASTER & SHEET ANAK)
    # ==========================================
    for gudang, gudang_df in df_filtered.groupby('Gudang', dropna=False):
        gudang_name = str(gudang)
        g_upper = gudang_name.upper()
        
        is_500_5000 = "500" in g_upper or "5000" in g_upper
        is_dyeing_lab = "DYEING" in g_upper or "LAB" in g_upper or "RND" in g_upper
        
        # --- A. BUAT SHEET MASTER PER GUDANG ---
        master_sheet_name = re.sub(r'[\\/*?:\[\]]', '_', f"Master - {gudang_name}")[:31]
        ws_master = wb.create_sheet(title=master_sheet_name)
        
        for c_idx, col_name in enumerate(gudang_df.columns, 1):
            cell = ws_master.cell(row=1, column=c_idx, value=col_name)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        for r_idx, row in enumerate(dataframe_to_rows(gudang_df, index=False, header=False), 2):
            for c_idx, value in enumerate(row, 1):
                ws_master.cell(row=r_idx, column=c_idx, value=value)
                
        summary_dict[gudang_name] = []
        
        # --- B. BUAT SHEET PEMECAHAN PER ITEM ---
        # Logika Grouping Sheet (Khusus 500/5000 pakai kombinasi Nama + Target Unit)
        if is_500_5000:
            grouped = gudang_df.groupby(['Nama Barang Jadi', 'Target Unit'], dropna=False)
        else:
            grouped = gudang_df.groupby(['Nama Barang Jadi'], dropna=False)
            
        for keys, group_data in grouped:
            if is_500_5000:
                nama_bj = str(keys[0])
                target_u = str(keys[1])
            else:
                nama_bj = str(keys)
                target_u = group_data['Target Unit'].iloc[0] if not group_data.empty and 'Target Unit' in group_data.columns else ""
                
            # Logika Penggalan Kata
            if "B. " in nama_bj:
                penggalan = nama_bj.split("B. ", 1)[1].strip()
            elif "-" in nama_bj:
                penggalan = nama_bj.split("-", 1)[1].strip()
            else:
                penggalan = nama_bj.strip()
                
            # Nama Tampilan (MERK) dan Nama Sheet
            if is_500_5000:
                sheet_name_raw = f"{penggalan} {target_u}"
                merk_display = f"{penggalan} {target_u}"
            else:
                sheet_name_raw = penggalan
                merk_display = penggalan
                
            # Mencegah duplikasi atau error nama sheet di Excel
            sheet_name_clean = re.sub(r'[\\/*?:\[\]]', '_', sheet_name_raw)[:31]
            base_sheet_name = sheet_name_clean
            counter = 1
            while base_sheet_name in wb.sheetnames:
                base_sheet_name = f"{sheet_name_clean[:28]}_{counter}"
                counter += 1
            final_sheet_name = base_sheet_name
            
            ws = wb.create_sheet(title=final_sheet_name)
            
            # Tulis Tabel Kiri (Data Mentah dari Sheet Ini)
            for r in dataframe_to_rows(group_data, index=False, header=True):
                ws.append(r)
            for cell in ws[1]:
                cell.font = Font(bold=True)
                
            # --- C. BUAT TABEL REKAP DETAIL (Di dalam Sheet) ---
            max_col_master = ws.max_column
            start_col_rekap = max_col_master + 3
            
            rekap_headers = [
                "Kode Barang Jadi", "Nama Barang Jadi", "Target Qty", "Target Unit",
                "Kode Item", "Nama Item", "Unit", "Qty", "Total Harga"
            ]
            for i, header in enumerate(rekap_headers):
                cell = ws.cell(row=1, column=start_col_rekap + i, value=header)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

            current_row = 2
            sub_grouped = group_data.groupby('Kode Barang Jadi', dropna=False)
            
            for kode_bj, sub_df in sub_grouped:
                nama_bj_val = sub_df['Nama Barang Jadi'].iloc[0]
                target_unit_val = sub_df['Target Unit'].iloc[0]
                
                # Logika Agregasi Rekap (Dyeing/Lab Menarik dari Kolom Standar)
                if is_dyeing_lab:
                    target_qty_sum = sub_df['Qty Target Standar (GR)'].sum()
                    item_grouped = sub_df.groupby(['Kode Item', 'Nama Item', 'Unit'], dropna=False).agg({
                        'Qty BB Standar (GR)': 'sum',
                        kolom_total: 'sum' 
                    }).reset_index()
                    item_grouped.rename(columns={'Qty BB Standar (GR)': 'Qty'}, inplace=True)
                else:
                    target_qty_sum = sub_df['Target Qty'].sum()
                    item_grouped = sub_df.groupby(['Kode Item', 'Nama Item', 'Unit'], dropna=False).agg({
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
                    for col_offset in range(4): # Merge 4 kolom awal (Level 1)
                        ws.merge_cells(start_row=start_merge_row, start_column=start_col_rekap + col_offset,
                                       end_row=end_merge_row, end_column=start_col_rekap + col_offset)
                        ws.cell(row=start_merge_row, column=start_col_rekap + col_offset).alignment = Alignment(vertical='center', horizontal='center')

            # --- D. TAMPUNG DATA UNTUK SUMMARY ---
            if is_dyeing_lab:
                t_qty_sum = group_data['Qty Target Standar (GR)'].sum()
                q_gram_sum = group_data['Qty Target Standar (GR)'].sum()
            else:
                t_qty_sum = group_data['Target Qty'].sum()
                q_gram_sum = group_data['Target Qty'].sum()

            satuan = target_u
            
            # Khusus 500/5000: Menjumlahkan nilai anak Qty dan mengambil anak Unit
            if is_500_5000:
                q_gram_sum = group_data['Qty'].sum()
                satuan = group_data['Unit'].iloc[0] if not group_data.empty else ""

            summary_dict[gudang_name].append({
                'DIVISI': gudang_name,
                'MERK': merk_display,
                'SATUAN TARGET': target_u,
                'QTY TARGET': t_qty_sum,
                'SATUAN': satuan,
                'QTY GRAM': q_gram_sum,
                'Sheet_Link': final_sheet_name # ID Hyperlink Rahasia
            })
            
    # ==========================================
    # 4. PENULISAN SHEET SUMMARY MENYAMPING
    # ==========================================
    summary_headers = ['DIVISI', 'MERK', 'SATUAN TARGET', 'QTY TARGET', 'SATUAN', 'QTY GRAM']
    current_start_col = 1
    
    for gudang, rows in summary_dict.items():
        # Title/Judul Tabel per Gudang
        ws_summary.cell(row=1, column=current_start_col, value=f"REKAP - {gudang}").font = Font(bold=True)
        
        # Tulis Header
        for i, h in enumerate(summary_headers):
            c = ws_summary.cell(row=2, column=current_start_col + i, value=h)
            c.font = Font(bold=True)
            c.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            
        # Tulis Baris Data
        for r_idx, row_dict in enumerate(rows, 3):
            ws_summary.cell(row=r_idx, column=current_start_col, value=row_dict['DIVISI'])
            
            # Tulis MERK & Pasang HYPERLINK Sakti
            c_merk = ws_summary.cell(row=r_idx, column=current_start_col + 1, value=row_dict['MERK'])
            c_merk.hyperlink = f"#'{row_dict['Sheet_Link']}'!A1"
            c_merk.font = Font(color="0563C1", underline="single")
            
            ws_summary.cell(row=r_idx, column=current_start_col + 2, value=row_dict['SATUAN TARGET'])
            ws_summary.cell(row=r_idx, column=current_start_col + 3, value=row_dict['QTY TARGET'])
            ws_summary.cell(row=r_idx, column=current_start_col + 4, value=row_dict['SATUAN'])
            ws_summary.cell(row=r_idx, column=current_start_col + 5, value=row_dict['QTY GRAM'])
            
        # Geser posisi kolom untuk Gudang selanjutnya (+6 kolom header + 2 kolom kosong)
        current_start_col += len(summary_headers) + 2 
        
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
        # Bersihkan & Gabung
        list_df_bersih = []
        for file in uploaded_files:
            df_bersih = clean_single_file(file)
            list_df_bersih.append(df_bersih)
            
        df_gabungan = pd.concat(list_df_bersih, ignore_index=True)
        
        # Eksekusi Pembuatan Excel
        processed_excel = generate_excel_report(df_gabungan)
        
        st.success("Semua logika berhasil dieksekusi! Laporan siap diunduh.")
        
        st.download_button(
            label="📥 Download Laporan Akhir (Excel)",
            data=processed_excel,
            file_name="Laporan_Otomatis_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Terjadi kesalahan teknis: {e}")
        st.warning("Pastikan nama kolom sesuai format dan file Excel tidak diubah strukturnya.")
