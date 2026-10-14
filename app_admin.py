import streamlit as st
import pandas as pd
import sqlite3
import io
import calendar
import requests
from datetime import datetime, timedelta, date
import plotly.express as px
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

try:
    from docx import Document
    from fpdf import FPDF
except ImportError:
    pass

# =====================================================================
# PENGATURAN & KONEKSI
# =====================================================================
PIN_ADMIN_RAHASIA = "123456" 
DB_PATH = "klinik_pipp.db"
TELEGRAM_BOT_TOKEN = "PASTE_TOKEN_DISINI"

def get_conn_write(): return sqlite3.connect(DB_PATH, check_same_thread=False)

def init_db():
    conn = get_conn_write()
    conn.execute('''CREATE TABLE IF NOT EXISTS pipp (id INTEGER PRIMARY KEY AUTOINCREMENT, waktu_input DATETIME, waktu_selesai DATETIME, no_bpjs TEXT, nama_pasien TEXT, alamat TEXT, no_telp TEXT, jenis_laporan TEXT, keluhan TEXT, solusi TEXT, link_bukti TEXT, is_real INTEGER, contact_id INTEGER, status TEXT DEFAULT 'Baru', priority TEXT DEFAULT 'Normal', short_link TEXT, consent_wa INTEGER DEFAULT 0)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS contacts (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, whatsapp TEXT UNIQUE, bpjs TEXT, nik TEXT, consent_wa INTEGER DEFAULT 0, first_interaction_at DATETIME, last_interaction_at DATETIME)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS main_categories (id INTEGER PRIMARY KEY AUTOINCREMENT, nama TEXT, icon TEXT, deskripsi TEXT, accent TEXT, light TEXT, tipe TEXT, urutan INTEGER)''')
    conn.commit()

def upgrade_db():
    conn = get_conn_write()
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(pipp)")
    columns = [info[1] for info in cursor.fetchall()]
    if 'petugas' not in columns: conn.execute("ALTER TABLE pipp ADD COLUMN petugas TEXT DEFAULT ''")
    if 'subkategori' not in columns: conn.execute("ALTER TABLE pipp ADD COLUMN subkategori TEXT DEFAULT ''")
    if 'is_audited' not in columns: conn.execute("ALTER TABLE pipp ADD COLUMN is_audited INTEGER DEFAULT 0")
    if 'jenis_pasien' not in columns: conn.execute("ALTER TABLE pipp ADD COLUMN jenis_pasien TEXT DEFAULT 'Umum'")
    if 'tensi' not in columns: conn.execute("ALTER TABLE pipp ADD COLUMN tensi TEXT DEFAULT ''")
    if 'suhu' not in columns: conn.execute("ALTER TABLE pipp ADD COLUMN suhu TEXT DEFAULT ''")
    conn.commit()

init_db()
upgrade_db()

# =====================================================================
# FUNGSI CEK DUPLIKASI GANDA
# =====================================================================
def cek_duplikasi(conn, nama, keluhan, tgl_str):
    cur = conn.cursor()
    cur.execute("SELECT id FROM pipp WHERE LOWER(nama_pasien) = ? AND LOWER(keluhan) = ? AND waktu_input LIKE ?", (str(nama).lower().strip(), str(keluhan).lower().strip(), f"{tgl_str}%"))
    return cur.fetchone() is not None

@st.cache_data(ttl=2) 
def get_prepared_data():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    df = pd.read_sql_query("SELECT * FROM pipp ORDER BY id DESC", conn).fillna('')
    conn.close()
    
    if df.empty: return df
    
    kata_kotor = ['-', '', 'nan', 'Belum Diklasifikasi', 'Belum Ditentukan', '[MENUNGGU TINDAKAN]', 'None']
    cols_to_clean = ['nama_pasien', 'no_bpjs', 'no_telp', 'alamat', 'keluhan', 'jenis_laporan', 'subkategori', 'solusi', 'petugas']
    for c in cols_to_clean:
        if c in df.columns: df[c] = df[c].replace(kata_kotor, '-')
            
    df['waktu_input_dt'] = pd.to_datetime(df['waktu_input'], errors='coerce')
    df['date_only'] = df['waktu_input_dt'].dt.date
    
    batas_24_jam = datetime.now() - timedelta(days=1)
    df['Smart_Status'] = df['status']
    is_baru = df['status'] == 'Baru'
    is_fresh = df['waktu_input_dt'] >= batas_24_jam
    df.loc[is_baru & is_fresh, 'Smart_Status'] = 'Baru Masuk'
    df.loc[is_baru & ~is_fresh, 'Smart_Status'] = 'Overdue'
    
    return df

# =====================================================================
# FUNGSI KALENDER HTML MURNI
# =====================================================================
def generate_html_calendar(year, month, data_dict):
    cal = calendar.monthcalendar(year, month)
    html = "<table style='width:100%; table-layout:fixed; border-collapse:collapse; text-align:center; font-family:Inter, sans-serif; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); margin-bottom: 20px;'>"
    html += "<tr style='background-color:#1e293b; color:white;'><th style='padding:10px;'>Sen</th><th>Sel</th><th>Rab</th><th>Kam</th><th>Jum</th><th>Sab</th><th>Min</th></tr>"
    for week in cal:
        html += "<tr>"
        for day in week:
            if day == 0: html += "<td style='border:1px solid #e2e8f0; padding:15px; background-color:#f8fafc;'></td>"
            else:
                d = date(year, month, day)
                count = data_dict.get(d, 0)
                if count > 0: 
                    html += f"<td style='border:1px solid #e2e8f0; padding:10px; background-color:#e0f2fe; font-weight:bold;'><span style='font-size:16px; color:#0f172a;'>{day}</span><br><br><span style='background-color:#ef4444; color:white; padding:3px 6px; border-radius:12px; font-size:11px; display:inline-block; margin-top:5px;'>{count} Lap</span></td>"
                else: 
                    html += f"<td style='border:1px solid #e2e8f0; padding:10px; color:#94a3b8; font-size:16px;'>{day}</td>"
        html += "</tr>"
    html += "</table>"
    return html

# =====================================================================
# FUNGSI EXPORT 
# =====================================================================
def style_excel_sheet(worksheet, dataframe):
    header_fill = PatternFill(start_color="0f172a", end_color="0f172a", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    for col_num, column_title in enumerate(dataframe.columns, 1):
        cell = worksheet.cell(row=1, column=col_num)
        cell.fill = header_fill; cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center"); cell.border = thin_border
        col_letter = get_column_letter(col_num)
        col_len = max(dataframe[column_title].astype(str).map(len).max(), len(str(column_title)))
        worksheet.column_dimensions[col_letter].width = min(col_len + 2, 45)
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
        for cell in row: cell.border = thin_border; cell.alignment = Alignment(vertical="top", wrap_text=True)

def prepare_export_df(df, cols_to_export):
    mapping = {'id': 'ID Laporan', 'waktu_input': 'Tgl Laporan', 'nama_pasien': 'Nama Pasien', 'no_bpjs': 'No. BPJS/NIK', 'no_telp': 'WhatsApp', 'alamat': 'Alamat', 'jenis_pasien': 'Jenis Pasien', 'keluhan': 'Keluhan/Diagnosa', 'jenis_laporan': 'Kategori Utama', 'subkategori': 'Subkategori', 'tensi': 'Tensi Darah', 'suhu': 'Suhu Badan', 'solusi': 'Tindakan/Solusi Akhir', 'petugas': 'Petugas/Dokter', 'status': 'Status Akhir'}
    df_eks = df.rename(columns=mapping)
    valid_cols = [mapping.get(c, c) for c in cols_to_export]
    return df_eks[[c for c in valid_cols if c in df_eks.columns]]

def create_excel(df, title):
    output = io.BytesIO()
    export_df = df.copy().astype(str).replace('nan', '')
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name=title)
        style_excel_sheet(writer.sheets[title], export_df)
    return output.getvalue()

def create_word(df, title):
    doc = Document()
    doc.add_heading(f'Laporan Klinik: {title}', 1)
    export_df = df.copy().astype(str).replace('nan', '')
    table = doc.add_table(rows=1, cols=len(export_df.columns))
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    for i, col_name in enumerate(export_df.columns):
        hdr_cells[i].text = str(col_name); hdr_cells[i].paragraphs[0].runs[0].font.bold = True
    for index, row in export_df.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row): row_cells[i].text = str(val)
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

def create_pdf(df, title):
    export_df = df.copy().astype(str).replace('nan', '')
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page(); pdf.set_font("helvetica", "B", 14)
    pdf.cell(0, 10, f"Laporan Klinik - {title}", align="C", new_x="LMARGIN", new_y="NEXT"); pdf.ln(5)
    pdf.set_font("helvetica", "B", 8)
    col_widths = [270 / len(export_df.columns)] * len(export_df.columns)
    pdf.set_fill_color(15, 23, 42); pdf.set_text_color(255, 255, 255)
    for i, col in enumerate(export_df.columns):
        pdf.cell(col_widths[i], 8, str(col)[:20], border=1, fill=True, align='C')
    pdf.ln(); pdf.set_font("helvetica", "", 8); pdf.set_text_color(0, 0, 0)
    for index, row in export_df.iterrows():
        for i, val in enumerate(row): pdf.cell(col_widths[i], 8, str(val)[:25], border=1)
        pdf.ln()
    return bytes(pdf.output())

def create_template_excel():
    output = io.BytesIO()
    df_template = pd.DataFrame(columns=['Tanggal Laporan', 'Nama Pasien', 'No. Identitas (BPJS/NIK)', 'No. WhatsApp/HP', 'Keluhan Awal', 'Kategori', 'Tindakan / Solusi', 'Nama Petugas', 'Status'])
    df_template.loc[0] = ['2026-12-01 10:00:00', 'Contoh Nama', '000123456', '081234567', 'Contoh masalah...', 'Masalah', 'Sudah dibantu', 'Dr. Budi', 'Selesai']
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_template.astype(str).to_excel(writer, index=False, sheet_name='Template_Upload')
        style_excel_sheet(writer.sheets['Template_Upload'], df_template)
    return output.getvalue()

# =====================================================================
# PROGRAM UTAMA ADMIN
# =====================================================================
st.set_page_config(page_title="Dashboard Admin", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    .stApp { background-color: #f8fafc; font-family: 'Inter', sans-serif; }
    .custom-table { width: 100%; border-collapse: collapse; background-color: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    .custom-table th { background-color: #1e293b; color: white; padding: 10px; text-align: left; font-size: 13px; }
    .custom-table td { padding: 8px 10px; border-bottom: 1px solid #e2e8f0; font-size: 13px; color: #334155; }
    .custom-table tr:hover { background-color: #f1f5f9; }
    
    .badge-barumasuk { background-color: #dcfce7; color: #166534; padding: 4px 8px; border-radius: 12px; font-weight: bold; font-size: 11px; }
    .badge-overdue { background-color: #fee2e2; color: #991b1b; padding: 4px 8px; border-radius: 12px; font-weight: bold; font-size: 11px; }
    .badge-selesai { background-color: #eff6ff; color: #1e40af; padding: 4px 8px; border-radius: 12px; font-weight: bold; font-size: 11px; }
    .badge-proses { background-color: #fef9c3; color: #854d0e; padding: 4px 8px; border-radius: 12px; font-weight: bold; font-size: 11px; }
    
    .form-container { background: white; padding: 20px; border-radius: 8px; border: 1px solid #e2e8f0; border-top: 4px solid #0284c7; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    .insight-box { background-color: #f0fdf4; border-left: 4px solid #16a34a; padding: 15px; margin-top: 15px; margin-bottom: 15px; border-radius: 4px; font-size: 13.5px; color: #334155; box-shadow: 0 1px 2px rgba(0,0,0,0.05);}
    .warning-box { background-color: #fef2f2; border-left: 4px solid #ef4444; padding: 15px; margin-top: 20px; border-radius: 4px; font-size: 12.5px; color: #7f1d1d;}
    
    div[data-testid="stFormSubmitButton"] > button { background: #0ea5e9 !important; color: white !important; font-weight: bold !important; border-radius: 6px !important; transition: all 0.2s !important; }
</style>
""", unsafe_allow_html=True)

if "admin_logged_in" not in st.session_state: st.session_state.admin_logged_in = False

if not st.session_state.admin_logged_in:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.info("🔐 Akses Manajemen Klinik")
        if st.button("Masuk"): 
            if st.text_input("PIN", type="password") == PIN_ADMIN_RAHASIA: st.session_state.admin_logged_in = True; st.rerun()
else:
    df_master = get_prepared_data()

    st.markdown("<h2 style='color:#0f172a; margin-top: -20px;'>🏥 Dashboard Klinik Pratama Papan Mas</h2>", unsafe_allow_html=True)
    
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "🚀 Kelola Pasien & Histori", 
        "📅 Kalender HTML Laporan", 
        "📊 Mega Analitik Historis", 
        "📑 Pusat Ekspor & Impor", 
        "✍️ Input Data Baru", 
        "👥 Interaksi & Kategori"
    ])

    # =========================================================
    # TAB 1: HISTORI LENGKAP & UPDATE
    # =========================================================
    with tab1:
        cari_pasien = st.text_input("🔍 Cari Pasien (Ketik Nama, No. BPJS, atau ID):", placeholder="Contoh: Sri rahayu...")
        if cari_pasien:
            sq = cari_pasien.lower()
            hasil = df_master[df_master['nama_pasien'].str.lower().str.contains(sq) | df_master['no_bpjs'].str.lower().str.contains(sq) | df_master['id'].astype(str).str.contains(sq)]
            
            if not hasil.empty:
                st.success(f"Ditemukan {len(hasil)} riwayat laporan untuk pencarian ini.")
                st.markdown(f"**Tabel Riwayat Seluruh Laporan: {hasil.iloc[0]['nama_pasien']}**")
                df_histori = hasil[['id', 'waktu_input', 'keluhan', 'jenis_laporan', 'petugas', 'Smart_Status', 'solusi']].rename(columns={'id':'ID', 'waktu_input':'Tgl Logbook', 'keluhan':'Keluhan', 'jenis_laporan':'Kategori', 'petugas':'Petugas', 'Smart_Status':'Status', 'solusi':'Solusi Diberikan'})
                st.dataframe(df_histori, use_container_width=True, hide_index=True)
                
                selected_id = st.selectbox("Pilih ID Laporan yang ingin Anda perbarui:", hasil['id'].tolist())
                p = hasil[hasil['id'] == selected_id].iloc[0]
                
                st.markdown("<div class='form-container'>", unsafe_allow_html=True)
                with st.form(f"form_update_{p['id']}"):
                    st.markdown(f"**Formulir Update Pasien (ID: {p['id']})**")
                    try: tgl_asli = pd.to_datetime(p['waktu_input']).date(); jam_asli = str(p['waktu_input'])[11:19] if len(str(p['waktu_input'])) > 10 else "00:00:00"
                    except: tgl_asli = datetime.now().date(); jam_asli = "00:00:00"
                    
                    c1, c2, c3, c4 = st.columns(4)
                    with c1: n_tgl = st.date_input("Tanggal Logbook", value=tgl_asli)
                    with c2: n_nama = st.text_input("Nama Pasien", value=p['nama_pasien'] if p['nama_pasien'] != '-' else '')
                    with c3: n_bpjs = st.text_input("No. BPJS/NIK", value=p['no_bpjs'] if p['no_bpjs'] != '-' else '')
                    with c4: n_wa = st.text_input("No. WhatsApp", value=p['no_telp'] if p['no_telp'] != '-' else '')
                    
                    c5, c6, c7, c8 = st.columns(4)
                    with c5: n_tensi = st.text_input("Tensi Darah", value=p.get('tensi', '') if p.get('tensi', '') != '-' else '')
                    with c6: n_suhu = st.text_input("Suhu", value=p.get('suhu', '') if p.get('suhu', '') != '-' else '')
                    with c7: n_jns = st.selectbox("Jenis Pasien", ["Umum", "BPJS", "Asuransi Lain", "VIP"], index=["Umum", "BPJS", "Asuransi Lain", "VIP"].index(p.get('jenis_pasien', 'Umum')) if p.get('jenis_pasien', 'Umum') in ["Umum", "BPJS", "Asuransi Lain", "VIP"] else 0)
                    with c8: n_petugas = st.text_input("Petugas", value=p['petugas'] if p['petugas'] != '-' else '')
                    
                    c9, c10 = st.columns([1, 1])
                    with c9: n_kat = st.text_input("Kategori Masalah", value=p['jenis_laporan'] if p['jenis_laporan'] != '-' else '')
                    with c10: 
                        stat_list = ["Baru", "Diproses", "Selesai", "Perlu Cek Ulang", "Ditutup"]
                        n_stat = st.selectbox("Status Sistem", stat_list, index=stat_list.index(p['status']) if p['status'] in stat_list else 0)
                    
                    n_solusi = st.text_area("Tindakan / Solusi", value=p['solusi'] if p['solusi'] != '-' else '', height=68)
                    
                    if st.form_submit_button("✔ Simpan Pembaruan Data"):
                        w_skrg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        w_input_baru = f"{n_tgl.strftime('%Y-%m-%d')} {jam_asli}"
                        w_sls = p['waktu_selesai']
                        if n_stat in ["Selesai", "Ditutup"] and w_sls == '': w_sls = w_skrg
                        conn = get_conn_write()
                        conn.execute("""UPDATE pipp SET waktu_input=?, nama_pasien=?, no_bpjs=?, no_telp=?, jenis_pasien=?, tensi=?, suhu=?, jenis_laporan=?, petugas=?, solusi=?, status=?, waktu_selesai=? WHERE id=?""", (w_input_baru, n_nama, n_bpjs, n_wa, n_jns, n_tensi, n_suhu, n_kat, n_petugas, n_solusi, n_stat, w_sls, p['id']))
                        conn.commit(); st.cache_data.clear(); st.rerun()
                st.markdown("</div><br>", unsafe_allow_html=True)
            else: st.warning("Nama pasien tidak ditemukan di database.")

        st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
        st.markdown("#### 📋 Validasi Cepat (Centang = Selesai)")
        
        df_audit = df_master.copy()
        df_audit['✔ Selesai'] = df_audit['status'].isin(['Selesai', 'Ditutup']) | (df_audit['is_audited'] == 1)
        df_edit_view = df_audit[['✔ Selesai', 'id', 'waktu_input', 'nama_pasien', 'keluhan', 'Smart_Status', 'petugas']].rename(columns={'waktu_input':'Tgl Logbook', 'nama_pasien':'Nama Pasien', 'Smart_Status':'Status'})
        
        edited_df = st.data_editor(df_edit_view, hide_index=True, use_container_width=True, disabled=['id', 'Tgl Logbook', 'Nama Pasien', 'keluhan', 'Status', 'petugas'])
        
        if st.button("💾 Simpan Validasi Centang", type="primary"):
            w_skrg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            conn = get_conn_write()
            for index, row in edited_df.iterrows():
                if row['✔ Selesai']: conn.execute("UPDATE pipp SET status = 'Selesai', is_audited = 1, waktu_selesai = ? WHERE id = ? AND status != 'Selesai'", (w_skrg, row['id']))
            conn.commit(); st.cache_data.clear(); st.success("Pembaruan tersimpan!"); st.rerun()

    # =========================================================
    # TAB 2: KALENDER (DENGAN QUICK FILTER TANGGAL LINTAS TAHUN)
    # =========================================================
    with tab2:
        st.markdown("<h4 style='color:#0f172a;'>📅 Kalender Pemantauan Laporan</h4>", unsafe_allow_html=True)
        
        df_cal = df_master.copy()
        
        cal_left, cal_right = st.columns([1, 1], gap="large")
        with cal_left:
            st.markdown("##### Visualisasi Kalender Harian")
            try: default_date = df_cal['date_only'].max() if not df_cal.empty else datetime.now().date()
            except: default_date = datetime.now().date()
            
            c_nav1, c_nav2 = st.columns([1, 3])
            with c_nav1: sel_tahun = st.selectbox("Pilih Tahun:", range(2020, datetime.now().year + 3), index=(default_date.year - 2020), key="cal_y")
            with c_nav2:
                bulan_list = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
                sel_bulan = st.selectbox("Pilih Bulan:", range(1, 13), format_func=lambda x: bulan_list[x-1], index=(default_date.month - 1), key="cal_m")
            
            df_bln_ini = df_cal[(df_cal['waktu_input_dt'].dt.year == sel_tahun) & (df_cal['waktu_input_dt'].dt.month == sel_bulan)]
            laporan_per_hari = df_bln_ini.groupby('date_only').size().to_dict()
            html_cal = generate_html_calendar(sel_tahun, sel_bulan, laporan_per_hari)
            st.markdown(html_cal, unsafe_allow_html=True)
            
        with cal_right:
            st.markdown("##### 🔍 Mesin Pencari Laporan (Rentang Waktu Lintas Tahun)")
            st.write("Gunakan pilihan cepat di bawah ini untuk mem-filter tabel, atau tentukan tanggal lintas tahun secara manual.")
            
            # QUICK FILTERS UNTUK MEMPERMUDAH RENTANG TAHUNAN
            c_qk1, c_qk2, c_qk3, c_qk4, c_qk5 = st.columns(5)
            today = datetime.now().date()
            if 'quick_date' not in st.session_state: st.session_state.quick_date = ()
            
            if c_qk1.button("Hari Ini", use_container_width=True): st.session_state.quick_date = (today, today)
            if c_qk2.button("7 Hari", use_container_width=True): st.session_state.quick_date = (today - timedelta(days=7), today)
            if c_qk3.button("Bulan Ini", use_container_width=True): st.session_state.quick_date = (today.replace(day=1), today)
            if c_qk4.button("Tahun Ini", use_container_width=True): st.session_state.quick_date = (today.replace(month=1, day=1), today)
            if c_qk5.button("Semua", use_container_width=True): st.session_state.quick_date = ()
            
            tgl_pilihan = st.date_input("Filter Rentang Tanggal Spesifik:", value=st.session_state.quick_date)
            
            f_col1, f_col2 = st.columns(2)
            with f_col1:
                f_stat_cal = st.multiselect("Filter Status:", options=[s for s in df_cal['status'].unique() if s != '-'])
                f_kat_cal = st.multiselect("Filter Kategori:", options=[k for k in df_cal['jenis_laporan'].unique() if k != '-'])
            with f_col2:
                f_pet_cal = st.multiselect("Filter Petugas:", options=[p for p in df_cal['petugas'].unique() if p != '-'])
            
            df_table_view = df_cal.copy()
            if len(tgl_pilihan) == 2: df_table_view = df_table_view[(df_table_view['date_only'] >= tgl_pilihan[0]) & (df_table_view['date_only'] <= tgl_pilihan[1])]
            elif len(tgl_pilihan) == 1: df_table_view = df_table_view[df_table_view['date_only'] == tgl_pilihan[0]]
            
            if f_stat_cal: df_table_view = df_table_view[df_table_view['status'].isin(f_stat_cal)]
            if f_kat_cal: df_table_view = df_table_view[df_table_view['jenis_laporan'].isin(f_kat_cal)]
            if f_pet_cal: df_table_view = df_table_view[df_table_view['petugas'].isin(f_pet_cal)]
            
            df_table_view = df_table_view[['id', 'waktu_input', 'nama_pasien', 'keluhan', 'Smart_Status', 'petugas']].rename(columns={'id':'ID', 'waktu_input':'Tgl Logbook', 'nama_pasien':'Pasien', 'keluhan':'Keluhan', 'Smart_Status':'Status', 'petugas':'Petugas'})
            
            st.markdown(f"**Menemukan {len(df_table_view)} Laporan Sesuai Kriteria:**")
            st.dataframe(df_table_view, use_container_width=True, hide_index=True)

    # =========================================================
    # TAB 3: MEGA ANALITIK (SESUAI WIREFRAME REVISI)
    # =========================================================
    with tab3:
        if "ana_drill_col" not in st.session_state: st.session_state.ana_drill_col = None
        if "ana_drill_val" not in st.session_state: st.session_state.ana_drill_val = None
        
        df_an = df_master.copy()
        kata_kosong = ['-', '', 'nan', 'Belum Diklasifikasi', 'Belum Ditentukan', '[MENUNGGU TINDAKAN]', 'None', 'KOSONG (Tanpa Petugas)', 'KOSONG (Belum Diklasifikasi)']
        
        cols_check = ['nama_pasien', 'no_bpjs', 'keluhan', 'jenis_laporan', 'solusi', 'petugas']
        for c in cols_check:
            if c in df_an.columns: df_an[c] = df_an[c].replace(kata_kosong, 'KOSONG')
        
        try: min_d = df_an['date_only'].min().strftime('%d %b %Y'); max_d = df_an['date_only'].max().strftime('%d %b %Y')
        except: min_d = "-"; max_d = "-"
        
        st.markdown("#### ⚙️ Filter Global & Status Data")
        st.markdown(f"**🗓️ Rentang Tanggal Historis Database:** {min_d} hingga {max_d}")
        
        col_f1, col_f2 = st.columns([2, 1])
        with col_f1:
            tgl_global = st.date_input("Date Input Range: Filter Global", value=(), key="ana_tgl_global")
        with col_f2:
            st.markdown("<br>", unsafe_allow_html=True)
            cf1, cf2 = st.columns(2)
            with cf1:
                if st.button("🔄 Reset Tanggal Global", use_container_width=True):
                    st.session_state.ana_drill_col = None; st.session_state.ana_drill_val = None; st.rerun()
            with cf2:
                if st.button("❌ Hapus Drill-Down", use_container_width=True):
                    st.session_state.ana_drill_col = None; st.session_state.ana_drill_val = None; st.rerun()

        df_filtered = df_an.copy()
        if len(tgl_global) == 2: df_filtered = df_filtered[(df_filtered['date_only'] >= tgl_global[0]) & (df_filtered['date_only'] <= tgl_global[1])]
        elif len(tgl_global) == 1: df_filtered = df_filtered[df_filtered['date_only'] == tgl_global[0]]
        
        total_db = len(df_an)
        total_filtered = len(df_filtered)
        is_drill_active = st.session_state.ana_drill_col is not None
        
        df_drilldown = df_filtered.copy()
        if is_drill_active:
            col_filter = st.session_state.ana_drill_col
            val_filter = st.session_state.ana_drill_val
            if col_filter == "Periode_Bulan_Str":
                df_drilldown['Periode_Bulan_Str'] = df_drilldown['waktu_input_dt'].dt.to_period('M').astype(str)
                df_drilldown = df_drilldown[df_drilldown['Periode_Bulan_Str'] == val_filter]
            elif col_filter == "nama_clean":
                df_drilldown['nama_clean'] = df_drilldown['nama_pasien'].str.lower().str.strip()
                df_drilldown = df_drilldown[df_drilldown['nama_clean'] == val_filter]
            else:
                df_drilldown = df_drilldown[df_drilldown[col_filter] == val_filter]
                
        total_drilldown = len(df_drilldown)

        st.info(f"**Indikator Data:**\n🗄️ Total Seluruh Database: **{total_db}** laporan\n📅 Total Setelah Filter Tanggal: **{total_filtered}** laporan\n🔍 Total Setelah Drill-Down: **{total_drilldown}** laporan")

        # =======================================================
        # BAGIAN A: Ringkasan Utama & Kualitas Data
        # =======================================================
        st.markdown("<hr style='margin: 20px 0;'>", unsafe_allow_html=True)
        st.markdown("#### Bagian A: Ringkasan Utama & Kualitas Data")
        
        jml_sol_kosong = len(df_filtered[df_filtered['solusi'] == 'KOSONG'])
        pct_sol_kosong = round((jml_sol_kosong / total_filtered * 100), 1) if total_filtered > 0 else 0
        jml_audit = len(df_filtered[df_filtered['is_audited'] == 1])
        pct_audit = round((jml_audit / total_filtered * 100), 1) if total_filtered > 0 else 0
        jml_id_terisi = len(df_filtered[df_filtered['no_bpjs'] != 'KOSONG'])
        pct_identitas = round((jml_id_terisi / total_filtered * 100), 1) if total_filtered > 0 else 0
        
        ka1, ka2, ka3, ka4 = st.columns(4)
        ka1.metric("Total Laporan (Terfilter)", f"{total_filtered} Data")
        ka2.metric("Dokumentasi Solusi Kosong", f"{jml_sol_kosong}", f"({pct_sol_kosong}%)", delta_color="inverse")
        ka3.metric("Progres Audit Historis", f"{pct_audit}%", f"{jml_audit} Divalidasi")
        ka4.metric("Kelengkapan Identitas", f"{pct_identitas}%", "Ada NIK/BPJS")
        
        st.markdown(f"<div class='insight-box'>💡 <b>INSIGHT KONDISI DATA:</b><br>Dari {total_filtered} data pada periode yang dipilih, sebanyak {jml_sol_kosong} laporan ({pct_sol_kosong}%) belum memiliki dokumentasi tindakan/solusi yang tercatat di sistem (Catatan: ini hanya menunjukkan ketiadaan teks dokumentasi, bukan bukti laporan tidak ditangani). Di sisi lain, kelengkapan pengisian NIK/BPJS mencapai {pct_identitas}%.</div>", unsafe_allow_html=True)

        # =======================================================
        # BAGIAN B: Tren dan Pola Waktu Pencatatan
        # =======================================================
        st.markdown("<hr style='margin: 20px 0;'>", unsafe_allow_html=True)
        st.markdown("#### Bagian B: Tren dan Pola Waktu Pencatatan")
        
        layout_cfg = dict(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(family="Inter", color="#334155"), margin=dict(t=20, b=10, l=10, r=10))
        cb1, cb2 = st.columns(2)
        
        bln_puncak = "-"
        jml_puncak_bln = 0
        hr_puncak = "-"
        
        with cb1:
            if total_filtered > 0:
                df_filtered['Periode_Bulan'] = df_filtered['waktu_input_dt'].dt.to_period('M')
                tren_bulan = df_filtered['Periode_Bulan'].value_counts().sort_index().reset_index()
                tren_bulan.columns = ['Bulan', 'Jumlah']
                tren_bulan['Bulan_Str'] = tren_bulan['Bulan'].astype(str)
                
                fig_bulan = px.line(tren_bulan, x='Bulan_Str', y='Jumlah', markers=True, color_discrete_sequence=['#ef4444'], title="Grafik Tren Pencatatan Bulanan")
                fig_bulan.update_layout(**layout_cfg)
                fig_bulan.update_xaxes(title_text="Bulan (Kronologis)")
                fig_bulan.update_yaxes(title_text="Jumlah Laporan Tercatat")
                st.plotly_chart(fig_bulan, use_container_width=True)
                
                if not tren_bulan.empty:
                    bln_puncak = tren_bulan.loc[tren_bulan['Jumlah'].idxmax()]['Bulan_Str']
                    jml_puncak_bln = tren_bulan['Jumlah'].max()
            else: st.info("Data bulanan tidak tersedia.")
            
        with cb2:
            if total_filtered > 0:
                df_filtered['Hari'] = df_filtered['waktu_input_dt'].dt.day_name().map({'Monday':'Senin', 'Tuesday':'Selasa', 'Wednesday':'Rabu', 'Thursday':'Kamis', 'Friday':'Jumat', 'Saturday':'Sabtu', 'Sunday':'Minggu'})
                hari_c = df_filtered['Hari'].value_counts().reindex(['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']).fillna(0).reset_index()
                hari_c.columns = ['Hari', 'Jumlah']
                
                fig_hari = px.bar(hari_c, x='Hari', y='Jumlah', color_discrete_sequence=['#8b5cf6'], title="Pola Hari Pencatatan")
                fig_hari.update_layout(**layout_cfg)
                fig_hari.update_xaxes(title_text="Senin -> Minggu")
                fig_hari.update_yaxes(title_text="Total Laporan")
                st.plotly_chart(fig_hari, use_container_width=True)
                
                if not hari_c.empty:
                    hr_puncak = hari_c.loc[hari_c['Jumlah'].idxmax()]['Hari']
            else: st.info("Data harian tidak tersedia.")
            
        st.markdown(f"<div class='insight-box'>💡 <b>INSIGHT POLA WAKTU:</b><br>Pencatatan data tertinggi berada pada bulan {bln_puncak} sebanyak {jml_puncak_bln} data. Untuk pola mingguan, {hr_puncak} adalah hari dengan volume pencatatan paling banyak. Penyebab lonjakan atau penurunan ini belum dapat disimpulkan murni dari data historis.</div>", unsafe_allow_html=True)
        if bln_puncak != "-" and st.button(f"🔍 Tampilkan Laporan di Bulan {bln_puncak} pada Bagian E", use_container_width=True):
            st.session_state.ana_drill_col = "Periode_Bulan_Str"
            st.session_state.ana_drill_val = bln_puncak
            st.rerun()

        # =======================================================
        # BAGIAN C: Pasien dan Pola Keluhan
        # =======================================================
        st.markdown("<hr style='margin: 20px 0;'>", unsafe_allow_html=True)
        st.markdown("#### Bagian C: Pasien dan Pola Keluhan")
        
        cc1, cc2 = st.columns(2)
        pas_terbanyak = "-"
        jml_pas_terbanyak = 0
        kel_terbanyak = "-"
        
        with cc1:
            st.markdown("**Top 10 Nama Paling Sering**")
            if total_filtered > 0:
                df_pas = df_filtered[df_filtered['nama_pasien'] != 'KOSONG'].copy()
                df_pas['nama_clean'] = df_pas['nama_pasien'].str.lower().str.strip()
                pas_c = df_pas['nama_clean'].value_counts().head(10).reset_index()
                pas_c.columns = ['Nama Tercatat', 'Jumlah Kemunculan']
                st.dataframe(pas_c, use_container_width=True, hide_index=True)
                if not pas_c.empty:
                    pas_terbanyak = pas_c.iloc[0]['Nama Tercatat']
                    jml_pas_terbanyak = pas_c.iloc[0]['Jumlah Kemunculan']
            else: st.info("Data pasien kosong.")

        with cc2:
            st.markdown("**Top 10 Teks Keluhan Mentah**")
            if total_filtered > 0:
                df_kel = df_filtered[df_filtered['keluhan'] != 'KOSONG']
                kel_c = df_kel['keluhan'].value_counts().head(10).reset_index()
                kel_c.columns = ['Teks Keluhan Asli', 'Jumlah Ditemukan']
                st.dataframe(kel_c, use_container_width=True, hide_index=True)
                if not kel_c.empty:
                    kel_terbanyak = kel_c.iloc[0]['Teks Keluhan Asli']
            else: st.info("Data keluhan kosong.")
            
        st.markdown(f"<div class='insight-box'>💡 <b>INSIGHT POLA PASIEN & KELUHAN:</b><br>Nama yang paling sering tercatat adalah '{pas_terbanyak}' dengan {jml_pas_terbanyak} kemunculan. Teks keluhan mentah yang paling banyak diketik adalah '{kel_terbanyak}'.<br>Keterbatasan: Sistem menggunakan pencocokan teks persis. Variasi ejaan (misal 'Sri Rahayu' vs 'sri rahayu') atau kesalahan ketik dihitung terpisah, sehingga tidak bisa diklaim mutlak sebagai individu yang sama tanpa validasi NIK.</div>", unsafe_allow_html=True)
        if pas_terbanyak != "-" and st.button(f"🔍 Tampilkan Riwayat Nama '{pas_terbanyak}' pada Bagian E", use_container_width=True):
            st.session_state.ana_drill_col = "nama_clean"
            st.session_state.ana_drill_val = pas_terbanyak
            st.rerun()

        # =======================================================
        # BAGIAN D: Kualitas & Distribusi Klasifikasi Data
        # =======================================================
        st.markdown("<hr style='margin: 20px 0;'>", unsafe_allow_html=True)
        st.markdown("#### Bagian D: Kualitas & Distribusi Klasifikasi Data")
        
        cd1, cd2 = st.columns(2)
        kat_terbanyak = "-"
        jml_kat_kosong = 0
        
        with cd1:
            if total_filtered > 0:
                kat_c = df_filtered['jenis_laporan'].value_counts().reset_index()
                kat_c.columns = ['Kategori', 'Total']
                kat_c['Warna'] = kat_c['Kategori'].apply(lambda x: '#cbd5e1' if x == 'KOSONG' else '#0ea5e9')
                
                fig_kat = px.bar(kat_c.head(10), x='Total', y='Kategori', orientation='h', color='Warna', color_discrete_map="identity", title="Distribusi Kategori Valid")
                fig_kat.update_layout(yaxis={'categoryorder':'total ascending'}, showlegend=False, **layout_cfg)
                st.plotly_chart(fig_kat, use_container_width=True)
                
                jml_kat_kosong = len(df_filtered[df_filtered['jenis_laporan'] == 'KOSONG'])
                kat_valid = kat_c[kat_c['Kategori'] != 'KOSONG']
                if not kat_valid.empty: kat_terbanyak = kat_valid.iloc[0]['Kategori']
            else: st.info("Data klasifikasi kosong.")

        with cd2:
            if total_filtered > 0:
                p_id = round((len(df_filtered[df_filtered['no_bpjs'] != 'KOSONG']) / total_filtered * 100), 1)
                p_kel = round((len(df_filtered[df_filtered['keluhan'] != 'KOSONG']) / total_filtered * 100), 1)
                p_kat = round((len(df_filtered[df_filtered['jenis_laporan'] != 'KOSONG']) / total_filtered * 100), 1)
                p_sol = round((len(df_filtered[df_filtered['solusi'] != 'KOSONG']) / total_filtered * 100), 1)
                
                df_bar = pd.DataFrame({
                    'Atribut': ['Identitas NIK', 'Teks Keluhan Asli', 'Kategori Klasifikasi', 'Dokumentasi Solusi'],
                    'Terisi (%)': [p_id, p_kel, p_kat, p_sol]
                })
                fig_bar_stack = px.bar(df_bar, x='Terisi (%)', y='Atribut', orientation='h', color_discrete_sequence=['#10b981'], title="Evaluasi Kelengkapan Dok. (Terisi vs Kosong)")
                fig_bar_stack.update_layout(xaxis=dict(range=[0, 100]), **layout_cfg)
                st.plotly_chart(fig_bar_stack, use_container_width=True)

        st.markdown(f"<div class='insight-box'>💡 <b>INSIGHT KELENGKAPAN:</b><br>Kategori klasifikasi yang paling banyak digunakan adalah '{kat_terbanyak}'. Namun, masih terdapat {jml_kat_kosong} data yang dilabeli sebagai 'Belum Diklasifikasi' (KOSONG). Admin diharapkan melengkapi kolom kategori dan solusi agar dokumentasi laporan historis semakin akurat.</div>", unsafe_allow_html=True)
        if st.button("🔍 Tampilkan Laporan yang Kategori/Solusinya Masih Kosong pada Bagian E", use_container_width=True):
            st.session_state.ana_drill_col = "solusi"
            st.session_state.ana_drill_val = "KOSONG"
            st.rerun()

        # =======================================================
        # BAGIAN E: Detail Laporan Tersinkronisasi
        # =======================================================
        st.markdown("<hr style='margin: 20px 0; border-top: 3px solid #cbd5e1;'>", unsafe_allow_html=True)
        st.markdown("#### Bagian E: Detail Laporan Tersinkronisasi")
        
        cari_teks = st.text_input("🔍 Kotak Pencarian Teks Bebas: Cari ID, Nama, NIK, atau Keluhan...")
        
        if cari_teks:
            sq = cari_teks.lower()
            df_drilldown = df_drilldown[df_drilldown['id'].astype(str).str.contains(sq) | df_drilldown['nama_pasien'].str.lower().str.contains(sq) | df_drilldown['no_bpjs'].str.lower().str.contains(sq) | df_drilldown['keluhan'].str.lower().str.contains(sq)]

        df_view_final = df_drilldown[['id', 'waktu_input', 'nama_pasien', 'no_bpjs', 'keluhan', 'jenis_laporan', 'solusi']].copy()
        df_view_final = df_view_final.rename(columns={'id':'ID', 'waktu_input':'Tgl Laporan', 'nama_pasien':'Nama Pasien', 'no_bpjs':'NIK/BPJS', 'keluhan':'Keluhan', 'jenis_laporan':'Kategori', 'solusi':'Solusi Terdata'})
        
        st.dataframe(df_view_final, use_container_width=True, hide_index=True)
        
        st.markdown("<div class='warning-box'>⚠️ <b>Peringatan Keterbatasan Historis:</b> Statistik di atas difokuskan pada evaluasi KUALITAS dan KELENGKAPAN PENCATATAN data di sistem. Karena alur operasional historis tidak direkam secara memadai, metrik ini tidak ditujukan untuk mengukur keberhasilan penyelesaian kasus, kinerja petugas medis, atau status 'terbengkalai'.</div>", unsafe_allow_html=True)

    # =========================================================
    # TAB 4: PUSAT EKSPOR & IMPOR DENGAN ANTI GANDA
    # =========================================================
    with tab4:
        st.markdown("<h4 style='color:#0f172a;'>📑 Pusat Ekspor & Impor Data Klinik</h4>", unsafe_allow_html=True)
        with st.container(border=True):
            st.markdown("##### 🔍 1. Ekspor Data Spesifik Terfilter")
            df_eks = df_master.copy()
            f1, f2, f3 = st.columns(3)
            with f1: fil_tgl = st.date_input("🗓️ Rentang Waktu:", value=())
            with f2: fil_stat = st.multiselect("🚦 Filter Status:", options=df_eks['status'].unique())
            with f3: fil_kat = st.multiselect("🩺 Filter Kategori:", options=[k for k in df_eks['jenis_laporan'].unique() if k != '-'])
            
            if len(fil_tgl) == 2: df_eks = df_eks[(df_eks['date_only'] >= fil_tgl[0]) & (df_eks['date_only'] <= fil_tgl[1])]
            elif len(fil_tgl) == 1: df_eks = df_eks[df_eks['date_only'] == fil_tgl[0]]
            if fil_stat: df_eks = df_eks[df_eks['status'].isin(fil_stat)]
            if fil_kat: df_eks = df_eks[df_eks['jenis_laporan'].isin(fil_kat)]
            
            kolom_pilih = st.multiselect("Kolom yang dicetak:", options=['id', 'waktu_input', 'nama_pasien', 'no_bpjs', 'no_telp', 'alamat', 'jenis_pasien', 'keluhan', 'tensi', 'suhu', 'jenis_laporan', 'subkategori', 'solusi', 'petugas', 'status'], default=['waktu_input', 'nama_pasien', 'no_bpjs', 'keluhan', 'solusi', 'status'])
            df_siap_ekspor = prepare_export_df(df_eks, kolom_pilih)
            
            dl1, dl2, dl3 = st.columns(3)
            nm_file = f"Laporan_{datetime.now().strftime('%Y%m%d')}"
            with dl1: st.download_button("📊 Download Excel", create_excel(df_siap_ekspor, "Data"), f"{nm_file}.xlsx", use_container_width=True)
            with dl2:
                try: st.download_button("📝 Download Word", create_word(df_siap_ekspor, "Laporan"), f"{nm_file}.docx", use_container_width=True)
                except: pass
            with dl3:
                try: st.download_button("📄 Download PDF", create_pdf(df_siap_ekspor, "Laporan"), f"{nm_file}.pdf", mime="application/pdf", use_container_width=True)
                except: pass

        c_ex1, c_ex2 = st.columns(2)
        with c_ex1:
            with st.container(border=True):
                st.markdown("##### 📥 2. Impor Data Massal (Sistem Anti Ganda)")
                st.download_button("⬇️ Download Template Excel", create_template_excel(), "Template_Upload.xlsx")
                uploaded_file = st.file_uploader("Upload file Excel:", type=['xlsx'])
                if uploaded_file is not None:
                    if st.button("🚀 Proses Database", use_container_width=True):
                        try:
                            df_import = pd.read_excel(uploaded_file)
                            conn = get_conn_write()
                            if 'Nama Pasien' in df_import.columns and 'Keluhan Awal' in df_import.columns:
                                dup_count = 0
                                dup_names = []
                                ins_count = 0
                                for index, row in df_import.iterrows():
                                    if str(row['Nama Pasien']) == 'Contoh Nama': continue
                                    imp_waktu = str(row['Tanggal Laporan']) if pd.notna(row['Tanggal Laporan']) else datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    imp_nama = str(row['Nama Pasien']) if pd.notna(row['Nama Pasien']) else ""
                                    imp_keluhan = str(row['Keluhan Awal']) if pd.notna(row['Keluhan Awal']) else ""
                                    
                                    if imp_nama and imp_keluhan:
                                        tgl_cek = imp_waktu[:10]
                                        if cek_duplikasi(conn, imp_nama, imp_keluhan, tgl_cek):
                                            dup_count += 1
                                            if imp_nama not in dup_names: dup_names.append(imp_nama)
                                        else:
                                            conn.execute('''INSERT INTO pipp (waktu_input, waktu_selesai, no_bpjs, nama_pasien, no_telp, jenis_laporan, keluhan, solusi, petugas, status) VALUES (?,?,?,?,?,?,?,?,?,?)''', (imp_waktu, imp_waktu, str(row.get('No. Identitas (BPJS/NIK)','')), imp_nama, str(row.get('No. WhatsApp/HP','')), str(row.get('Kategori','')), imp_keluhan, str(row.get('Tindakan / Solusi','')), str(row.get('Nama Petugas','')), str(row.get('Status','Baru'))))
                                            ins_count += 1
                                conn.commit(); st.cache_data.clear()
                                st.success(f"✅ Impor Berhasil! ({ins_count} Data masuk).")
                                if dup_count > 0:
                                    st.error(f"⛔ SISTEM MENCEGAH {dup_count} DATA GANDA! \nPasien berikut mencoba menginput keluhan yang sama di hari yang sama: {', '.join(dup_names)}")
                            else: st.error("Gagal! Gunakan template excel yang disediakan.")
                        except Exception as e: st.error(f"Error membaca Excel: {e}")

        with c_ex2:
            with st.container(border=True):
                st.markdown("##### 📤 3. Backup Full Database")
                try:
                    df_arsip_rapi = prepare_export_df(df_master, ['id', 'waktu_input', 'nama_pasien', 'no_bpjs', 'no_telp', 'alamat', 'jenis_pasien', 'keluhan', 'tensi', 'suhu', 'jenis_laporan', 'subkategori', 'solusi', 'petugas', 'status'])
                    st.download_button("⬇️ Download Full Backup", create_excel(df_arsip_rapi, "Arsip"), f"Backup_Full_{datetime.now().strftime('%Y%m%d')}.xlsx", type="primary", use_container_width=True)
                except Exception as e: st.error(f"Gagal memproses file: {e}")

    # =========================================================
    # TAB 5: INPUT MANUAL & GEMINI IMPORTER (ANTI GANDA)
    # =========================================================
    with tab5:
        tab_man, tab_gem = st.tabs(["✍️ Input Manual Form", "🤖 Import Data Cepat dari Gemini AI"])
        
        with tab_man:
            with st.form("form_manual_staf"):
                st.info("Sistem ini dilengkapi sensor Pencegah Duplikasi. Pasien tidak bisa lapor keluhan yang sama dua kali di hari yang sama.")
                c_tgl, c_nama = st.columns([1, 2])
                with c_tgl: m_tgl = st.date_input("🗓️ Tgl Logbook *", value=datetime.now().date())
                with c_nama: m_nama = st.text_input("👤 Nama Pasien *")
                col_m1, col_m2 = st.columns(2)
                with col_m1: m_bpjs = st.text_input("🏥 Nomor BPJS/NIK"); m_jns = st.selectbox("Jenis Pasien", ["Umum", "BPJS", "Asuransi Lain"])
                with col_m2: m_telp = st.text_input("📞 Nomor WA/HP"); m_prio = st.selectbox("Prioritas", ["Normal", "Penting", "Mendesak"])
                m_jenis = st.text_input("Kategori Laporan / Masalah")
                m_keluhan = st.text_area("💬 Detail Keluhan *")
                m_solusi = st.text_area("✅ Tindakan/Solusi")
                
                if st.form_submit_button("Simpan Laporan Baru"):
                    if m_nama and m_keluhan:
                        conn = get_conn_write()
                        tgl_str = m_tgl.strftime('%Y-%m-%d')
                        if cek_duplikasi(conn, m_nama, m_keluhan, tgl_str):
                            st.error(f"⛔ DUPLIKASI DICEGAH: Pasien atas nama '{m_nama}' sudah pernah melaporkan keluhan yang sama persis pada tanggal {tgl_str}. Mohon periksa kembali!")
                        else:
                            wkt_skrg = f"{tgl_str} {datetime.now().strftime('%H:%M:%S')}"
                            conn.execute('''INSERT INTO pipp (waktu_input, waktu_selesai, no_bpjs, nama_pasien, alamat, no_telp, jenis_laporan, keluhan, solusi, status, priority, jenis_pasien) VALUES (?,?,?,?,'-',?,?,?,?,'Baru',?,?)''', (wkt_skrg, wkt_skrg, m_bpjs, m_nama, m_telp, m_jenis, m_keluhan, m_solusi, m_prio, m_jns))
                            conn.commit(); st.cache_data.clear(); st.success("Data berhasil ditambahkan!"); st.rerun()
                    else: st.error("Nama dan Keluhan wajib diisi.")

        with tab_gem:
            st.info("Anda bisa mengunggah foto logbook langsung ke kolom obrolan **Gemini AI**, lalu gunakan perintah teks (Prompt) di bawah ini agar Gemini mengubah foto tersebut menjadi format teks yang rapi. Setelah itu, *paste* hasilnya ke kotak di bawah.")
            st.code("Tolong ekstrak foto data logbook ini menjadi teks. Susun dalam format persis seperti ini (pisahkan dengan garis lurus '|'):\nTanggal (YYYY-MM-DD) | Nama Pasien | NIK/BPJS | No HP | Kategori | Keluhan | Solusi\n\nJika ada yang kosong, tulis strip (-).", language="text")
            
            gemini_text = st.text_area("Paste hasil teks dari Gemini di sini:", height=150, placeholder="2026-09-14 | Sri Rahayu | 123456789 | 08123 | Umum | Pusing | Paracetamol\n2026-09-14 | Budi | 987654321 | 08999 | BPJS | Demam | Rujuk")
            
            if st.button("🔍 Review & Cek Duplikasi", type="primary"):
                if gemini_text:
                    lines = gemini_text.strip().split('\n')
                    parsed_data = []
                    conn = get_conn_write()
                    
                    for line in lines:
                        if '|' in line:
                            parts = [p.strip() for p in line.split('|')]
                            if len(parts) >= 6:
                                tgl = parts[0] if parts[0] != '-' else datetime.now().strftime('%Y-%m-%d')
                                nama = parts[1]
                                nik = parts[2]
                                hp = parts[3]
                                kat = parts[4]
                                keluhan = parts[5]
                                solusi = parts[6] if len(parts) > 6 else "-"
                                
                                is_dup = cek_duplikasi(conn, nama, keluhan, tgl[:10])
                                status_dup = "❌ GANDA (Ditolak)" if is_dup else "✅ AMAN"
                                
                                parsed_data.append({
                                    "Tgl": tgl, "Nama": nama, "NIK": nik, "No_HP": hp, 
                                    "Kategori": kat, "Keluhan": keluhan, "Solusi": solusi,
                                    "Status Cek": status_dup, "is_dup": is_dup
                                })
                    
                    if parsed_data:
                        st.session_state.gemini_parsed = parsed_data
                        st.success("Teks berhasil dibaca! Silakan periksa hasil pratinjau di bawah ini sebelum menyimpan.")
                    else:
                        st.error("Format tidak dikenali. Pastikan Anda menyalin teks hasil Gemini yang menggunakan pemisah garis lurus (|).")
                        
            if 'gemini_parsed' in st.session_state and st.session_state.gemini_parsed:
                df_parsed = pd.DataFrame(st.session_state.gemini_parsed)
                st.dataframe(df_parsed.drop(columns=['is_dup']), use_container_width=True)
                
                if st.button("💾 Simpan Data 'Aman' ke Database", type="primary"):
                    conn = get_conn_write()
                    masuk = 0
                    ditolak = 0
                    wkt_skrg_time = datetime.now().strftime('%H:%M:%S')
                    
                    for item in st.session_state.gemini_parsed:
                        if not item['is_dup']:
                            wkt_gabung = f"{item['Tgl'][:10]} {wkt_skrg_time}" if len(item['Tgl']) >= 10 else datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            conn.execute('''INSERT INTO pipp (waktu_input, waktu_selesai, no_bpjs, nama_pasien, alamat, no_telp, jenis_laporan, keluhan, solusi, status, priority, jenis_pasien) VALUES (?,?,?,?,'-',?,?,?,?,'Baru','Normal','Umum')''', (wkt_gabung, wkt_gabung, item['NIK'], item['Nama'], item['No_HP'], item['Kategori'], item['Keluhan'], item['Solusi']))
                            masuk += 1
                        else:
                            ditolak += 1
                    
                    conn.commit(); st.cache_data.clear()
                    st.session_state.gemini_parsed = None 
                    st.success(f"Proses Selesai! {masuk} data berhasil disimpan. {ditolak} data ganda ditolak dan dibuang.")
                    st.rerun()

    # =========================================================
    # TAB 6: KONTAK & KATEGORI
    # =========================================================
    with tab6:
        st.markdown("#### 👥 Database Kontak & Kategori")
        try: st.dataframe(pd.read_sql_query("SELECT * FROM contacts ORDER BY last_interaction_at DESC", conn), use_container_width=True)
        except: pass
        try: st.dataframe(pd.read_sql_query("SELECT * FROM main_categories", conn), use_container_width=True)
        except: pass
