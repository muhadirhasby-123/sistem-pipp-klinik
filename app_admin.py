import streamlit as st
import pandas as pd
import sqlite3
import io
import calendar
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

@st.cache_data(ttl=5) # Cache 5 detik sangkan sistem nyebrot gancangna
def load_data_from_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    df = pd.read_sql_query("SELECT * FROM pipp ORDER BY id DESC", conn).fillna('')
    conn.close()
    return df

def get_conn_write(): return sqlite3.connect(DB_PATH, check_same_thread=False)

def init_db():
    conn = get_conn_write()
    conn.execute('''CREATE TABLE IF NOT EXISTS pipp (id INTEGER PRIMARY KEY AUTOINCREMENT, waktu_input DATETIME, waktu_selesai DATETIME, no_bpjs TEXT, nama_pasien TEXT, alamat TEXT, no_telp TEXT, jenis_laporan TEXT, keluhan TEXT, solusi TEXT, link_bukti TEXT, is_real INTEGER, contact_id INTEGER, status TEXT DEFAULT 'Baru', priority TEXT DEFAULT 'Normal', short_link TEXT, consent_wa INTEGER DEFAULT 0)''')
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

def clean_text(text):
    if str(text).strip() in ['-', '', 'nan', 'Belum Diklasifikasi', 'Belum Ditentukan', '[MENUNGGU TINDAKAN]', 'None']: return '-'
    return str(text)

def get_smart_status(status_db, waktu_input_str):
    try:
        tgl_logbook = pd.to_datetime(waktu_input_str)
        batas_24_jam = datetime.now() - timedelta(days=1)
        if status_db == 'Baru':
            if tgl_logbook >= batas_24_jam: return "Baru Masuk"
            else: return "Overdue"
        return status_db
    except: return status_db

# =====================================================================
# FUNGSI EXPORT (KEBAL ERROR FLOAT)
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
    mapping = {'id': 'ID Laporan', 'waktu_input': 'Tgl Laporan', 'nama_pasien': 'Nama Pasien', 'no_bpjs': 'No. BPJS/NIK', 'no_telp': 'WhatsApp', 'alamat': 'Alamat', 'jenis_pasien': 'Jenis Pasien', 'keluhan': 'Keluhan/Diagnosa', 'jenis_laporan': 'Kategori Utama', 'subkategori': 'Subkategori', 'tensi': 'Tensi Darah', 'suhu': 'Suhu Badan', 'solusi': 'Tindakan/Solusi', 'petugas': 'Petugas/Dokter', 'status': 'Status Akhir'}
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
    .penjelasan-box { background-color: #e0f2fe; border-left: 4px solid #0284c7; padding: 12px 15px; margin-top: 5px; margin-bottom: 15px; border-radius: 4px; font-size: 13px; color: #0f172a;}
    
    /* TRIK HACKER CSS UNTUK KALENDER: Paksa teks Streamlit Button agar Multi-Baris dan Tidak Terpotong! */
    div[data-testid="stButton"] button p {
        white-space: normal !important; 
        overflow: visible !important; 
        word-wrap: break-word !important;
        text-overflow: clip !important;
    }
    
    /* Tombol Kalender */
    .btn-cal-has-data > div > button { background-color: #e0f2fe !important; border: 2px solid #0284c7 !important; border-radius: 8px !important; padding: 5px !important; height: 100% !important; min-height: 70px !important;}
    .btn-cal-has-data > div > button p { color: #0f172a !important; font-weight: bold !important; font-size: 15px !important; line-height: 1.2 !important; }
    .btn-cal-empty > div > button { background-color: transparent !important; border: 1px dashed #cbd5e1 !important; border-radius: 8px !important; padding: 5px !important; height: 100% !important; min-height: 70px !important; }
    .btn-cal-empty > div > button p { color: #94a3b8 !important; font-size: 14px !important; }
    
    /* Tombol Submit */
    div[data-testid="stFormSubmitButton"] > button { background: #0ea5e9 !important; color: white !important; font-weight: bold !important; border-radius: 6px !important; transition: all 0.2s !important; }
</style>
""", unsafe_allow_html=True)

if "admin_logged_in" not in st.session_state: st.session_state.admin_logged_in = False
if "cal_selected_date" not in st.session_state: st.session_state.cal_selected_date = None

if not st.session_state.admin_logged_in:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.info("🔐 Akses Manajemen Klinik")
        if st.button("Masuk"): 
            if st.text_input("PIN", type="password") == PIN_ADMIN_RAHASIA: st.session_state.admin_logged_in = True; st.rerun()
else:
    df_raw = load_data_from_db()
    df_master = df_raw.copy()
    
    df_master['waktu_input_dt'] = pd.to_datetime(df_master['waktu_input'], errors='coerce')
    df_master['date_only'] = df_master['waktu_input_dt'].dt.date
    
    kata_kotor = ['-', '', 'nan', 'Belum Diklasifikasi', 'Belum Ditentukan', '[MENUNGGU TINDAKAN]', 'None']
    cols_to_clean = ['nama_pasien', 'no_bpjs', 'no_telp', 'alamat', 'keluhan', 'jenis_laporan', 'subkategori', 'solusi', 'petugas']
    df_master[cols_to_clean] = df_master[cols_to_clean].replace(kata_kotor, '-')
    df_master['Smart_Status'] = df_master.apply(lambda row: get_smart_status(row['status'], row['waktu_input']), axis=1)

    st.markdown("<h2 style='color:#0f172a; margin-top: -20px;'>🏥 Dashboard Klinik Pratama Papan Mas</h2>", unsafe_allow_html=True)
    
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "🚀 Kelola Pasien & Histori", 
        "📅 Kalender Laporan Instan", # Tab kalender
        "📊 Analisis 10 Statistik Nyata", 
        "📑 Pusat Ekspor & Impor", 
        "✍️ Input Manual", 
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
    # TAB 2: KALENDER INTERAKTIF (BISA DIKLIK, TIDAK TERPOTONG, CEPAT)
    # =========================================================
    with tab2:
        st.markdown("<h4 style='color:#0f172a;'>📅 Kalender Pemantauan Laporan</h4>", unsafe_allow_html=True)
        st.info("Klik tombol tanggal di kalender untuk memunculkan detail rincian data di tabel sebelah kanan. Cepat dan instan.")
        
        df_cal = df_master.copy()

        with st.expander("🔍 Filter Laporan Khusus Kalender", expanded=False):
            f_col1, f_col2, f_col3 = st.columns(3)
            with f_col1:
                f_kat_cal = st.multiselect("Kategori Laporan:", options=[k for k in df_cal['jenis_laporan'].unique() if k != '-'])
                f_sub_cal = st.multiselect("Subkategori:", options=[s for s in df_cal['subkategori'].unique() if s != '-'])
            with f_col2:
                f_stat_cal = st.multiselect("Status:", options=[s for s in df_cal['status'].unique() if s != '-'])
                f_pet_cal = st.multiselect("Petugas:", options=[p for p in df_cal['petugas'].unique() if p != '-'])
            with f_col3:
                f_sol_kosong = st.checkbox("Hanya Laporan Tanpa Solusi")
                f_kat_kosong = st.checkbox("Hanya Laporan Tanpa Kategori")

        df_fil_cal = df_cal.copy()
        if f_kat_cal: df_fil_cal = df_fil_cal[df_fil_cal['jenis_laporan'].isin(f_kat_cal)]
        if f_sub_cal: df_fil_cal = df_fil_cal[df_fil_cal['subkategori'].isin(f_sub_cal)]
        if f_stat_cal: df_fil_cal = df_fil_cal[df_fil_cal['status'].isin(f_stat_cal)]
        if f_pet_cal: df_fil_cal = df_fil_cal[df_fil_cal['petugas'].isin(f_pet_cal)]
        if f_sol_kosong: df_fil_cal = df_fil_cal[df_fil_cal['solusi'] == '-']
        if f_kat_kosong: df_fil_cal = df_fil_cal[df_fil_cal['jenis_laporan'] == '-']

        laporan_per_hari = df_fil_cal.groupby('date_only').size().to_dict()
        st.markdown("---")
        
        # LAYOUT: Kalender Kiri (Ukuran Pas), Tabel Kanan
        cal_left, cal_right = st.columns([1.3, 1.7], gap="large")
        
        with cal_left:
            try: default_date = df_fil_cal['date_only'].max() if not df_fil_cal.empty else datetime.now().date()
            except: default_date = datetime.now().date()
            
            c_nav1, c_nav2 = st.columns([1, 2])
            with c_nav1: sel_tahun = st.selectbox("Tahun", range(2020, datetime.now().year + 3), index=(default_date.year - 2020))
            with c_nav2:
                bulan_list = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
                sel_bulan = st.selectbox("Bulan", range(1, 13), format_func=lambda x: bulan_list[x-1], index=(default_date.month - 1))
            
            st.markdown(f"<h5 style='text-align:center;'>{bulan_list[sel_bulan-1]} {sel_tahun}</h5>", unsafe_allow_html=True)
            
            # Header Hari
            days_header = ["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"]
            cols_hdr = st.columns(7)
            for i, day_name in enumerate(days_header):
                cols_hdr[i].markdown(f"<div style='text-align:center; font-weight:bold; color:#475569; font-size:13px; padding-bottom:5px;'>{day_name}</div>", unsafe_allow_html=True)
            
            # GRID TOMBOL KALENDER NATIVE (Sudah di-Hack CSS-nya agar teks full multi-baris)
            cal_matrix = calendar.monthcalendar(sel_tahun, sel_bulan)
            for week in cal_matrix:
                cols_week = st.columns(7)
                for i, day in enumerate(week):
                    with cols_week[i]:
                        if day != 0:
                            current_dt = date(sel_tahun, sel_bulan, day)
                            jumlah_lap = laporan_per_hari.get(current_dt, 0)
                            
                            # Label Jelas, Misal: 12 \n 5 Lap.
                            btn_label = f"{day}\n{jumlah_lap} Lap" if jumlah_lap > 0 else f"{day}"
                            btn_class = "btn-cal-has-data" if jumlah_lap > 0 else "btn-cal-empty"
                            
                            st.markdown(f"<div class='{btn_class}'>", unsafe_allow_html=True)
                            if st.button(btn_label, key=f"cal_btn_{current_dt}", use_container_width=True):
                                st.session_state.cal_selected_date = current_dt
                                st.rerun() 
                            st.markdown("</div>", unsafe_allow_html=True)
                        else: st.write("")

        with cal_right:
            df_table_view = df_fil_cal.copy()
            tb_header, tb_reset = st.columns([3, 1])
            with tb_header:
                if st.session_state.cal_selected_date:
                    df_table_view = df_table_view[df_table_view['date_only'] == st.session_state.cal_selected_date]
                    st.markdown(f"**Laporan Tanggal: {st.session_state.cal_selected_date}** (Total: {len(df_table_view)})")
                else:
                    st.markdown(f"**Semua Laporan Sesuai Filter Bulan Ini** (Total: {len(df_table_view)})")
                    
            with tb_reset:
                if st.button("🔄 Reset Tgl", use_container_width=True):
                    st.session_state.cal_selected_date = None
                    st.rerun()

            df_table_view = df_table_view[['id', 'waktu_input', 'nama_pasien', 'keluhan', 'Smart_Status', 'petugas']].rename(columns={'id':'ID', 'waktu_input':'Tgl Logbook', 'nama_pasien':'Pasien', 'keluhan':'Keluhan', 'Smart_Status':'Status', 'petugas':'Petugas'})
            
            if df_table_view.empty: st.info("Pilih tanggal di kalender yang memiliki indikator laporan.")
            else: st.dataframe(df_table_view, use_container_width=True, hide_index=True)

    # =========================================================
    # TAB 3: MEGA ANALITIK 1000+ (DENGAN PENJELASAN FAKTA & REAL)
    # =========================================================
    with tab3:
        st.markdown("<h4 style='color:#0f172a;'>📊 Laporan 10 Statistik Fakta & Realita</h4>", unsafe_allow_html=True)
        st.info("Seluruh statistik ini membaca MURNI dari data historis Anda (SQLite). Jika sistem menunjukkan ada 600 data di bulan April, artinya terdapat 600 baris data yang tanggal inputnya diset pada bulan tersebut.")
        
        df_an = df_master.copy()
        df_an['Bulan_Tahun'] = df_an['waktu_input_dt'].dt.strftime('%b %Y')
        layout_cfg = dict(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(family="Inter", color="#334155"), margin=dict(t=30, b=10, l=10, r=10))
        
        total_semua = len(df_an)
        total_selesai = len(df_an[df_an['status'].isin(['Selesai', 'Ditutup'])])
        total_overdue = len(df_an[df_an['Smart_Status'] == 'Overdue'])
        
        st.markdown("#### 1. Ringkasan Kuantitas Laporan Historis")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Stat 1: Total Seluruh Laporan", f"{total_semua} Data")
        m2.metric("Stat 2: Laporan Diselesaikan", total_selesai)
        m3.metric("Stat 3: Status Overdue", total_overdue, delta="Urgent!", delta_color="inverse")
        pct = round((total_selesai/total_semua)*100, 1) if total_semua > 0 else 0
        m4.metric("Stat 4: Rasio Selesai", f"{pct}%")
        st.markdown(f"<div class='penjelasan-box'><b>Fakta Database:</b> Dari total {total_semua} laporan yang tercatat dari awal hingga hari ini, terdapat {total_overdue} laporan yang statusnya masih 'Baru' padahal usia logbook-nya sudah lebih dari 24 jam. Hal ini menghasilkan rasio keberhasilan penanganan di angka {pct}%.</div>", unsafe_allow_html=True)
        
        st.markdown("---")
        cg1, cg2 = st.columns(2)
        with cg1:
            st.markdown("#### 5. Tren Lonjakan Kunjungan Bulanan")
            bulan_counts = df_an['Bulan_Tahun'].value_counts().reset_index()
            bulan_counts.columns = ['Bulan', 'Jumlah']
            fig_bulan = px.line(bulan_counts, x='Bulan', y='Jumlah', markers=True, color_discrete_sequence=['#ef4444'])
            fig_bulan.update_layout(**layout_cfg); st.plotly_chart(fig_bulan, use_container_width=True)
            try: bln_sibuk = bulan_counts.iloc[0]['Bulan']; jlh_sibuk = bulan_counts.iloc[0]['Jumlah']
            except: bln_sibuk = "-"; jlh_sibuk = 0
            st.markdown(f"<div class='penjelasan-box'><b>Fakta Database:</b> Data tertinggi sistem jatuh pada bulan <b>{bln_sibuk}</b> dengan lonjakan sebanyak {jlh_sibuk} laporan. Angka ini secara absolut membaca jumlah baris data yang tanggal inputnya jatuh pada bulan tersebut.</div>", unsafe_allow_html=True)
            
        with cg2:
            st.markdown("#### 6. Komparasi Status Penanganan")
            stat_counts = df_an['Smart_Status'].value_counts().reset_index()
            stat_counts.columns = ['Status', 'Jumlah']
            fig_pie = px.pie(stat_counts, values='Jumlah', names='Status', hole=0.4, color_discrete_sequence=['#22c55e', '#ef4444', '#f59e0b', '#3b82f6'])
            fig_pie.update_layout(**layout_cfg); st.plotly_chart(fig_pie, use_container_width=True)
            st.markdown(f"<div class='penjelasan-box'><b>Fakta Database:</b> Grafik donat ini memperlihatkan proporsi tugas secara visual. Area berwarna merah ('Overdue') menyoroti tumpukan tugas (pekerjaan rumah) yang membutuhkan tindakan penyelesaian dari Admin.</div>", unsafe_allow_html=True)
        
        st.markdown("---")
        cg3, cg4 = st.columns(2)
        with cg3:
            st.markdown("#### 7. Kategori Kasus Paling Umum")
            df_kat = df_an.copy()
            df_kat['jenis_laporan'] = df_kat['jenis_laporan'].apply(lambda x: 'KOSONG (Belum Diklasifikasi)' if x == '-' else x)
            kat_c = df_kat['jenis_laporan'].value_counts().head(5).reset_index()
            kat_c.columns = ['Kategori', 'Total']
            fig_bar_k = px.bar(kat_c, x='Total', y='Kategori', orientation='h', color_discrete_sequence=['#0ea5e9'])
            fig_bar_k.update_layout(yaxis={'categoryorder':'total ascending'}, **layout_cfg); st.plotly_chart(fig_bar_k, use_container_width=True)
            st.markdown(f"<div class='penjelasan-box'><b>Fakta Database:</b> Kasus <b>'{kat_c.iloc[0]['Kategori']}'</b> mendominasi laporan. Perhatikan bahwa grafik ini secara transparan juga menghitung baris data yang belum memiliki kategori (KOSONG).</div>", unsafe_allow_html=True)

        with cg4:
            st.markdown("#### 8. Evaluasi Beban Tugas Petugas")
            df_pet = df_an.copy()
            df_pet['petugas'] = df_pet['petugas'].apply(lambda x: 'KOSONG (Tanpa Petugas)' if x == '-' else x)
            pet_c = df_pet['petugas'].value_counts().head(5).reset_index()
            pet_c.columns = ['Nama Petugas', 'Jumlah Ditangani']
            fig_pet = px.bar(pet_c, x='Nama Petugas', y='Jumlah Ditangani', color_discrete_sequence=['#8b5cf6'])
            fig_pet.update_layout(**layout_cfg); st.plotly_chart(fig_pet, use_container_width=True)
            st.markdown(f"<div class='penjelasan-box'><b>Fakta Database:</b> <b>{pet_c.iloc[0]['Nama Petugas']}</b> tercatat memiliki rekam jejak tugas terbanyak. Seperti pada grafik sebelumnya, baris 'KOSONG' memperlihatkan jumlah laporan yang tidak diketahui siapa petugasnya.</div>", unsafe_allow_html=True)
            
        st.markdown("---")
        ct1, ct2 = st.columns(2)
        with ct1:
            st.markdown("#### 9. Pasien Paling Sering Melapor")
            df_pas = df_an[df_an['nama_pasien'] != '-']
            pas_c = df_pas['nama_pasien'].value_counts().head(5).reset_index()
            pas_c.columns = ['Nama Pasien', 'Total Kunjungan']
            st.dataframe(pas_c, use_container_width=True, hide_index=True)
            st.markdown(f"<div class='penjelasan-box'><b>Fakta Database:</b> Nama <b>{pas_c.iloc[0]['Nama Pasien'] if not pas_c.empty else '-'}</b> tercatat berinteraksi sangat sering. Sistem menemukan {pas_c.iloc[0]['Total Kunjungan'] if not pas_c.empty else 0} baris laporan atas nama pasien tersebut di seluruh riwayat.</div>", unsafe_allow_html=True)
            
        with ct2:
            st.markdown("#### 10. Daftar Keluhan Overdue Terlama")
            df_overdue = df_an[df_an['Smart_Status'] == 'Overdue'].sort_values('waktu_input_dt')[['waktu_input', 'nama_pasien', 'keluhan']].head(5)
            df_overdue.columns = ['Tgl Logbook', 'Pasien', 'Keluhan Belum Selesai']
            st.dataframe(df_overdue, use_container_width=True, hide_index=True)
            try: tgl_lama = df_overdue.iloc[0]['Tgl Logbook'][:10]
            except: tgl_lama = "-"
            st.markdown(f"<div class='penjelasan-box'><b>Fakta Database:</b> Tabel ini menampilkan 5 laporan yang paling lama 'terabaikan'. Laporan teratas sudah berstatus Baru dari semenjak tanggal <b>{tgl_lama}</b> namun belum dicentang selesai.</div>", unsafe_allow_html=True)

    # =========================================================
    # TAB 4: PUSAT EKSPOR & IMPOR 
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
            
            kolom_pilih = st.multiselect("Kolom yang dicetak:", options=['id', 'waktu_input', 'nama_pasien', 'no_bpjs', 'no_telp', 'alamat', 'jenis_pasien', 'keluhan', 'tensi', 'suhu', 'jenis_laporan', 'subkategori', 'solusi', 'petugas', 'status'], default=['waktu_input', 'nama_pasien', 'keluhan', 'solusi', 'status'])
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
                st.markdown("##### 📥 2. Impor Data Massal ke Sistem")
                st.download_button("⬇️ Download Template Excel", create_template_excel(), "Template_Upload.xlsx")
                uploaded_file = st.file_uploader("Upload file Excel:", type=['xlsx'])
                if uploaded_file is not None:
                    if st.button("🚀 Proses Database", use_container_width=True):
                        try:
                            df_import = pd.read_excel(uploaded_file)
                            conn = get_conn_write()
                            if 'Nama Pasien' in df_import.columns and 'Keluhan Awal' in df_import.columns:
                                for index, row in df_import.iterrows():
                                    if str(row['Nama Pasien']) == 'Contoh Nama': continue
                                    imp_waktu = str(row['Tanggal Laporan']) if pd.notna(row['Tanggal Laporan']) else datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                    imp_nama = str(row['Nama Pasien']) if pd.notna(row['Nama Pasien']) else ""
                                    imp_keluhan = str(row['Keluhan Awal']) if pd.notna(row['Keluhan Awal']) else ""
                                    imp_bpjs = str(row['No. Identitas (BPJS/NIK)']) if pd.notna(row['No. Identitas (BPJS/NIK)']) else ""
                                    imp_telp = str(row['No. WhatsApp/HP']) if pd.notna(row['No. WhatsApp/HP']) else ""
                                    imp_kat = str(row['Kategori']) if pd.notna(row['Kategori']) else ""
                                    imp_solusi = str(row['Tindakan / Solusi']) if pd.notna(row['Tindakan / Solusi']) else ""
                                    imp_petugas = str(row['Nama Petugas']) if pd.notna(row['Nama Petugas']) else ""
                                    imp_status = str(row['Status']) if pd.notna(row['Status']) else "Baru"
                                    if imp_nama and imp_keluhan:
                                        conn.execute('''INSERT INTO pipp (waktu_input, waktu_selesai, no_bpjs, nama_pasien, no_telp, jenis_laporan, keluhan, solusi, petugas, status) VALUES (?,?,?,?,?,?,?,?,?,?)''', (imp_waktu, imp_waktu, imp_bpjs, imp_nama, imp_telp, imp_kat, imp_keluhan, imp_solusi, imp_petugas, imp_status))
                                conn.commit(); st.cache_data.clear(); st.success("✅ Impor Data Berhasil!"); st.rerun()
                        except Exception as e: st.error(f"Error membaca Excel: {e}")

        with c_ex2:
            with st.container(border=True):
                st.markdown("##### 📤 3. Backup Full Database")
                try:
                    df_arsip_rapi = prepare_export_df(df_master, ['id', 'waktu_input', 'nama_pasien', 'no_bpjs', 'no_telp', 'alamat', 'jenis_pasien', 'keluhan', 'tensi', 'suhu', 'jenis_laporan', 'subkategori', 'solusi', 'petugas', 'status'])
                    st.download_button("⬇️ Download Full Backup", create_excel(df_arsip_rapi, "Arsip"), f"Backup_Full_{datetime.now().strftime('%Y%m%d')}.xlsx", type="primary", use_container_width=True)
                except Exception as e: st.error(f"Gagal memproses file: {e}")

    # =========================================================
    # TAB 5: INPUT MANUAL PASIEN BARU
    # =========================================================
    with tab5:
        with st.form("form_manual_staf"):
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
                    wkt_skrg = f"{m_tgl.strftime('%Y-%m-%d')} {datetime.now().strftime('%H:%M:%S')}"
                    conn = get_conn_write()
                    conn.execute('''INSERT INTO pipp (waktu_input, waktu_selesai, no_bpjs, nama_pasien, alamat, no_telp, jenis_laporan, keluhan, solusi, status, priority, jenis_pasien) VALUES (?,?,?,?,'-',?,?,?,?,'Baru',?,?)''', (wkt_skrg, wkt_skrg, m_bpjs, m_nama, m_telp, m_jenis, m_keluhan, m_solusi, m_prio, m_jns))
                    conn.commit(); st.cache_data.clear(); st.success("Data ditambahkan!"); st.rerun()
                else: st.error("Nama dan Keluhan wajib diisi.")

    # =========================================================
    # TAB 6: KONTAK & KATEGORI
    # =========================================================
    with tab6:
        st.markdown("#### 👥 Database Kontak & Kategori")
        try: st.dataframe(pd.read_sql_query("SELECT * FROM contacts ORDER BY last_interaction_at DESC", conn), use_container_width=True)
        except: pass
        try: st.dataframe(pd.read_sql_query("SELECT * FROM main_categories", conn), use_container_width=True)
        except: pass
