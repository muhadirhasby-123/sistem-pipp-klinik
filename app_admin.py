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
PIN_ADMIN_RAHASIA = "123456" # <-- GANTI DENGAN PIN ASLI
DB_PATH = "klinik_pipp.db"

@st.cache_data(ttl=2)
def load_data_from_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    df = pd.read_sql_query("SELECT * FROM pipp ORDER BY id DESC", conn).fillna('')
    conn.close()
    return df

def get_conn_write(): return sqlite3.connect(DB_PATH, check_same_thread=False)

def init_db():
    conn = get_conn_write()
    conn.execute('''CREATE TABLE IF NOT EXISTS pipp (id INTEGER PRIMARY KEY AUTOINCREMENT, waktu_input DATETIME, waktu_selesai DATETIME, no_bpjs TEXT, nama_pasien TEXT, alamat TEXT, no_telp TEXT, jenis_laporan TEXT, keluhan TEXT, solusi TEXT, link_bukti TEXT, is_real INTEGER, contact_id INTEGER, status TEXT DEFAULT 'Baru', priority TEXT DEFAULT 'Normal', short_link TEXT, consent_wa INTEGER DEFAULT 0)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS contacts (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, whatsapp TEXT UNIQUE, bpjs TEXT, nik TEXT, consent_wa INTEGER DEFAULT 0, first_interaction_at DATETIME, last_interaction_at DATETIME)''')
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
# FUNGSI EXPORT (KEBAL ERROR)
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
    mapping = {'id': 'ID', 'waktu_input': 'Tgl Laporan', 'nama_pasien': 'Nama Pasien', 'no_bpjs': 'No. BPJS/NIK', 'no_telp': 'WhatsApp', 'alamat': 'Alamat', 'jenis_pasien': 'Jenis Pasien', 'keluhan': 'Keluhan', 'jenis_laporan': 'Kategori', 'subkategori': 'Subkategori', 'tensi': 'Tensi', 'suhu': 'Suhu', 'solusi': 'Tindakan', 'petugas': 'Petugas', 'status': 'Status'}
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
    
    /* Kalender Canggih Button */
    .btn-cal-has-data > button { background-color: #0284c7 !important; color: white !important; font-weight: bold !important; border: none !important; font-size:14px!important; padding:15px 0px!important; margin-bottom:8px; border-radius: 8px !important;}
    .btn-cal-empty > button { background-color: #f8fafc !important; color: #94a3b8 !important; border: 1px dashed #cbd5e1 !important; font-size:14px!important; padding:15px 0px!important; margin-bottom:8px; border-radius: 8px !important;}
    div[data-testid="stFormSubmitButton"] > button { background: #0ea5e9 !important; color: white !important; font-weight: bold !important; border-radius: 6px !important; transition: all 0.2s !important; }
</style>
""", unsafe_allow_html=True)

if "admin_logged_in" not in st.session_state: st.session_state.admin_logged_in = False

# State Kalender Canggih
if "cal_filter_type" not in st.session_state: st.session_state.cal_filter_type = None
if "cal_filter_val" not in st.session_state: st.session_state.cal_filter_val = None

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
        "🚀 Kelola Pasien & Audit", 
        "📅 Kalender Laporan Canggih", 
        "📊 Mega Analitik 10 Statistik", 
        "📑 Pusat Ekspor & Impor", 
        "✍️ Input Manual", 
        "👥 Database Kontak & Kategori"
    ])

    # =========================================================
    # TAB 1: PENCARIAN, UPDATE & AUDIT 
    # =========================================================
    with tab1:
        cari_pasien = st.text_input("🔍 Cari Pasien untuk Diupdate (Ketik Nama, No. BPJS, atau ID):", placeholder="Contoh: Sri rahayu...")
        if cari_pasien:
            sq = cari_pasien.lower()
            hasil = df_master[df_master['nama_pasien'].str.lower().str.contains(sq) | df_master['no_bpjs'].str.lower().str.contains(sq) | df_master['id'].astype(str).str.contains(sq)]
            
            if not hasil.empty:
                p = hasil.iloc[0]
                st.markdown("<div class='form-container'>", unsafe_allow_html=True)
                with st.form(f"form_update_{p['id']}"):
                    st.markdown(f"**Mengelola Laporan Pasien: {p['nama_pasien']} (ID: {p['id']})**")
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
                    
                    if st.form_submit_button("✔ Simpan Perubahan"):
                        w_skrg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        w_input_baru = f"{n_tgl.strftime('%Y-%m-%d')} {jam_asli}"
                        w_sls = p['waktu_selesai']
                        if n_stat in ["Selesai", "Ditutup"] and w_sls == '': w_sls = w_skrg
                        conn = get_conn_write()
                        conn.execute("""UPDATE pipp SET waktu_input=?, nama_pasien=?, no_bpjs=?, no_telp=?, jenis_pasien=?, tensi=?, suhu=?, jenis_laporan=?, petugas=?, solusi=?, status=?, waktu_selesai=? WHERE id=?""", (w_input_baru, n_nama, n_bpjs, n_wa, n_jns, n_tensi, n_suhu, n_kat, n_petugas, n_solusi, n_stat, w_sls, p['id']))
                        conn.commit(); st.cache_data.clear(); st.rerun()
                st.markdown("</div><br>", unsafe_allow_html=True)

        st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
        st.markdown("#### 📋 Data Antrean & Validasi (Audit 1-Klik)")
        st.info("💡 **Tips:** Centang kotak `✔ Validasi` di tabel, lalu klik 'Simpan' untuk merubah status laporan menjadi Selesai tanpa perlu buka form di atas.")
        
        df_audit = df_master.copy()
        df_audit['✔ Validasi'] = df_audit['status'].isin(['Selesai', 'Ditutup']) | (df_audit['is_audited'] == 1)
        df_edit_view = df_audit[['✔ Validasi', 'id', 'waktu_input', 'nama_pasien', 'keluhan', 'Smart_Status', 'petugas']]
        df_edit_view = df_edit_view.rename(columns={'waktu_input':'Tgl Logbook', 'nama_pasien':'Nama Pasien', 'Smart_Status':'Status'})
        
        edited_df = st.data_editor(df_edit_view, hide_index=True, use_container_width=True, disabled=['id', 'Tgl Logbook', 'Nama Pasien', 'keluhan', 'Status', 'petugas'])
        
        if st.button("💾 Simpan Validasi Tabel", type="primary"):
            w_skrg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            conn = get_conn_write()
            for index, row in edited_df.iterrows():
                if row['✔ Validasi']:
                    conn.execute("UPDATE pipp SET status = 'Selesai', is_audited = 1, waktu_selesai = ? WHERE id = ? AND status != 'Selesai'", (w_skrg, row['id']))
                    conn.execute("UPDATE pipp SET is_audited = 1 WHERE id = ?", (row['id'],))
                else: conn.execute("UPDATE pipp SET is_audited = 0 WHERE id = ?", (row['id'],))
            conn.commit(); st.cache_data.clear(); st.success("Pembaruan tersimpan!"); st.rerun()

    # =========================================================
    # TAB 2: KALENDER CANGGIH 3 DIMENSI (HARIAN, BULANAN, TAHUNAN)
    # =========================================================
    with tab2:
        st.markdown("<h4 style='color:#0f172a;'>📅 Kalender Laporan Multi-Dimensi</h4>", unsafe_allow_html=True)
        
        mode_kalender = st.radio("Pilih Tingkat Tampilan Kalender:", ["📅 Harian (Lihat per Bulan)", "📊 Bulanan (Lihat per Tahun)", "📈 Tahunan (Lihat Semua Tahun)"], horizontal=True)
        
        df_cal = df_master.copy()
        with st.expander("🔍 Filter Khusus Kalender", expanded=False):
            f_col1, f_col2 = st.columns(2)
            with f_col1:
                f_kat_cal = st.multiselect("Kategori:", options=[k for k in df_cal['jenis_laporan'].unique() if k != '-'])
                f_stat_cal = st.multiselect("Status:", options=[s for s in df_cal['status'].unique() if s != '-'])
            with f_col2:
                f_pet_cal = st.multiselect("Petugas:", options=[p for p in df_cal['petugas'].unique() if p != '-'])
                f_sol_kosong = st.checkbox("Tampilkan HANYA Laporan Tanpa Solusi")

        if f_kat_cal: df_cal = df_cal[df_cal['jenis_laporan'].isin(f_kat_cal)]
        if f_stat_cal: df_cal = df_cal[df_cal['status'].isin(f_stat_cal)]
        if f_pet_cal: df_cal = df_cal[df_cal['petugas'].isin(f_pet_cal)]
        if f_sol_kosong: df_cal = df_cal[df_cal['solusi'] == '-']

        st.markdown("---")
        cal_left, cal_right = st.columns([1.5, 2], gap="large")
        
        try: default_date = df_cal['date_only'].max() if not df_cal.empty else datetime.now().date()
        except: default_date = datetime.now().date()
        
        with cal_left:
            if mode_kalender == "📅 Harian (Lihat per Bulan)":
                c_nav1, c_nav2 = st.columns(2)
                with c_nav1: sel_thn = st.selectbox("Pilih Tahun:", range(2020, datetime.now().year + 3), index=(default_date.year - 2020))
                with c_nav2:
                    bln_list = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
                    sel_bln = st.selectbox("Pilih Bulan:", range(1, 13), format_func=lambda x: bln_list[x-1], index=(default_date.month - 1))
                
                lap_harian = df_cal[(df_cal['waktu_input_dt'].dt.year == sel_thn) & (df_cal['waktu_input_dt'].dt.month == sel_bln)].groupby('date_only').size().to_dict()
                
                st.markdown(f"<h5 style='text-align:center;'>{bln_list[sel_bln-1]} {sel_thn}</h5>", unsafe_allow_html=True)
                cols_hdr = st.columns(7)
                for i, day_name in enumerate(["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"]):
                    cols_hdr[i].markdown(f"<div style='text-align:center; font-weight:bold; color:#475569; font-size:13px; padding-bottom:5px;'>{day_name}</div>", unsafe_allow_html=True)
                
                cal_matrix = calendar.monthcalendar(sel_thn, sel_bln)
                for week in cal_matrix:
                    cols_week = st.columns(7)
                    for i, day in enumerate(week):
                        with cols_week[i]:
                            if day != 0:
                                current_dt = date(sel_thn, sel_bln, day)
                                jml = lap_harian.get(current_dt, 0)
                                btn_label = f"{day}\n({jml})" if jml > 0 else f"{day}"
                                btn_class = "btn-cal-has-data" if jml > 0 else "btn-cal-empty"
                                st.markdown(f"<div class='{btn_class}'>", unsafe_allow_html=True)
                                if st.button(btn_label, key=f"d_{current_dt}", use_container_width=True):
                                    st.session_state.cal_filter_type = 'date'
                                    st.session_state.cal_filter_val = current_dt
                                    st.rerun()
                                st.markdown("</div>", unsafe_allow_html=True)
                            else: st.write("")

            elif mode_kalender == "📊 Bulanan (Lihat per Tahun)":
                sel_thn = st.selectbox("Pilih Tahun:", range(2020, datetime.now().year + 3), index=(default_date.year - 2020))
                st.markdown(f"<h5 style='text-align:center;'>Rekap Tahun {sel_thn}</h5>", unsafe_allow_html=True)
                df_thn = df_cal[df_cal['waktu_input_dt'].dt.year == sel_thn]
                lap_bln = df_thn.groupby(df_thn['waktu_input_dt'].dt.month).size().to_dict()
                
                bln_nama = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
                for row_i in range(4):
                    cols_b = st.columns(3)
                    for col_i in range(3):
                        b_idx = row_i * 3 + col_i + 1
                        jml = lap_bln.get(b_idx, 0)
                        with cols_b[col_i]:
                            btn_label = f"{bln_nama[b_idx-1]}\n({jml} Lap)"
                            btn_class = "btn-cal-has-data" if jml > 0 else "btn-cal-empty"
                            st.markdown(f"<div class='{btn_class}'>", unsafe_allow_html=True)
                            if st.button(btn_label, key=f"m_{sel_thn}_{b_idx}", use_container_width=True):
                                st.session_state.cal_filter_type = 'month'
                                st.session_state.cal_filter_val = (sel_thn, b_idx)
                                st.rerun()
                            st.markdown("</div>", unsafe_allow_html=True)

            elif mode_kalender == "📈 Tahunan (Lihat Semua Tahun)":
                st.markdown(f"<h5 style='text-align:center;'>Rekap Semua Tahun</h5>", unsafe_allow_html=True)
                lap_thn = df_cal.groupby(df_cal['waktu_input_dt'].dt.year).size().to_dict()
                available_years = sorted(list(set([y for y in lap_thn.keys() if pd.notna(y)])))
                if not available_years: available_years = [datetime.now().year]
                
                cols_y = st.columns(3)
                for i, y in enumerate(available_years):
                    jml = lap_thn.get(y, 0)
                    with cols_y[i % 3]:
                        btn_label = f"Tahun {int(y)}\n({jml} Lap)"
                        btn_class = "btn-cal-has-data" if jml > 0 else "btn-cal-empty"
                        st.markdown(f"<div class='{btn_class}'>", unsafe_allow_html=True)
                        if st.button(btn_label, key=f"y_{y}", use_container_width=True):
                            st.session_state.cal_filter_type = 'year'
                            st.session_state.cal_filter_val = y
                            st.rerun()
                        st.markdown("</div>", unsafe_allow_html=True)

        with cal_right:
            df_table_view = df_cal.copy()
            tb_header, tb_reset = st.columns([3, 1])
            
            with tb_header:
                if st.session_state.cal_filter_type == 'date':
                    dt_val = st.session_state.cal_filter_val
                    df_table_view = df_table_view[df_table_view['date_only'] == dt_val]
                    st.markdown(f"**Laporan Tanggal: {dt_val}** (Total: {len(df_table_view)})")
                elif st.session_state.cal_filter_type == 'month':
                    y, m = st.session_state.cal_filter_val
                    df_table_view = df_table_view[(df_table_view['waktu_input_dt'].dt.year == y) & (df_table_view['waktu_input_dt'].dt.month == m)]
                    bln_str = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agt", "Sep", "Okt", "Nov", "Des"][m-1]
                    st.markdown(f"**Laporan Bulan: {bln_str} {y}** (Total: {len(df_table_view)})")
                elif st.session_state.cal_filter_type == 'year':
                    y = st.session_state.cal_filter_val
                    df_table_view = df_table_view[df_table_view['waktu_input_dt'].dt.year == y]
                    st.markdown(f"**Laporan Tahun: {int(y)}** (Total: {len(df_table_view)})")
                else: st.markdown(f"**Menampilkan Semua Laporan** (Total: {len(df_table_view)})")
                    
            with tb_reset:
                if st.button("🔄 Reset Pilihan", use_container_width=True):
                    st.session_state.cal_filter_type = None; st.session_state.cal_filter_val = None; st.rerun()

            df_table_view = df_table_view[['id', 'waktu_input', 'nama_pasien', 'keluhan', 'Smart_Status', 'petugas']]
            df_table_view = df_table_view.rename(columns={'id':'ID', 'waktu_input':'Tgl Logbook', 'nama_pasien':'Pasien', 'keluhan':'Keluhan', 'Smart_Status':'Status', 'petugas':'Petugas'})
            
            if df_table_view.empty: st.info("Pilih kotak berwarna biru di kalender sebelah kiri untuk melihat rincian laporan.")
            else: st.dataframe(df_table_view, use_container_width=True, hide_index=True)

    # =========================================================
    # TAB 3: MEGA ANALITIK (10 STATISTIK NYATA & RAPIH)
    # =========================================================
    with tab3:
        st.markdown("<h4 style='color:#0f172a;'>📊 Laporan 10 Statistik Fakta & Realita</h4>", unsafe_allow_html=True)
        df_an = df_master.copy()
        df_an['Bulan_Tahun'] = df_an['waktu_input_dt'].dt.strftime('%b %Y')
        
        # PENGELOMPOKKAN SUB-TAB AGAR TIDAK NUMPUK KE HANDAP
        atab1, atab2, atab3 = st.tabs(["📌 Stat 1-4 (Kinerja & Tren)", "🏆 Stat 5-7 (Kategori & Petugas)", "👥 Stat 8-10 (Pasien & Overdue)"])
        layout_cfg = dict(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(family="Inter", color="#334155"), margin=dict(t=30, b=10, l=10, r=10))
        
        with atab1:
            total_semua = len(df_an)
            total_selesai = len(df_an[df_an['status'].isin(['Selesai', 'Ditutup'])])
            total_overdue = len(df_an[df_an['Smart_Status'] == 'Overdue'])
            
            st.markdown("**Statistik 1: Ringkasan Total Data Database**")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Seluruh Laporan", f"{total_semua}")
            m2.metric("Laporan Diselesaikan", total_selesai)
            m3.metric("Laporan Overdue", total_overdue)
            pct = round((total_selesai/total_semua)*100, 1) if total_semua > 0 else 0
            m4.metric("Rasio Keberhasilan", f"{pct}%")
            st.markdown(f"<div class='penjelasan-box'><b>Fakta Nyata:</b> Tina {total_semua} data asli anu lebet ti awal dugi ka ayeuna, aya {total_overdue} laporan anu statusna masih 'Baru' padahal umurna atos langkung ti 24 jam (Overdue). Ieu ngajantenkeun tingkat kasuksésan penanganan di klinik mung ngahontal {pct}%.</div>", unsafe_allow_html=True)
            
            st.markdown("---")
            cg1, cg2 = st.columns(2)
            with cg1:
                st.markdown("**Statistik 2: Tren Kunjungan Harian/Bulanan**")
                bulan_counts = df_an['Bulan_Tahun'].value_counts().reset_index()
                bulan_counts.columns = ['Bulan', 'Jumlah']
                fig_bulan = px.line(bulan_counts, x='Bulan', y='Jumlah', markers=True, color_discrete_sequence=['#ef4444'])
                fig_bulan.update_layout(**layout_cfg); st.plotly_chart(fig_bulan, use_container_width=True)
                try: bln_sibuk = bulan_counts.iloc[0]['Bulan']; jlh_sibuk = bulan_counts.iloc[0]['Jumlah']
                except: bln_sibuk = "-"; jlh_sibuk = 0
                st.markdown(f"<div class='penjelasan-box'><b>Fakta Nyata:</b> Grafik garis ieu maca dumasar tanggal kajadian di logbook. Puncak komplain atanapi kadatangan paling luhur murag dina bulan <b>{bln_sibuk}</b> ku total {jlh_sibuk} laporan.</div>", unsafe_allow_html=True)
                
            with cg2:
                st.markdown("**Statistik 3: Komparasi Status Penanganan**")
                stat_counts = df_an['Smart_Status'].value_counts().reset_index()
                stat_counts.columns = ['Status', 'Jumlah']
                fig_pie = px.pie(stat_counts, values='Jumlah', names='Status', hole=0.4, color_discrete_sequence=['#22c55e', '#ef4444', '#f59e0b', '#3b82f6'])
                fig_pie.update_layout(**layout_cfg); st.plotly_chart(fig_pie, use_container_width=True)
                st.markdown(f"<div class='penjelasan-box'><b>Fakta Nyata:</b> Diagram ieu maparkeun babandingan sacara visual antara nu atos didamelan ('Selesai') sareng nu ngagantung ('Overdue'). Bagian warna beureum/oranye nunjukkeun tumpukan PR (Pekerjaan Rumah) admin klinik anu kedah diberesan.</div>", unsafe_allow_html=True)
            
            st.markdown("---")
            st.markdown("**Statistik 4: Peta Prioritas & Jenis Pasien**")
            cp1, cp2 = st.columns(2)
            with cp1:
                jns_pas = df_an['jenis_pasien'].value_counts().reset_index()
                jns_pas.columns = ['Jenis Pasien', 'Total']
                fig_jp = px.bar(jns_pas, x='Jenis Pasien', y='Total', color='Jenis Pasien', color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_jp.update_layout(showlegend=False, **layout_cfg); st.plotly_chart(fig_jp, use_container_width=True)
            with cp2:
                st.markdown(f"<div class='penjelasan-box' style='margin-top:50px;'><b>Fakta Nyata:</b> Bar chart di gigir misahkeun mana laporan ti pasien Umum, BPJS, sareng Asuransi. Dumasar database, pasien <b>{jns_pas.iloc[0]['Jenis Pasien'] if not jns_pas.empty else '-'}</b> mangrupikeun pelanggan panglobana ({jns_pas.iloc[0]['Total'] if not jns_pas.empty else 0} data).</div>", unsafe_allow_html=True)

        with atab2:
            st.markdown("**Statistik 5: Top 10 Kategori Kasus Terbanyak**")
            df_kat = df_an.copy()
            df_kat['jenis_laporan'] = df_kat['jenis_laporan'].apply(lambda x: 'KOSONG (Belum Diklasifikasi)' if x == '-' else x)
            kat_c = df_kat['jenis_laporan'].value_counts().head(10).reset_index()
            kat_c.columns = ['Kategori', 'Total']
            fig_bar_k = px.bar(kat_c, x='Total', y='Kategori', orientation='h', color_discrete_sequence=['#0ea5e9'])
            fig_bar_k.update_layout(yaxis={'categoryorder':'total ascending'}, **layout_cfg); st.plotly_chart(fig_bar_k, use_container_width=True)
            st.markdown(f"<div class='penjelasan-box'><b>Fakta Nyata:</b> Masalah no. 1 di klinik ieu nyaéta <b>'{kat_c.iloc[0]['Kategori']}'</b> ku total {kat_c.iloc[0]['Total']} laporan. Upami aya tulisan 'KOSONG', éta hartosna admin teu acan ngeusian kolom kategori pas nginput data.</div>", unsafe_allow_html=True)
            
            st.markdown("---")
            cg3, cg4 = st.columns(2)
            with cg3:
                st.markdown("**Statistik 6: Top 5 Petugas/Dokter Tersibuk**")
                df_pet = df_an.copy()
                df_pet['petugas'] = df_pet['petugas'].apply(lambda x: 'Tanpa Petugas' if x == '-' else x)
                pet_c = df_pet['petugas'].value_counts().head(5).reset_index()
                pet_c.columns = ['Nama Petugas', 'Jumlah Ditangani']
                fig_pet = px.bar(pet_c, x='Nama Petugas', y='Jumlah Ditangani', color_discrete_sequence=['#8b5cf6'])
                fig_pet.update_layout(**layout_cfg); st.plotly_chart(fig_pet, use_container_width=True)
                
            with cg4:
                st.markdown(f"<div class='penjelasan-box' style='margin-top:50px;'><b>Fakta Nyata:</b> <b>{pet_c.iloc[0]['Nama Petugas']}</b> mangrupikeun petugas/dokter pangrajinna anu ngadamelan/dicatet dina {pet_c.iloc[0]['Jumlah Ditangani']} laporan. Grafik ieu fungsina salaku ngukur Beban Kerja (Kinerja) SDM di Papan Mas.</div>", unsafe_allow_html=True)
            
            st.markdown("---")
            st.markdown("**Statistik 7: Evaluasi Kelengkapan Data Solusi**")
            df_sol = df_an.copy()
            df_sol['Status Solusi'] = df_sol['solusi'].apply(lambda x: 'Belum Diisi' if x == '-' else 'Sudah Diberi Solusi')
            sol_c = df_sol['Status Solusi'].value_counts().reset_index()
            sol_c.columns = ['Status Solusi', 'Total']
            fig_sol = px.pie(sol_c, values='Total', names='Status Solusi', hole=0.6, color_discrete_sequence=['#94a3b8', '#10b981'])
            fig_sol.update_layout(**layout_cfg); st.plotly_chart(fig_sol, use_container_width=True)
            st.markdown(f"<div class='penjelasan-box'><b>Fakta Nyata:</b> Penting pisan pikeun ngalengkepan kolom Solusi! Tina sakabeh data, aya <b>{len(df_sol[df_sol['Status Solusi'] == 'Belum Diisi'])} laporan</b> anu solusina masih kosong blong.</div>", unsafe_allow_html=True)

        with atab3:
            ct1, ct2 = st.columns(2)
            with ct1:
                st.markdown("**Statistik 8: Daftar 10 Pasien Paling Sering Datang**")
                df_pas = df_an[df_an['nama_pasien'] != '-']
                pas_c = df_pas['nama_pasien'].value_counts().head(10).reset_index()
                pas_c.columns = ['Nama Pasien', 'Total Kunjungan']
                st.dataframe(pas_c, use_container_width=True, hide_index=True)
                st.markdown(f"<div class='penjelasan-box'><b>Fakta Nyata:</b> Sistem nyaring nami pasien. Hasilna, pasien an. <b>{pas_c.iloc[0]['Nama Pasien'] if not pas_c.empty else '-'}</b> kacatet dongkap/ngalapor sababaraha kali (Total {pas_c.iloc[0]['Total Kunjungan'] if not pas_c.empty else 0}x). Ieu tiasa disebat pasien kronis atanapi VIP.</div>", unsafe_allow_html=True)
                
            with ct2:
                st.markdown("**Statistik 9: Daftar Laporan Paling Basi (Overdue Terlama)**")
                df_overdue = df_an[df_an['Smart_Status'] == 'Overdue'].sort_values('waktu_input_dt')[['waktu_input', 'nama_pasien', 'keluhan']].head(10)
                df_overdue.columns = ['Tgl Logbook', 'Pasien', 'Keluhan Belum Selesai']
                st.dataframe(df_overdue, use_container_width=True, hide_index=True)
                try: tgl_lama = df_overdue.iloc[0]['Tgl Logbook'][:10]
                except: tgl_lama = "-"
                st.markdown(f"<div class='penjelasan-box'><b>Fakta Nyata:</b> Ieu daptar 10 laporan 'dosa' admin anu teu acan diberesan. Laporan panglami kacatet asup kaping <b>{tgl_lama}</b> tapi dugi ka ayeuna masih ngagantung.</div>", unsafe_allow_html=True)
            
            st.markdown("---")
            st.markdown("**Statistik 10: Analisis Hari Tersibuk**")
            df_hari = df_an.copy()
            df_hari['Hari'] = df_hari['waktu_input_dt'].dt.day_name().map({'Monday':'Senin', 'Tuesday':'Selasa', 'Wednesday':'Rabu', 'Thursday':'Kamis', 'Friday':'Jumat', 'Saturday':'Sabtu', 'Sunday':'Minggu'})
            hari_c = df_hari['Hari'].value_counts().reindex(['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']).fillna(0).reset_index()
            hari_c.columns = ['Hari', 'Jumlah']
            fig_hari = px.bar(hari_c, x='Hari', y='Jumlah', color_discrete_sequence=['#f43f5e'])
            fig_hari.update_layout(**layout_cfg); st.plotly_chart(fig_hari, use_container_width=True)
            st.markdown(f"<div class='penjelasan-box'><b>Fakta Nyata:</b> Berdasar data historis, poé anu paling seueur narima keluhan pasien ragrag dina poé <b>{df_hari['Hari'].mode()[0] if not df_hari['Hari'].empty else '-'}</b>.</div>", unsafe_allow_html=True)

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
            with dl1: st.download_button("📊 Buat & Download Excel", create_excel(df_siap_ekspor, "Data"), f"{nm_file}.xlsx", use_container_width=True)
            with dl2:
                try: st.download_button("📝 Buat & Download Word", create_word(df_siap_ekspor, "Laporan"), f"{nm_file}.docx", use_container_width=True)
                except: pass
            with dl3:
                try: st.download_button("📄 Buat & Download PDF", create_pdf(df_siap_ekspor, "Laporan"), f"{nm_file}.pdf", mime="application/pdf", use_container_width=True)
                except: pass

        c_ex1, c_ex2 = st.columns(2)
        with c_ex1:
            with st.container(border=True):
                st.markdown("##### 📥 2. Impor Data Massal ke Sistem")
                st.download_button("⬇️ Download Template Excel", create_template_excel(), "Template_Upload.xlsx")
                uploaded_file = st.file_uploader("Upload file Excel yang sudah diisi:", type=['xlsx'])
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
                                    if imp_nama and imp_keluhan:
                                        conn.execute('''INSERT INTO pipp (waktu_input, waktu_selesai, no_bpjs, nama_pasien, no_telp, jenis_laporan, keluhan, solusi, petugas, status) VALUES (?,?,?,?,?,?,?,?,?,?)''', (imp_waktu, imp_waktu, str(row['No. Identitas (BPJS/NIK)']), imp_nama, str(row['No. WhatsApp/HP']), str(row['Kategori']), imp_keluhan, str(row['Tindakan / Solusi']), str(row['Nama Petugas']), str(row['Status'])))
                                conn.commit(); st.cache_data.clear(); st.success("✅ Impor Data Berhasil!"); st.rerun()
                        except Exception as e: st.error(f"Error membaca Excel: {e}")

        with c_ex2:
            with st.container(border=True):
                st.markdown("##### 📤 3. Backup Full Database Sistem")
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
