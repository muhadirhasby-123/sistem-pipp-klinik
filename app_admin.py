import streamlit as st
import pandas as pd
import sqlite3
import io
import requests
import calendar
from datetime import datetime, timedelta, date
import plotly.express as px
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# Import library untuk Word dan PDF
try:
    from docx import Document
    from docx.shared import Pt
    from fpdf import FPDF
except ImportError:
    pass # Akan di-handle oleh st.warning di UI jika belum terinstall

# =====================================================================
# PENGATURAN & KONEKSI
# =====================================================================
PIN_ADMIN_RAHASIA = "123456" # <-- GANTI DENGAN PIN ASLI ANDA
TELEGRAM_BOT_TOKEN = "PASTE_TOKEN_DISINI" # <-- GANTI DENGAN TOKEN ASLI ANDA
APP_URL = "https://pipp-klinik.streamlit.app"
DB_PATH = "klinik_pipp.db"

def get_conn(): return sqlite3.connect(DB_PATH, check_same_thread=False)

def init_db():
    conn = get_conn()
    conn.execute('''CREATE TABLE IF NOT EXISTS pipp (id INTEGER PRIMARY KEY AUTOINCREMENT, waktu_input DATETIME, waktu_selesai DATETIME, no_bpjs TEXT, nama_pasien TEXT, alamat TEXT, no_telp TEXT, jenis_laporan TEXT, keluhan TEXT, solusi TEXT, link_bukti TEXT, is_real INTEGER, contact_id INTEGER, status TEXT DEFAULT 'Baru', priority TEXT DEFAULT 'Normal', short_link TEXT, consent_wa INTEGER DEFAULT 0)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS contacts (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, whatsapp TEXT UNIQUE, bpjs TEXT, nik TEXT, consent_wa INTEGER DEFAULT 0, first_interaction_at DATETIME, last_interaction_at DATETIME)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS interactions (id INTEGER PRIMARY KEY AUTOINCREMENT, contact_id INTEGER, module TEXT, details TEXT, is_ticketed INTEGER DEFAULT 0, created_at DATETIME)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS ticket_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, ticket_id INTEGER, actor TEXT, action TEXT, note TEXT, is_internal INTEGER DEFAULT 1, created_at DATETIME)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS main_categories (id INTEGER PRIMARY KEY AUTOINCREMENT, nama TEXT, icon TEXT, deskripsi TEXT, accent TEXT, light TEXT, tipe TEXT, urutan INTEGER)''')
    conn.commit()

def upgrade_db():
    conn = get_conn()
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
    if str(text).strip() in ['-', '', 'nan', 'Belum Diklasifikasi', 'Belum Ditentukan', '[MENUNGGU TINDAKAN]', 'None']:
        return '-'
    return str(text)

def get_smart_status(status_db, waktu_input_str):
    try:
        tgl_logbook = pd.to_datetime(waktu_input_str)
        batas_24_jam = datetime.now() - timedelta(days=1)
        if status_db == 'Baru':
            if tgl_logbook >= batas_24_jam: return "Baru Masuk"
            else: return "Overdue"
        return status_db
    except:
        return status_db

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
        column_letter = get_column_letter(col_num)
        column_length = max(dataframe[column_title].astype(str).map(len).max(), len(str(column_title)))
        worksheet.column_dimensions[column_letter].width = min(column_length + 2, 45)
    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
        for cell in row:
            cell.border = thin_border; cell.alignment = Alignment(vertical="top", wrap_text=True)

def prepare_export_df(df, cols_to_export):
    df_eks = df.copy()
    for col in df_eks.columns: df_eks[col] = df_eks[col].apply(clean_text)
    mapping = {
        'id': 'ID Laporan', 'waktu_input': 'Tgl Laporan', 'nama_pasien': 'Nama Pasien',
        'no_bpjs': 'No. BPJS/NIK', 'no_telp': 'No. WhatsApp', 'alamat': 'Alamat',
        'jenis_pasien': 'Jenis Pasien', 'keluhan': 'Keluhan/Diagnosa Awal', 
        'jenis_laporan': 'Kategori Utama', 'subkategori': 'Subkategori',
        'tensi': 'Tensi Darah', 'suhu': 'Suhu Badan',
        'solusi': 'Tindakan/Solusi Akhir', 'petugas': 'Petugas/Dokter', 'status': 'Status Akhir'
    }
    df_eks = df_eks.rename(columns=mapping)
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
    df_template.loc[0] = ['2026-12-01 10:00:00', 'Contoh Nama', '000123456', '081234567', 'Contoh masalah...', 'Masalah Aplikasi', 'Sudah dibantu reset', 'Dr. Budi', 'Selesai']
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_template.astype(str).to_excel(writer, index=False, sheet_name='Template_Upload')
        style_excel_sheet(writer.sheets['Template_Upload'], df_template)
    return output.getvalue()

# =====================================================================
# AKSES BUKTI LAMPIRAN VIA TELEGRAM
# =====================================================================
if "file_id" in st.query_params:
    file_id = st.query_params["file_id"]
    st.title("📥 Akses Lampiran Pasien")
    pin = st.text_input("Masukkan PIN Keamanan Admin:", type="password")
    if pin == PIN_ADMIN_RAHASIA:
        with st.spinner("Mengambil file dari Telegram..."):
            file_info_url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getFile?file_id={file_id}"
            res = requests.get(file_info_url).json()
            if res.get("ok"):
                file_path = res["result"]["file_path"]
                download_url = f"https://api.telegram.org/file/bot{TELEGRAM_BOT_TOKEN}/{file_path}"
                file_bytes = requests.get(download_url).content
                ext = file_path.split(".")[-1] if "." in file_path else "file"
                st.download_button(label="⬇️ Download File Lampiran", data=file_bytes, file_name=f"Bukti_Laporan.{ext}")
            else: st.error("Gagal! File tidak ditemukan.")
    elif pin: st.error("PIN Salah!")
    st.stop()

# =====================================================================
# PROGRAM UTAMA ADMIN
# =====================================================================
st.set_page_config(page_title="Dashboard Admin", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    .stApp { background-color: #f1f5f9; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .custom-table { width: 100%; border-collapse: collapse; background-color: white; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    .custom-table th { background-color: #0f172a; color: white; padding: 12px; text-align: left; font-size: 14px; font-weight: 600; }
    .custom-table td { padding: 10px 12px; border-bottom: 1px solid #e2e8f0; font-size: 14px; color: #334155; }
    .custom-table tr:hover { background-color: #f8fafc; }
    
    .badge-barumasuk { background-color: #dcfce7; color: #166534; border: 1px solid #bbf7d0; padding: 4px 8px; border-radius: 12px; font-weight: bold; font-size: 12px; }
    .badge-overdue { background-color: #fee2e2; color: #991b1b; border: 1px solid #fecaca; padding: 4px 8px; border-radius: 12px; font-weight: bold; font-size: 12px; }
    .badge-selesai { background-color: #eff6ff; color: #1e40af; border: 1px solid #bfdbfe; padding: 4px 8px; border-radius: 12px; font-weight: bold; font-size: 12px; }
    .badge-proses { background-color: #fef9c3; color: #854d0e; border: 1px solid #fef08a; padding: 4px 8px; border-radius: 12px; font-weight: bold; font-size: 12px; }
    
    .form-container { background: white; padding: 20px; border-radius: 12px; border-top: 4px solid #0ea5e9; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); margin-top: -10px; }
    div[data-testid="stFormSubmitButton"] > button { background: #0ea5e9 !important; color: white !important; font-weight: bold !important; border-radius: 6px !important; transition: all 0.2s !important; }
    div[data-testid="stFormSubmitButton"] > button:hover { background: #0284c7 !important; transform: scale(1.02); }
    
    /* Tombol Kalender */
    .btn-cal-has-data > button { background-color: #0ea5e9 !important; color: white !important; font-weight: bold !important; border: none !important; }
    .btn-cal-empty > button { background-color: transparent !important; color: #94a3b8 !important; border: 1px dashed #cbd5e1 !important; }
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
    conn = get_conn()
    st.markdown("<h2 style='color:#0f172a; margin-top: -20px;'>🏥 Dashboard Klinik Pratama Papan Mas</h2>", unsafe_allow_html=True)
    
    # === TOTAL 8 TABS (SEMUA FITUR AMAN & LENGKAP) ===
    tab1, tab2, tab3, tab_kalender, tab4, tab5, tab6, tab7 = st.tabs([
        "🚀 Kelola Pasien", 
        "📝 Audit Logbook", 
        "📑 Generator Laporan", 
        "📅 Kalender Laporan", # <--- FITUR BARU SESUAI INSTRUKSI
        "📊 Mega Analitik 1000+", 
        "✍️ Input Data Baru",
        "📂 Manajemen Excel",
        "👥 Interaksi & Kategori"
    ])

    # =========================================================
    # TAB 1: PENCARIAN & UPDATE
    # =========================================================
    with tab1:
        df_all = pd.read_sql_query("SELECT * FROM pipp ORDER BY id DESC", conn).fillna('')
        cari_pasien = st.text_input("🔍 Cari Pasien untuk Diupdate (Ketik Nama, No. BPJS, atau ID):", placeholder="Contoh: Hasnadi atau 00004051...", key="search_main")
        
        if cari_pasien:
            sq = cari_pasien.lower()
            hasil = df_all[df_all['nama_pasien'].str.lower().str.contains(sq) | df_all['no_bpjs'].str.lower().str.contains(sq) | df_all['id'].astype(str).str.contains(sq)]
            
            if not hasil.empty:
                st.success(f"Ditemukan {len(hasil)} pasien. Silakan lengkapi datanya di bawah ini:")
                p = hasil.iloc[0]
                
                st.markdown("<div class='form-container'>", unsafe_allow_html=True)
                with st.form(f"form_update_{p['id']}"):
                    st.markdown(f"**Mengelola Pasien: {p['nama_pasien']} (ID: {p['id']})**")
                    
                    try: tgl_asli = pd.to_datetime(p['waktu_input']).date(); jam_asli = str(p['waktu_input'])[11:19] if len(str(p['waktu_input'])) > 10 else "00:00:00"
                    except: tgl_asli = datetime.now().date(); jam_asli = "00:00:00"
                    
                    c1, c2, c3, c4 = st.columns(4)
                    with c1: n_tgl = st.date_input("Tanggal Logbook", value=tgl_asli)
                    with c2: n_nama = st.text_input("Nama Pasien", value=clean_text(p['nama_pasien']) if clean_text(p['nama_pasien']) != '-' else '')
                    with c3: n_bpjs = st.text_input("No. BPJS / NIK", value=clean_text(p['no_bpjs']) if clean_text(p['no_bpjs']) != '-' else '')
                    with c4: n_wa = st.text_input("No. WhatsApp", value=clean_text(p['no_telp']) if clean_text(p['no_telp']) != '-' else '')
                    
                    c5, c6, c7, c8 = st.columns(4)
                    with c5: n_tensi = st.text_input("Tensi Darah", value=clean_text(p.get('tensi', '')) if clean_text(p.get('tensi', '')) != '-' else '')
                    with c6: n_suhu = st.text_input("Suhu", value=clean_text(p.get('suhu', '')) if clean_text(p.get('suhu', '')) != '-' else '')
                    with c7: n_jns = st.selectbox("Jenis Pasien", ["Umum", "BPJS", "Asuransi Lain", "VIP"], index=["Umum", "BPJS", "Asuransi Lain", "VIP"].index(p.get('jenis_pasien', 'Umum')) if p.get('jenis_pasien', 'Umum') in ["Umum", "BPJS", "Asuransi Lain", "VIP"] else 0)
                    with c8: n_petugas = st.text_input("Petugas", value=clean_text(p['petugas']) if clean_text(p['petugas']) != '-' else '')
                    
                    c9, c10 = st.columns([1, 1])
                    with c9: n_kat = st.text_input("Kategori Masalah", value=clean_text(p['jenis_laporan']) if clean_text(p['jenis_laporan']) != '-' else '')
                    with c10: 
                        stat_list = ["Baru", "Diproses", "Selesai", "Perlu Cek Ulang", "Ditutup"]
                        n_stat = st.selectbox("Status Sistem", stat_list, index=stat_list.index(p['status']) if p['status'] in stat_list else 0)
                    
                    n_solusi = st.text_area("Tindakan / Solusi", value=clean_text(p['solusi']) if clean_text(p['solusi']) != '-' else '', height=68)
                    
                    if st.form_submit_button("✔ Simpan Update"):
                        w_skrg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        w_input_baru = f"{n_tgl.strftime('%Y-%m-%d')} {jam_asli}"
                        w_sls = p['waktu_selesai']
                        if n_stat in ["Selesai", "Ditutup"] and w_sls == '': w_sls = w_skrg
                        conn.execute("""UPDATE pipp SET waktu_input=?, nama_pasien=?, no_bpjs=?, no_telp=?, jenis_pasien=?, tensi=?, suhu=?, jenis_laporan=?, petugas=?, solusi=?, status=?, waktu_selesai=? WHERE id=?""", (w_input_baru, n_nama, n_bpjs, n_wa, n_jns, n_tensi, n_suhu, n_kat, n_petugas, n_solusi, n_stat, w_sls, p['id']))
                        conn.commit(); st.rerun()
                st.markdown("</div><br>", unsafe_allow_html=True)

        st.markdown("**Data Terkini (Preview 15 Teratas):**")
        df_view = df_all.head(15) 
        html_table = "<table class='custom-table'><thead><tr><th>ID</th><th>Tgl Logbook</th><th>No. BPJS/NIK</th><th>Nama Pasien</th><th>Keluhan Awal</th><th>Petugas</th><th>Status</th></tr></thead><tbody>"
        for _, r in df_view.iterrows():
            smart_stt = get_smart_status(r['status'], r['waktu_input'])
            if smart_stt == "Baru Masuk": bdg = "badge-barumasuk"
            elif smart_stt == "Overdue": bdg = "badge-overdue"
            elif smart_stt in ["Selesai", "Ditutup"]: bdg = "badge-selesai"
            else: bdg = "badge-proses"
            html_table += f"<tr><td>{r['id']}</td><td>{str(r['waktu_input'])[:10]}</td><td>{clean_text(r['no_bpjs'])}</td><td>{clean_text(r['nama_pasien'])}</td><td>{clean_text(r['keluhan'])}</td><td>{clean_text(r['petugas'])}</td><td><span class='{bdg}'>{smart_stt}</span></td></tr>"
        html_table += "</tbody></table>"
        st.markdown(html_table, unsafe_allow_html=True)

    # =========================================================
    # TAB 2: AUDIT LOGBOOK 
    # =========================================================
    with tab2:
        st.markdown("<h4 style='color:#0f172a;'>📝 Validasi Logbook 1-Klik</h4>", unsafe_allow_html=True)
        st.info("Cukup centang kotak 'Tandai Valid' di tabel bawah jika data sudah sesuai dengan logbook fisik.")
        df_audit = pd.read_sql_query("SELECT id, waktu_input, no_bpjs, nama_pasien, keluhan, status, is_audited FROM pipp ORDER BY id DESC", conn)
        sdh_cek = len(df_audit[df_audit['is_audited'] == 1])
        st.write(f"📊 **Progress Audit:** {sdh_cek} dari {len(df_audit)} data telah divalidasi.")
        
        df_editor = df_audit.copy()
        df_editor['Tandai Valid'] = df_editor['is_audited'].apply(lambda x: True if x == 1 else False)
        df_editor['Status Pintar'] = df_editor.apply(lambda r: get_smart_status(r['status'], r['waktu_input']), axis=1)
        df_editor = df_editor[['Tandai Valid', 'id', 'waktu_input', 'no_bpjs', 'nama_pasien', 'keluhan', 'Status Pintar']]
        
        edited_df = st.data_editor(df_editor, hide_index=True, use_container_width=True, column_config={"Tandai Valid": st.column_config.CheckboxColumn("✔ Validasi Logbook"), "id": st.column_config.NumberColumn("ID", disabled=True), "waktu_input": st.column_config.TextColumn("Tanggal", disabled=True), "no_bpjs": st.column_config.TextColumn("No. BPJS/NIK", disabled=True), "nama_pasien": st.column_config.TextColumn("Nama", disabled=True), "keluhan": st.column_config.TextColumn("Keluhan", disabled=True), "Status Pintar": st.column_config.TextColumn("Status", disabled=True)})
        if st.button("💾 Simpan Semua Perubahan Centang", type="primary"):
            for index, row in edited_df.iterrows():
                conn.execute("UPDATE pipp SET is_audited = ? WHERE id = ?", (1 if row['Tandai Valid'] else 0, row['id']))
            conn.commit(); st.success("Validasi berhasil disimpan masal!")

    # =========================================================
    # TAB 3: GENERATOR EKSPOR SPESIFIK
    # =========================================================
    with tab3:
        st.markdown("<h4 style='color:#0f172a;'>📑 Generator Laporan Fleksibel</h4>", unsafe_allow_html=True)
        df_eks = pd.read_sql_query("SELECT * FROM pipp ORDER BY id DESC", conn).fillna('')
        df_eks['waktu_input_dt'] = pd.to_datetime(df_eks['waktu_input'], errors='coerce')
        
        f1, f2, f3, f4 = st.columns(4)
        with f1: fil_tgl = st.date_input("🗓️ Rentang Waktu:", value=())
        with f2: fil_stat = st.multiselect("🚦 Filter Status:", options=df_eks['status'].unique())
        with f3: fil_kat = st.multiselect("🩺 Filter Kategori:", options=[k for k in df_eks['jenis_laporan'].unique() if clean_text(k) != '-'])
        with f4: fil_petugas = st.multiselect("👨‍⚕️ Filter Petugas:", options=[p for p in df_eks['petugas'].unique() if clean_text(p) != '-'])
        
        if len(fil_tgl) == 2: df_eks = df_eks[(df_eks['waktu_input_dt'].dt.date >= fil_tgl[0]) & (df_eks['waktu_input_dt'].dt.date <= fil_tgl[1])]
        elif len(fil_tgl) == 1: df_eks = df_eks[df_eks['waktu_input_dt'].dt.date == fil_tgl[0]]
        if fil_stat: df_eks = df_eks[df_eks['status'].isin(fil_stat)]
        if fil_kat: df_eks = df_eks[df_eks['jenis_laporan'].isin(fil_kat)]
        if fil_petugas: df_eks = df_eks[df_eks['petugas'].isin(fil_petugas)]
        
        st.markdown("---")
        semua_k = ['id', 'waktu_input', 'nama_pasien', 'no_bpjs', 'no_telp', 'alamat', 'jenis_pasien', 'keluhan', 'tensi', 'suhu', 'jenis_laporan', 'subkategori', 'solusi', 'petugas', 'status']
        kolom_pilih = st.multiselect("Kolom Laporan:", options=semua_k, default=['waktu_input', 'no_bpjs', 'nama_pasien', 'keluhan', 'solusi', 'status'])
        
        df_siap_ekspor = prepare_export_df(df_eks, kolom_pilih)
        st.write(f"Preview Laporan (Ditemukan **{len(df_siap_ekspor)}** data):")
        st.dataframe(df_siap_ekspor.head(3), use_container_width=True)
        
        dl1, dl2, dl3 = st.columns(3)
        nm_file = f"Laporan_Spesifik_{datetime.now().strftime('%Y%m%d')}"
        with dl1: st.download_button("📊 Download Excel", create_excel(df_siap_ekspor, "Data"), f"{nm_file}.xlsx", type="primary", use_container_width=True)
        with dl2:
            try: st.download_button("📝 Download Word", create_word(df_siap_ekspor, "Laporan"), f"{nm_file}.docx", type="primary", use_container_width=True)
            except: pass
        with dl3:
            try: st.download_button("📄 Download PDF", create_pdf(df_siap_ekspor, "Laporan"), f"{nm_file}.pdf", mime="application/pdf", type="primary", use_container_width=True)
            except: pass

    # =========================================================
    # TAB 4: KALENDER LAPORAN (FITUR BARU INTERAKTIF)
    # =========================================================
    with tab_kalender:
        st.markdown("<h4 style='color:#0f172a;'>📅 Kalender Laporan Interaktif</h4>", unsafe_allow_html=True)
        
        df_cal = pd.read_sql_query("SELECT id, waktu_input, waktu_selesai, no_bpjs, nama_pasien, keluhan, jenis_laporan, subkategori, solusi, petugas, status FROM pipp", conn).fillna('')
        df_cal['waktu_input_dt'] = pd.to_datetime(df_cal['waktu_input'], errors='coerce')
        df_cal['date_only'] = df_cal['waktu_input_dt'].dt.date
        for col in ['jenis_laporan', 'subkategori', 'solusi', 'petugas']: df_cal[col] = df_cal[col].apply(lambda x: str(x).strip() if str(x).strip() not in ['', '-'] else 'Belum Tercatat')

        with st.expander("🔍 Filter Laporan Khusus Kalender", expanded=True):
            f_col1, f_col2, f_col3 = st.columns(3)
            with f_col1:
                f_tgl_rentang = st.date_input("Rentang Tanggal Bebas (Menimpa Klik Harian):", value=(), key="cal_rentang")
                f_kat = st.multiselect("Kategori Laporan:", options=[k for k in df_cal['jenis_laporan'].unique() if k != 'Belum Tercatat'])
                f_sub = st.multiselect("Subkategori:", options=[s for s in df_cal['subkategori'].unique() if s != 'Belum Tercatat'])
            with f_col2:
                f_stat = st.multiselect("Status Penyelesaian:", options=[s for s in df_cal['status'].unique() if s != 'Belum Tercatat'])
                f_petugas = st.multiselect("Petugas/Dokter:", options=[p for p in df_cal['petugas'].unique() if p != 'Belum Tercatat'])
            with f_col3:
                st.markdown("**Filter Kelengkapan Data**")
                f_solusi_kosong = st.checkbox("Tampilkan HANYA Laporan Tanpa Solusi")
                f_kat_kosong = st.checkbox("Tampilkan HANYA Laporan Tanpa Kategori")

        df_filtered_cal = df_cal.copy()
        is_rentang_aktif = False
        if len(f_tgl_rentang) == 2:
            df_filtered_cal = df_filtered_cal[(df_filtered_cal['date_only'] >= f_tgl_rentang[0]) & (df_filtered_cal['date_only'] <= f_tgl_rentang[1])]
            is_rentang_aktif = True
        elif len(f_tgl_rentang) == 1:
            df_filtered_cal = df_filtered_cal[df_filtered_cal['date_only'] == f_tgl_rentang[0]]
            is_rentang_aktif = True

        if f_kat: df_filtered_cal = df_filtered_cal[df_filtered_cal['jenis_laporan'].isin(f_kat)]
        if f_sub: df_filtered_cal = df_filtered_cal[df_filtered_cal['subkategori'].isin(f_sub)]
        if f_stat: df_filtered_cal = df_filtered_cal[df_filtered_cal['status'].isin(f_stat)]
        if f_petugas: df_filtered_cal = df_filtered_cal[df_filtered_cal['petugas'].isin(f_petugas)]
        if f_solusi_kosong: df_filtered_cal = df_filtered_cal[df_filtered_cal['solusi'] == 'Belum Tercatat']
        if f_kat_kosong: df_filtered_cal = df_filtered_cal[df_filtered_cal['jenis_laporan'] == 'Belum Tercatat']

        laporan_per_hari = df_filtered_cal.groupby('date_only').size().to_dict()

        st.markdown("---")
        try: default_date = df_filtered_cal['date_only'].max() if not df_filtered_cal.empty else datetime.now().date()
        except: default_date = datetime.now().date()
        
        c_nav1, c_nav2, c_nav3 = st.columns([1, 2, 1])
        with c_nav1: sel_tahun = st.selectbox("Tahun", range(2020, datetime.now().year + 3), index=(default_date.year - 2020))
        with c_nav2:
            bulan_list = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
            sel_bulan = st.selectbox("Bulan", range(1, 13), format_func=lambda x: bulan_list[x-1], index=(default_date.month - 1))
        
        st.markdown(f"#### Bulan: {bulan_list[sel_bulan-1]} {sel_tahun}")
        days_header = ["Sen", "Sel", "Rab", "Kam", "Jum", "Sab", "Min"]
        cols_hdr = st.columns(7)
        for i, day_name in enumerate(days_header):
            cols_hdr[i].markdown(f"<div style='text-align:center; font-weight:bold; background:#0f172a; color:white; padding:5px; border-radius:5px;'>{day_name}</div>", unsafe_allow_html=True)
        
        cal_matrix = calendar.monthcalendar(sel_tahun, sel_bulan)
        for week in cal_matrix:
            cols_week = st.columns(7)
            for i, day in enumerate(week):
                with cols_week[i]:
                    if day != 0:
                        current_dt = date(sel_tahun, sel_bulan, day)
                        jumlah_lap = laporan_per_hari.get(current_dt, 0)
                        
                        btn_label = f"{day}\n({jumlah_lap})" if jumlah_lap > 0 else f"{day}"
                        btn_class = "btn-cal-has-data" if jumlah_lap > 0 else "btn-cal-empty"
                        
                        st.markdown(f"<div class='{btn_class}'>", unsafe_allow_html=True)
                        if st.button(btn_label, key=f"cal_btn_{current_dt}", use_container_width=True):
                            st.session_state.cal_selected_date = current_dt
                        st.markdown("</div>", unsafe_allow_html=True)
                    else: st.write("")

        st.markdown("---")
        df_table_view = df_filtered_cal.copy()
        
        col_tb1, col_tb2 = st.columns([3, 1])
        with col_tb1:
            if is_rentang_aktif: st.markdown(f"**Menampilkan Filter Rentang Waktu** (Total: {len(df_table_view)} Laporan)")
            elif st.session_state.cal_selected_date:
                df_table_view = df_table_view[df_table_view['date_only'] == st.session_state.cal_selected_date]
                st.markdown(f"**Menampilkan Laporan pada: {st.session_state.cal_selected_date}** (Total: {len(df_table_view)} Laporan)")
            else: st.markdown(f"**Menampilkan Seluruh Laporan Bulan Ini/Filter** (Total: {len(df_table_view)} Laporan)")
                
        with col_tb2:
            if st.button("🔄 Reset / Lihat Semua", use_container_width=True):
                st.session_state.cal_selected_date = None
                st.rerun()

        df_table_view = df_table_view[['id', 'waktu_input', 'nama_pasien', 'keluhan', 'jenis_laporan', 'status', 'solusi', 'petugas']]
        df_table_view = df_table_view.rename(columns={'id':'ID', 'waktu_input':'Tanggal Laporan', 'nama_pasien':'Nama Pasien', 'keluhan':'Keluhan', 'jenis_laporan':'Kategori', 'status':'Status', 'solusi':'Solusi', 'petugas':'Petugas'})
        if df_table_view.empty: st.info("Tidak ada laporan.")
        else: st.dataframe(df_table_view, use_container_width=True, hide_index=True)

    # =========================================================
    # TAB 5: MEGA ANALITIK 1000+
    # =========================================================
    with tab4:
        st.markdown("<h4 style='color:#0f172a;'>📊 Dashboard 10 Metrik Utama Klinik</h4>", unsafe_allow_html=True)
        df_an = pd.read_sql_query("SELECT * FROM pipp", conn).fillna('')
        df_an['waktu_input_dt'] = pd.to_datetime(df_an['waktu_input'], errors='coerce')
        df_an['Bulan_Tahun'] = df_an['waktu_input_dt'].dt.strftime('%b %Y')
        df_an['Smart_Status'] = df_an.apply(lambda row: get_smart_status(row['status'], row['waktu_input']), axis=1)
        
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("1. Total Database", f"{len(df_an)} Pasien")
        m2.metric("2. Fresh (Masuk 24 Jam)", len(df_an[df_an['Smart_Status'] == 'Baru Masuk']), delta="Normal", delta_color="normal")
        m3.metric("3. Kasus Overdue/Terbengkalai", len(df_an[df_an['Smart_Status'] == 'Overdue']), delta="Urgent!", delta_color="inverse")
        m4.metric("4. Total Diselesaikan", len(df_an[df_an['status'].isin(['Selesai', 'Ditutup'])]))
        
        st.markdown("---")
        cg1, cg2 = st.columns(2)
        layout_cfg = dict(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(family="Inter", color="#334155"), margin=dict(t=30, b=10, l=10, r=10))
        
        with cg1:
            st.markdown("**5. Tren Kepadatan Laporan per Bulan**")
            bulan_counts = df_an['Bulan_Tahun'].value_counts().reset_index()
            bulan_counts.columns = ['Bulan', 'Jumlah']
            fig_bulan = px.line(bulan_counts, x='Bulan', y='Jumlah', markers=True, color_discrete_sequence=['#ef4444'])
            fig_bulan.update_layout(**layout_cfg); st.plotly_chart(fig_bulan, use_container_width=True)
            
        with cg2:
            st.markdown("**6. Rasio Kinerja Penanganan**")
            stat_counts = df_an['Smart_Status'].value_counts().reset_index()
            stat_counts.columns = ['Status', 'Jumlah']
            fig_pie = px.pie(stat_counts, values='Jumlah', names='Status', hole=0.4, color_discrete_sequence=['#22c55e', '#ef4444', '#f59e0b', '#3b82f6'])
            fig_pie.update_layout(**layout_cfg); st.plotly_chart(fig_pie, use_container_width=True)
            
        cg3, cg4 = st.columns(2)
        with cg3:
            st.markdown("**7. Top 10 Kategori/Keluhan Terbanyak**")
            df_kat_bersih = df_an[(df_an['jenis_laporan'] != '') & (df_an['jenis_laporan'] != '-')]
            kat_c = df_kat_bersih['jenis_laporan'].value_counts().head(10).reset_index()
            kat_c.columns = ['Kategori', 'Total']
            fig_bar = px.bar(kat_c, x='Total', y='Kategori', orientation='h', color_discrete_sequence=['#0ea5e9'])
            fig_bar.update_layout(**layout_cfg, yaxis={'categoryorder':'total ascending'}); st.plotly_chart(fig_bar, use_container_width=True)
            
        with cg4:
            st.markdown("**8. Top 5 Petugas/Dokter Tersibuk**")
            df_pet = df_an[(df_an['petugas'] != '') & (df_an['petugas'] != '-')]
            pet_c = df_pet['petugas'].value_counts().head(5).reset_index()
            pet_c.columns = ['Nama Petugas', 'Jumlah Ditangani']
            fig_pet = px.bar(pet_c, x='Nama Petugas', y='Jumlah Ditangani', color_discrete_sequence=['#8b5cf6'])
            fig_pet.update_layout(**layout_cfg); st.plotly_chart(fig_pet, use_container_width=True)
            
        ct1, ct2 = st.columns(2)
        with ct1:
            st.markdown("**9. Top 10 Pasien Paling Sering Lapor**")
            df_pas = df_an[(df_an['nama_pasien'] != '') & (df_an['nama_pasien'] != '-')]
            pas_c = df_pas['nama_pasien'].value_counts().head(10).reset_index()
            pas_c.columns = ['Nama Pasien', 'Frekuensi Datang']
            st.dataframe(pas_c, use_container_width=True, hide_index=True)
            
        with ct2:
            st.markdown("**10. Daftar Keluhan Paling Terbengkalai (Overdue)**")
            df_overdue = df_an[df_an['Smart_Status'] == 'Overdue'][['waktu_input', 'nama_pasien', 'keluhan']].head(10)
            df_overdue.columns = ['Tgl Logbook', 'Pasien', 'Keluhan Belum Selesai']
            st.dataframe(df_overdue, use_container_width=True, hide_index=True)

    # =========================================================
    # TAB 6: INPUT MANUAL PASIEN BARU
    # =========================================================
    with tab5:
        with st.form("form_manual_staf"):
            st.info("Input Pasien dari Buku Log (Bisa Set Tanggal Mundur)")
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
                    conn.execute('''INSERT INTO pipp (waktu_input, waktu_selesai, no_bpjs, nama_pasien, alamat, no_telp, jenis_laporan, keluhan, solusi, status, priority, jenis_pasien) VALUES (?,?,?,?,'-',?,?,?,?,'Baru',?,?)''', (wkt_skrg, wkt_skrg, m_bpjs, m_nama, m_telp, m_jenis, m_keluhan, m_solusi, m_prio, m_jns))
                    conn.commit(); st.success("Data berhasil ditambahkan!")
                else: st.error("Nama dan Keluhan wajib diisi.")

    # =========================================================
    # TAB 7: MANAJEMEN EXCEL (EXPORT & IMPORT MASAL)
    # =========================================================
    with tab6:
        st.markdown("#### 📑 Pusat Manajemen Data Excel")
        c_ex1, c_ex2 = st.columns(2)
        with c_ex1:
            st.markdown("##### 📤 Ekspor Keseluruhan Data")
            try:
                df_arsip = pd.read_sql_query("SELECT * FROM pipp ORDER BY id DESC", conn)
                cols_to_export = ['id', 'waktu_input', 'nama_pasien', 'no_bpjs', 'no_telp', 'alamat', 'jenis_pasien', 'keluhan', 'tensi', 'suhu', 'jenis_laporan', 'subkategori', 'solusi', 'petugas', 'status']
                df_arsip_rapi = prepare_export_df(df_arsip, cols_to_export)
                excel_file = create_excel(df_arsip_rapi, "Arsip_Klinik")
                st.download_button("⬇️ Download Excel Sekarang", excel_file, f"Arsip_Klinik.xlsx", type="primary", use_container_width=True)
            except Exception as e: st.error(f"Gagal memproses file: {e}")
            
        with c_ex2:
            st.markdown("##### 📥 Impor (Upload) Data Buku Log")
            template_file = create_template_excel()
            st.download_button("⬇️ Download Template Excel Kosong", template_file, "Template_Upload.xlsx")
            uploaded_file = st.file_uploader("Pilih file Excel yang sudah diisi:", type=['xlsx'])
            if uploaded_file is not None:
                if st.button("🚀 Proses & Masukkan ke Database", use_container_width=True):
                    try:
                        df_import = pd.read_excel(uploaded_file)
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
                            conn.commit(); st.success("✅ Impor Data Berhasil!")
                        else: st.error("Gagal! Gunakan Template yang disediakan.")
                    except Exception as e: st.error(f"Error membaca Excel: {e}")

    # =========================================================
    # TAB 8: KONTAK & KATEGORI
    # =========================================================
    with tab7:
        st.markdown("#### 👥 Database Kontak & Interaksi")
        try: st.dataframe(pd.read_sql_query("SELECT * FROM contacts ORDER BY last_interaction_at DESC", conn), use_container_width=True)
        except: pass
        st.markdown("#### ⚙️ Pengaturan Form Kategori")
        try: st.dataframe(pd.read_sql_query("SELECT * FROM main_categories", conn), use_container_width=True)
        except: pass
