import streamlit as st
import pandas as pd
import sqlite3
import io
import requests
from datetime import datetime
import plotly.express as px
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# Import library untuk Word dan PDF
try:
    from docx import Document
    from fpdf import FPDF
except ImportError:
    pass

# =====================================================================
# PENGATURAN & KONEKSI
# =====================================================================
PIN_ADMIN_RAHASIA = "123456" # <-- GANTI DENGAN PIN ASLI ANDA
TELEGRAM_BOT_TOKEN = "PASTE_TOKEN_DISINI" # <-- GANTI DENGAN TOKEN ASLI ANDA
APP_URL = "https://pipp-klinik.streamlit.app"
DB_PATH = "klinik_pipp.db"

def get_conn(): 
    return sqlite3.connect(DB_PATH, check_same_thread=False)

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

# =====================================================================
# FUNGSI EXPORT (PERBAIKAN MUTLAK ERROR FLOAT EXCEL)
# =====================================================================
def style_excel_sheet(worksheet, dataframe):
    header_fill = PatternFill(start_color="0f172a", end_color="0f172a", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    for col_num, column_title in enumerate(dataframe.columns, 1):
        cell = worksheet.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
        column_letter = get_column_letter(col_num)
        # PERBAIKAN: Memaksa semua nilai menjadi string terlebih dahulu sebelum dihitung len-nya
        # Ini mencegah error 'float has no len()' 100%
        column_length = max(dataframe[column_title].astype(str).map(len).max(), len(str(column_title)))
        worksheet.column_dimensions[column_letter].width = min(column_length + 2, 45)

    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="top", wrap_text=True)

def prepare_export_df(df, cols_to_export):
    df_eks = df.copy()
    for col in df_eks.columns:
        df_eks[col] = df_eks[col].apply(clean_text)
    
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
    # PERBAIKAN: Mengkonversi seluruh DataFrame menjadi tipe string MURNI sebelum dikirim ke openpyxl
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
        hdr_cells[i].text = str(col_name)
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        
    for index, row in export_df.iterrows():
        row_cells = table.add_row().cells
        for i, val in enumerate(row):
            row_cells[i].text = str(val)
            
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()

def create_pdf(df, title):
    export_df = df.copy().astype(str).replace('nan', '')
    pdf = FPDF(orientation='L', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font("helvetica", "B", 14)
    pdf.cell(0, 10, f"Laporan Klinik - {title}", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    
    pdf.set_font("helvetica", "B", 8)
    col_widths = [270 / len(export_df.columns)] * len(export_df.columns)
    
    pdf.set_fill_color(15, 23, 42)
    pdf.set_text_color(255, 255, 255)
    for i, col in enumerate(export_df.columns):
        pdf.cell(col_widths[i], 8, str(col)[:20], border=1, fill=True, align='C')
    pdf.ln()
    
    pdf.set_font("helvetica", "", 8)
    pdf.set_text_color(0, 0, 0)
    for index, row in export_df.iterrows():
        for i, val in enumerate(row):
            pdf.cell(col_widths[i], 8, str(val)[:25], border=1)
        pdf.ln()
        
    return bytes(pdf.output())

def create_template_excel():
    output = io.BytesIO()
    df_template = pd.DataFrame(columns=['Tanggal Laporan', 'Nama Pasien', 'No. Identitas (BPJS/NIK)', 'No. WhatsApp/HP', 'Keluhan Awal', 'Kategori', 'Tindakan / Solusi', 'Nama Petugas', 'Status'])
    df_template.loc[0] = ['2026-12-01 10:00:00', 'Contoh Nama', '000123456', '081234567', 'Contoh masalah...', 'Masalah Aplikasi', 'Sudah dibantu reset', 'Dr. Budi', 'Selesai']
    df_template = df_template.astype(str) # Paksa string
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False, sheet_name='Template_Upload')
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
            else: 
                st.error("Gagal! File tidak ditemukan.")
    elif pin: 
        st.error("PIN Salah!")
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
    
    .badge-baru { background-color: #fef2f2; color: #ef4444; border: 1px solid #fca5a5; padding: 4px 8px; border-radius: 12px; font-weight: bold; font-size: 12px; }
    .badge-selesai { background-color: #f0fdf4; color: #22c55e; border: 1px solid #86efac; padding: 4px 8px; border-radius: 12px; font-weight: bold; font-size: 12px; }
    .badge-proses { background-color: #eff6ff; color: #3b82f6; border: 1px solid #93c5fd; padding: 4px 8px; border-radius: 12px; font-weight: bold; font-size: 12px; }
    
    .form-container { background: white; padding: 20px; border-radius: 12px; border-top: 4px solid #0ea5e9; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); margin-top: -10px; }
    div[data-testid="stFormSubmitButton"] > button { background: #0ea5e9 !important; color: white !important; font-weight: bold !important; border-radius: 6px !important; transition: all 0.2s !important; }
    div[data-testid="stFormSubmitButton"] > button:hover { background: #0284c7 !important; transform: scale(1.02); }
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
    conn = get_conn()
    
    st.markdown("<h2 style='color:#0f172a; margin-top: -20px;'>🏥 Dashboard Klinik Pratama Papan Mas</h2>", unsafe_allow_html=True)
    
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "🚀 Kelola Pasien", 
        "📝 Audit Logbook", 
        "📑 Generator Laporan", 
        "📊 Visual Analitik", 
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
                    
                    c1, c2, c3 = st.columns(3)
                    with c1: n_nama = st.text_input("Nama Pasien", value=clean_text(p['nama_pasien']) if clean_text(p['nama_pasien']) != '-' else '')
                    with c2: n_bpjs = st.text_input("No. BPJS / NIK", value=clean_text(p['no_bpjs']) if clean_text(p['no_bpjs']) != '-' else '')
                    with c3: n_wa = st.text_input("No. WhatsApp/HP", value=clean_text(p['no_telp']) if clean_text(p['no_telp']) != '-' else '')
                    
                    c4, c5, c6 = st.columns(3)
                    with c4: n_tensi = st.text_input("Tensi Darah (mmHg)", value=clean_text(p.get('tensi', '')) if clean_text(p.get('tensi', '')) != '-' else '')
                    with c5: n_suhu = st.text_input("Suhu Badan (°C)", value=clean_text(p.get('suhu', '')) if clean_text(p.get('suhu', '')) != '-' else '')
                    with c6: n_jns = st.selectbox("Jenis Pasien", ["Umum", "BPJS", "Asuransi Lain", "VIP"], index=["Umum", "BPJS", "Asuransi Lain", "VIP"].index(p.get('jenis_pasien', 'Umum')) if p.get('jenis_pasien', 'Umum') in ["Umum", "BPJS", "Asuransi Lain", "VIP"] else 0)
                    
                    c7, c8, c9 = st.columns(3)
                    with c7: n_kat = st.text_input("Kategori", value=clean_text(p['jenis_laporan']) if clean_text(p['jenis_laporan']) != '-' else '')
                    with c8: n_petugas = st.text_input("Petugas/Dokter", value=clean_text(p['petugas']) if clean_text(p['petugas']) != '-' else '')
                    with c9: 
                        stat_list = ["Baru", "Diproses", "Selesai", "Perlu Cek Ulang", "Ditutup"]
                        n_stat = st.selectbox("Status", stat_list, index=stat_list.index(p['status']) if p['status'] in stat_list else 0)
                    
                    n_solusi = st.text_area("Tindakan / Solusi Medis", value=clean_text(p['solusi']) if clean_text(p['solusi']) != '-' else '', height=68)
                    
                    if st.form_submit_button("✔ Simpan Data Pasien"):
                        w_skrg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        w_sls = p['waktu_selesai']
                        if n_stat in ["Selesai", "Ditutup"] and w_sls == '': w_sls = w_skrg
                        
                        conn.execute("""UPDATE pipp SET 
                                     nama_pasien=?, no_bpjs=?, no_telp=?, jenis_pasien=?, tensi=?, suhu=?,
                                     jenis_laporan=?, petugas=?, solusi=?, status=?, waktu_selesai=? WHERE id=?""", 
                                     (n_nama, n_bpjs, n_wa, n_jns, n_tensi, n_suhu, n_kat, n_petugas, n_solusi, n_stat, w_sls, p['id']))
                        conn.commit(); st.rerun()
                st.markdown("</div><br>", unsafe_allow_html=True)

        st.markdown("**Data Terkini (Preview 15 Teratas):**")
        df_view = df_all.head(15) 
        
        html_table = "<table class='custom-table'><thead><tr><th>ID</th><th>Tgl Input</th><th>No. BPJS/NIK</th><th>Nama Pasien</th><th>Keluhan Awal</th><th>Petugas</th><th>Status</th></tr></thead><tbody>"
        for _, r in df_view.iterrows():
            stt = r['status']
            bdg = "badge-selesai" if stt in ["Selesai", "Ditutup"] else ("badge-baru" if stt == "Baru" else "badge-proses")
            html_table += f"<tr><td>{r['id']}</td><td>{str(r['waktu_input'])[:10]}</td><td>{clean_text(r['no_bpjs'])}</td><td>{clean_text(r['nama_pasien'])}</td><td>{clean_text(r['keluhan'])}</td><td>{clean_text(r['petugas'])}</td><td><span class='{bdg}'>{stt}</span></td></tr>"
        html_table += "</tbody></table>"
        st.markdown(html_table, unsafe_allow_html=True)

    # =========================================================
    # TAB 2: AUDIT LOGBOOK (CENTANG LANGSUNG DI TABEL)
    # =========================================================
    with tab2:
        st.markdown("<h4 style='color:#0f172a;'>📝 Validasi Logbook 1-Klik</h4>", unsafe_allow_html=True)
        st.info("Cukup centang kotak 'Tandai Valid' di tabel bawah jika data sudah sesuai dengan logbook fisik.")
        
        df_audit = pd.read_sql_query("SELECT id, waktu_input, no_bpjs, nama_pasien, keluhan, status, is_audited FROM pipp ORDER BY id DESC", conn)
        
        total_data = len(df_audit)
        sdh_cek = len(df_audit[df_audit['is_audited'] == 1])
        st.write(f"📊 **Progress Audit:** {sdh_cek} dari {total_data} data telah divalidasi.")
        
        df_editor = df_audit.copy()
        df_editor['Tandai Valid'] = df_editor['is_audited'].apply(lambda x: True if x == 1 else False)
        df_editor = df_editor[['Tandai Valid', 'id', 'waktu_input', 'no_bpjs', 'nama_pasien', 'keluhan', 'status']]
        
        edited_df = st.data_editor(
            df_editor,
            hide_index=True,
            use_container_width=True,
            column_config={
                "Tandai Valid": st.column_config.CheckboxColumn("✔ Validasi Logbook", help="Centang jika data sudah sesuai dengan buku manual"),
                "id": st.column_config.NumberColumn("ID", disabled=True),
                "waktu_input": st.column_config.TextColumn("Tanggal", disabled=True),
                "no_bpjs": st.column_config.TextColumn("No. BPJS/NIK", disabled=True),
                "nama_pasien": st.column_config.TextColumn("Nama", disabled=True),
                "keluhan": st.column_config.TextColumn("Keluhan", disabled=True),
                "status": st.column_config.TextColumn("Status", disabled=True)
            }
        )
        
        if st.button("💾 Simpan Semua Perubahan Centang", type="primary"):
            for index, row in edited_df.iterrows():
                val_num = 1 if row['Tandai Valid'] else 0
                conn.execute("UPDATE pipp SET is_audited = ? WHERE id = ?", (val_num, row['id']))
            conn.commit()
            st.success("Validasi berhasil disimpan masal!")

    # =========================================================
    # TAB 3: GENERATOR EKSPOR (Filter Melimpah & Fleksibel)
    # =========================================================
    with tab3:
        st.markdown("<h4 style='color:#0f172a;'>📑 Generator Laporan Fleksibel</h4>", unsafe_allow_html=True)
        st.info("Filter data secara spesifik sebelum men-generate laporan.")
        
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
        st.markdown("**Pilih Kolom yang ingin di-print:**")
        semua_k = ['id', 'waktu_input', 'nama_pasien', 'no_bpjs', 'no_telp', 'alamat', 'jenis_pasien', 'keluhan', 'tensi', 'suhu', 'jenis_laporan', 'subkategori', 'solusi', 'petugas', 'status']
        kolom_pilih = st.multiselect("Kolom Laporan:", options=semua_k, default=['waktu_input', 'no_bpjs', 'nama_pasien', 'keluhan', 'solusi', 'status'])
        
        df_siap_ekspor = prepare_export_df(df_eks, kolom_pilih)
        
        st.write(f"Preview Laporan (Ditemukan **{len(df_siap_ekspor)}** data):")
        st.dataframe(df_siap_ekspor.head(3), use_container_width=True)
        
        st.markdown("---")
        dl1, dl2, dl3 = st.columns(3)
        nm_file = f"Laporan_Spesifik_{datetime.now().strftime('%Y%m%d')}"
        with dl1: st.download_button("📊 Export Excel (.xlsx)", create_excel(df_siap_ekspor, "Data"), f"{nm_file}.xlsx", type="primary", use_container_width=True)
        with dl2:
            try: st.download_button("📝 Export Word (.docx)", create_word(df_siap_ekspor, "Laporan"), f"{nm_file}.docx", type="primary", use_container_width=True)
            except: st.error("Install python-docx")
        with dl3:
            try: st.download_button("📄 Export PDF (.pdf)", create_pdf(df_siap_ekspor, "Laporan"), f"{nm_file}.pdf", mime="application/pdf", type="primary", use_container_width=True)
            except: st.error("Install fpdf2")

    # ---------------------------------------------------------
    # TAB 4: ANALITIK (GRAFIK PROFESIONAL)
    # ---------------------------------------------------------
    with tab4:
        tampilkan_grafik = st.toggle("Aktifkan Grafik (ON/OFF)", value=True)
        if tampilkan_grafik:
            try:
                c_graf1, c_graf2 = st.columns(2)
                layout_style = dict(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(family="Inter", color="#334155"), margin=dict(t=40, b=20, l=20, r=20))
                with c_graf1:
                    with st.container(border=True):
                        st.markdown("**Rasio Penyelesaian Laporan**")
                        sc = df_all['status'].value_counts().reset_index()
                        sc.columns = ['Status', 'Jumlah']
                        fig1 = px.pie(sc, values='Jumlah', names='Status', hole=0.5, color_discrete_sequence=['#0ea5e9', '#3b82f6', '#f59e0b', '#10b981'])
                        fig1.update_layout(**layout_style, showlegend=False)
                        fig1.update_traces(textinfo='percent+label')
                        st.plotly_chart(fig1, use_container_width=True)
                with c_graf2:
                    with st.container(border=True):
                        st.markdown("**Kategori Masalah Terbanyak**")
                        dk = df_all[(df_all['jenis_laporan'] != '') & (df_all['jenis_laporan'] != '-')]
                        kc = dk['jenis_laporan'].value_counts().head(5).reset_index()
                        kc.columns = ['Kategori', 'Jumlah']
                        fig2 = px.bar(kc, x='Jumlah', y='Kategori', orientation='h', color_discrete_sequence=['#0284c7'])
                        fig2.update_layout(**layout_style, yaxis={'categoryorder':'total ascending'})
                        st.plotly_chart(fig2, use_container_width=True)
            except: st.warning("Pastikan `plotly` terinstal.")

    # ---------------------------------------------------------
    # TAB 5: INPUT MANUAL PASIEN BARU
    # ---------------------------------------------------------
    with tab5:
        with st.form("form_manual_staf"):
            st.info("Input Pasien Baru (Tidak ada di database lama)")
            m_nama = st.text_input("👤 Nama Pasien *")
            col_m1, col_m2 = st.columns(2)
            with col_m1: 
                m_bpjs = st.text_input("🏥 Nomor BPJS/NIK")
                m_jns = st.selectbox("Jenis Pasien", ["Umum", "BPJS", "Asuransi Lain"])
            with col_m2: 
                m_telp = st.text_input("📞 Nomor WA/HP")
                m_prio = st.selectbox("Prioritas", ["Normal", "Penting", "Mendesak"])
            m_jenis = st.text_input("Kategori Laporan / Masalah")
            m_keluhan = st.text_area("💬 Detail Keluhan *")
            m_solusi = st.text_area("✅ Tindakan/Solusi")
            
            if st.form_submit_button("Simpan Laporan Baru"):
                if m_nama and m_keluhan:
                    wkt_skrg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    conn.execute('''INSERT INTO pipp (waktu_input, waktu_selesai, no_bpjs, nama_pasien, alamat, no_telp, jenis_laporan, keluhan, solusi, status, priority, jenis_pasien)
                                    VALUES (?,?,?,?,'-',?,?,?,?,'Selesai',?,?)''',
                                 (wkt_skrg, wkt_skrg, m_bpjs, m_nama, m_telp, m_jenis, m_keluhan, m_solusi, m_prio, m_jns))
                    conn.commit(); st.success("Data berhasil ditambahkan!")
                else: st.error("Nama dan Keluhan wajib diisi.")

    # ---------------------------------------------------------
    # TAB 6: MANAJEMEN EXCEL (EXPORT & IMPORT MASAL)
    # ---------------------------------------------------------
    with tab6:
        st.markdown("#### 📑 Pusat Manajemen Data Excel")
        c_ex1, c_ex2 = st.columns(2)
        
        with c_ex1:
            st.markdown("##### 📤 Ekspor Keseluruhan Data Sistem")
            st.info("Unduh seluruh laporan pasien ke format Excel dengan tabel rapi.")
            try:
                # PERBAIKAN: Melempar dataframe ke fungsi prepare_export_df agar kolomnya rapi dan sesuai bahasa
                df_arsip = pd.read_sql_query("SELECT * FROM pipp ORDER BY id DESC", conn)
                cols_to_export = ['id', 'waktu_input', 'nama_pasien', 'no_bpjs', 'no_telp', 'alamat', 'jenis_pasien', 'keluhan', 'tensi', 'suhu', 'jenis_laporan', 'subkategori', 'solusi', 'petugas', 'status']
                df_arsip_rapi = prepare_export_df(df_arsip, cols_to_export)
                
                excel_file = create_excel(df_arsip_rapi, "Arsip_Klinik")
                st.download_button("⬇️ Download Excel Sekarang", excel_file, f"Arsip_Klinik_{datetime.now().strftime('%Y%m%d')}.xlsx", type="primary", use_container_width=True)
            except Exception as e: st.error(f"Gagal memproses file: {e}")
            
        with c_ex2:
            st.markdown("##### 📥 Impor (Upload) Data Buku Log")
            st.warning("Unggah file Excel (format .xlsx) untuk memasukkan data pasien masal ke sistem.")
            template_file = create_template_excel()
            st.download_button("⬇️ Download Template Excel Kosong", template_file, "Template_Upload_Klinik.xlsx", help="Gunakan file ini untuk mengetik data masal sebelum di-upload.")
            
            uploaded_file = st.file_uploader("Pilih file Excel yang sudah diisi dari komputer Anda:", type=['xlsx'])
            
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
                                imp_status = str(row['Status']) if pd.notna(row['Status']) else "Selesai"
                                
                                if imp_nama and imp_keluhan:
                                    conn.execute('''INSERT INTO pipp (waktu_input, waktu_selesai, no_bpjs, nama_pasien, no_telp, jenis_laporan, keluhan, solusi, petugas, status)
                                                    VALUES (?,?,?,?,?,?,?,?,?,?)''', 
                                                 (imp_waktu, imp_waktu, imp_bpjs, imp_nama, imp_telp, imp_kat, imp_keluhan, imp_solusi, imp_petugas, imp_status))
                            conn.commit()
                            st.success("✅ Impor Data Berhasil! Silakan cek di Tab Laporan.")
                        else: st.error("Gagal! Pastikan Anda menggunakan file Template yang disediakan di tombol atas.")
                    except Exception as e: st.error(f"Terjadi kesalahan membaca Excel: {e}")

    # ---------------------------------------------------------
    # TAB 7: FITUR LAMA KONTAK & KATEGORI
    # ---------------------------------------------------------
    with tab7:
        st.markdown("#### 👥 Database Kontak & Interaksi")
        try: st.dataframe(pd.read_sql_query("SELECT * FROM contacts ORDER BY last_interaction_at DESC", conn), use_container_width=True)
        except: pass
        
        st.markdown("#### ⚙️ Pengaturan Form Kategori")
        try: st.dataframe(pd.read_sql_query("SELECT * FROM main_categories", conn), use_container_width=True)
        except: pass
