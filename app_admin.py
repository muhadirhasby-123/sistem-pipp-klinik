import streamlit as st
import pandas as pd
import sqlite3
import io
import requests
from datetime import datetime
import plotly.express as px
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

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
    if 'petugas' not in columns: cursor.execute("ALTER TABLE pipp ADD COLUMN petugas TEXT DEFAULT ''")
    if 'subkategori' not in columns: cursor.execute("ALTER TABLE pipp ADD COLUMN subkategori TEXT DEFAULT ''")
    if 'is_audited' not in columns: cursor.execute("ALTER TABLE pipp ADD COLUMN is_audited INTEGER DEFAULT 0")
    conn.commit()

init_db()
upgrade_db()

# =====================================================================
# FUNGSI EXCEL PROFESIONAL (SUDAH DIPERBAIKI DARI ERROR FLOAT)
# =====================================================================
def style_excel_sheet(worksheet, dataframe):
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    for col_num, column_title in enumerate(dataframe.columns, 1):
        cell = worksheet.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
        column_letter = get_column_letter(col_num)
        # Menghindari error float has no len dengan konversi ke str yang aman
        column_length = max(dataframe[column_title].astype(str).apply(lambda x: len(x) if x != 'nan' else 0).max(), len(str(column_title)))
        worksheet.column_dimensions[column_letter].width = min(column_length + 2, 50) # Max lebar 50

    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="top", wrap_text=True)

def create_excel(df, title):
    output = io.BytesIO()
    export_df = df.copy()
    export_df.insert(0, 'No', range(1, 1 + len(export_df)))
    
    if 'link_bukti' in export_df.columns:
        export_df['Link Lampiran'] = export_df['link_bukti'].apply(lambda x: f"{APP_URL}/?file_id={x}" if pd.notna(x) and str(x) not in ["-", "", "nan"] else "Tidak ada")
    
    export_df['Durasi Penanganan'] = export_df.apply(lambda row: str(pd.to_datetime(row['waktu_selesai']) - pd.to_datetime(row['waktu_input'])).split('.')[0] if pd.notna(row['waktu_selesai']) and str(row['waktu_selesai']) != '' else "Belum Selesai", axis=1)
    
    cols = ['No', 'waktu_input', 'no_bpjs', 'nama_pasien', 'no_telp', 'jenis_laporan', 'subkategori', 'keluhan', 'solusi', 'petugas', 'status', 'Durasi Penanganan']
    if 'Link Lampiran' in export_df.columns: cols.append('Link Lampiran')
    
    cols_exist = [c for c in cols if c in export_df.columns]
    export_df = export_df[cols_exist]
    
    header_mapping = {
        'waktu_input': 'Tanggal Input', 'no_bpjs': 'No. Identitas', 'nama_pasien': 'Nama Pasien', 
        'no_telp': 'Kontak (WA)', 'jenis_laporan': 'Kategori Utama', 'subkategori': 'Subkategori',
        'keluhan': 'Detail Keluhan', 'solusi': 'Tindakan/Solusi', 'petugas': 'Nama Petugas', 'status': 'Status Akhir'
    }
    export_df = export_df.rename(columns=header_mapping)
    
    # Konversi semua data ke string untuk menghindari error openpyxl
    export_df = export_df.astype(str)
    export_df = export_df.replace('nan', '')
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name=title)
        workbook = writer.book
        worksheet = writer.sheets[title]
        style_excel_sheet(worksheet, export_df)

    return output.getvalue()

def create_template_excel():
    output = io.BytesIO()
    # Membuat template kosong dengan kolom yang dibutuhkan
    df_template = pd.DataFrame(columns=['Tanggal Laporan', 'Nama Pasien', 'No. Identitas (BPJS/NIK)', 'No. WhatsApp/HP', 'Keluhan Awal', 'Kategori', 'Tindakan / Solusi', 'Nama Petugas', 'Status'])
    
    # Memberi 1 baris contoh panduan
    df_template.loc[0] = ['2026-12-01 10:00:00', 'Contoh Nama', '000123456', '081234567', 'Contoh masalah...', 'Masalah Aplikasi', 'Sudah dibantu reset password', 'Dr. Budi', 'Selesai']
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_template.to_excel(writer, index=False, sheet_name='Template_Upload')
        worksheet = writer.sheets['Template_Upload']
        style_excel_sheet(worksheet, df_template)
        
    return output.getvalue()

# =====================================================================
# PROGRAM UTAMA
# =====================================================================
st.set_page_config(page_title="Dashboard Admin Papan Mas", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    .stApp { background-color: #f8fafc !important; font-family: 'Inter', sans-serif;}
    div[data-testid="stDataFrame"] > div { max-height: 550px !important; overflow-y: auto !important; border-radius: 8px; border: 1px solid #e2e8f0; }
    div[data-testid="stFormSubmitButton"] > button {
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%) !important; color: white !important; border-radius: 6px !important; width: 100% !important; font-weight: bold !important;
    }
</style>
""", unsafe_allow_html=True)

if "admin_logged_in" not in st.session_state: st.session_state.admin_logged_in = False

if not st.session_state.admin_logged_in:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br><br>", unsafe_allow_html=True)
        st.info("🔐 Silakan masukkan PIN untuk mengakses Database Manajemen Klinik.")
        pin_input = st.text_input("PIN Keamanan:", type="password")
        if st.button("Masuk / Login", type="primary"):
            if pin_input == PIN_ADMIN_RAHASIA: st.session_state.admin_logged_in = True; st.rerun()
            else: st.error("PIN Salah!")
else:
    conn = get_conn()
    
    # --- HEADER ---
    c_head1, c_head2 = st.columns([6, 1])
    with c_head1: st.markdown("<h3 style='margin-bottom: 0px;'>🏥 Dashboard Layanan Klinik Pratama Papan Mas</h3>", unsafe_allow_html=True)
    with c_head2:
        st.markdown('<div style="margin-top: 5px;"></div>', unsafe_allow_html=True)
        if st.button("🚪 Logout", use_container_width=True): 
            st.session_state.admin_logged_in = False; st.rerun()
    st.markdown("<hr style='border-top: 1px solid #e2e8f0; margin: 10px 0px;'>", unsafe_allow_html=True)

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "🎫 Laporan & Pencocokan", 
        "📝 Audit Logbook", 
        "📊 Analitik", 
        "✍️ Input Manual", 
        "📑 Manajemen Excel", 
        "⚙️ Interaksi & Kategori"
    ])

    # ---------------------------------------------------------
    # TAB 1: LAYAR TERBAGI DUA
    # ---------------------------------------------------------
    with tab1:
        try:
            df_all = pd.read_sql_query("SELECT * FROM pipp ORDER BY id DESC", conn)
            df_all = df_all.fillna('')
            df_all['waktu_input_dt'] = pd.to_datetime(df_all['waktu_input'], errors='coerce')
            
            col_kiri, col_kanan = st.columns([1.6, 1], gap="large")
            
            with col_kiri:
                st.markdown("<p style='font-size: 15px; font-weight: bold;'>📋 Tabel Pasien (Cari & Filter)</p>", unsafe_allow_html=True)
                
                f_col1, f_col2 = st.columns([2, 1])
                with f_col1: search_query = st.text_input("Pencarian:", placeholder="Cari Nama, NIK, atau Keluhan...", label_visibility="collapsed")
                with f_col2: tanggal_filter = st.date_input("Rentang Tanggal:", value=(), label_visibility="collapsed")

                df_filtered = df_all.copy()
                if search_query:
                    sq = search_query.lower()
                    df_filtered = df_filtered[df_filtered['id'].astype(str).str.contains(sq) | df_filtered['nama_pasien'].str.lower().str.contains(sq) | df_filtered['no_bpjs'].str.lower().str.contains(sq) | df_filtered['keluhan'].str.lower().str.contains(sq)]
                if len(tanggal_filter) == 2: df_filtered = df_filtered[(df_filtered['waktu_input_dt'].dt.date >= tanggal_filter[0]) & (df_filtered['waktu_input_dt'].dt.date <= tanggal_filter[1])]
                elif len(tanggal_filter) == 1: df_filtered = df_filtered[df_filtered['waktu_input_dt'].dt.date == tanggal_filter[0]]

                df_display = df_filtered[['id', 'waktu_input', 'no_bpjs', 'nama_pasien', 'keluhan', 'jenis_laporan', 'status']].copy()
                
                # Pembersihan kata default dari tabel agar bersih
                kata_buang = ['-', '', 'Belum Diklasifikasi', 'Belum Ditentukan', '[MENUNGGU TINDAKAN]']
                for col in ['jenis_laporan', 'nama_pasien', 'no_bpjs', 'keluhan']:
                    df_display[col] = df_display[col].apply(lambda x: '-' if str(x).strip() in kata_buang else x)
                    
                df_display = df_display.rename(columns={'id': 'ID', 'waktu_input': 'Tanggal', 'no_bpjs': 'Identitas', 'nama_pasien': 'Nama', 'keluhan': 'Keluhan', 'jenis_laporan': 'Kategori', 'status': 'Status'})
                st.dataframe(df_display, use_container_width=True, hide_index=True)
            
            with col_kanan:
                st.markdown("<p style='font-size: 15px; font-weight: bold;'>✍️ Form Pencocokan Logbook</p>", unsafe_allow_html=True)
                
                with st.container(border=True):
                    # Dropdown sekarang hanya menampilkan ID dari hasil filter/pencarian di tabel sebelah kiri
                    selected_id = st.selectbox("Masukkan ID Laporan dari tabel di kiri:", [None] + df_filtered['id'].tolist(), format_func=lambda x: f"Mengelola ID: {x}" if x else "Pilih ID di sini...")
                    
                    if selected_id:
                        data_terpilih = df_all[df_all['id'] == selected_id].iloc[0]
                        st.markdown(f"**Pasien:** {data_terpilih['nama_pasien']} | **Keluhan:** {data_terpilih['keluhan']}")
                        
                        with st.form(f"form_update_{selected_id}"):
                            # Pembersihan kata default agar form kosong saat diedit
                            kata_buang_form = ['-', 'Belum Diklasifikasi', 'Belum Ditentukan', '[MENUNGGU TINDAKAN]']
                            val_bpjs = "" if str(data_terpilih['no_bpjs']).strip() in kata_buang_form else data_terpilih['no_bpjs']
                            val_wa = "" if str(data_terpilih['no_telp']).strip() in kata_buang_form else data_terpilih['no_telp']
                            val_kat = "" if str(data_terpilih['jenis_laporan']).strip() in kata_buang_form else data_terpilih['jenis_laporan']
                            val_pet = "" if str(data_terpilih['petugas']).strip() in kata_buang_form else data_terpilih['petugas']
                            val_sol = "" if str(data_terpilih['solusi']).strip() in kata_buang_form else data_terpilih['solusi']
                            
                            c_f1, c_f2 = st.columns(2)
                            with c_f1:
                                bpjs_baru = st.text_input("Identitas (BPJS/NIK)", value=val_bpjs, placeholder="Ketik NIK/BPJS...")
                                kat_baru = st.text_input("Kategori / Masalah", value=val_kat, placeholder="Contoh: Lupa Sandi")
                            with c_f2:
                                telp_baru = st.text_input("No. WhatsApp/HP", value=val_wa, placeholder="08123...")
                                pet_baru = st.text_input("Nama Petugas", value=val_pet, placeholder="Nama dokter/staf")
                            
                            stat_saat_ini = data_terpilih['status']
                            list_status = ["Baru", "Diproses", "Selesai", "🚩 Perlu Cek Ulang", "Ditutup"]
                            idx_stat = list_status.index(stat_saat_ini) if stat_saat_ini in list_status else 0
                            
                            stat_baru = st.selectbox("Status Penanganan", list_status, index=idx_stat)
                            sol_baru = st.text_area("Tindakan / Solusi dari buku log", value=val_sol, placeholder="Ketik solusi yang diberikan di sini...", height=80)
                            
                            if st.form_submit_button("✔ Simpan & Cocokkan"):
                                wkt_skrng = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                wkt_selesai = data_terpilih['waktu_selesai']
                                if stat_baru in ["Selesai", "Ditutup"] and wkt_selesai == '': wkt_selesai = wkt_skrng
                                
                                conn.execute("""UPDATE pipp SET no_bpjs=?, no_telp=?, jenis_laporan=?, petugas=?, solusi=?, status=?, waktu_selesai=? WHERE id=?""", 
                                             (bpjs_baru, telp_baru, kat_baru, pet_baru, sol_baru, stat_baru, wkt_selesai, selected_id))
                                conn.commit(); st.success("Tersimpan!"); st.rerun()
                    else:
                        st.info("👈 Silakan cari nama pasien di tabel sebelah kiri, lalu pilih ID-nya di atas kotak ini untuk mengedit data.")
                        
        except Exception as e: st.error(f"Error: {e}")

    # ---------------------------------------------------------
    # TAB 2: AUDIT LOGBOOK
    # ---------------------------------------------------------
    with tab2:
        st.markdown("#### 📝 Validasi & Pencocokan Logbook Fisik")
        try:
            df_audit = pd.read_sql_query("SELECT id, waktu_input, nama_pasien, keluhan, is_audited FROM pipp ORDER BY id DESC", conn).fillna('')
            total_data = len(df_audit)
            sudah_dicek = len(df_audit[df_audit['is_audited'] == 1])
            belum_dicek = total_data - sudah_dicek
            
            c_met1, c_met2 = st.columns(2)
            c_met1.metric("Total Seluruh Data", total_data)
            c_met2.metric("Sisa Data Belum Divalidasi", belum_dicek, delta=f"{sudah_dicek} Sudah Valid", delta_color="normal")
            
            df_pending_audit = df_audit[df_audit['is_audited'] != 1].copy()
            ca1, ca2 = st.columns([1.5, 1])
            with ca1:
                st.dataframe(df_pending_audit[['id', 'waktu_input', 'nama_pasien', 'keluhan']].rename(columns={'waktu_input': 'Tanggal'}), use_container_width=True, hide_index=True)
            with ca2:
                val_id = st.selectbox("Pilih ID untuk divalidasi:", [None] + df_pending_audit['id'].tolist(), format_func=lambda x: f"Validasi ID: {x}" if x else "--- Pilih ID ---")
                if val_id:
                    pasien_val = df_pending_audit[df_pending_audit['id'] == val_id].iloc[0]
                    st.success(f"**Nama:** {pasien_val['nama_pasien']}")
                    if st.button("✔ Tandai Sudah Sesuai dengan Buku Log", type="primary", use_container_width=True):
                        conn.execute("UPDATE pipp SET is_audited = 1 WHERE id = ?", (val_id,))
                        conn.commit(); st.rerun()
        except: pass

    # ---------------------------------------------------------
    # TAB 3: ANALITIK 
    # ---------------------------------------------------------
    with tab3:
        tampilkan_grafik = st.toggle("Aktifkan Grafik Analitik (ON/OFF)", value=False)
        if tampilkan_grafik:
            try:
                df_stat = pd.read_sql_query("SELECT * FROM pipp", conn).fillna('')
                c_graf1, c_graf2 = st.columns(2)
                layout_style = dict(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(family="Inter", size=13), margin=dict(t=30, b=10, l=10, r=10))
                with c_graf1:
                    with st.container(border=True):
                        st.markdown("**Rasio Status Penanganan**")
                        status_counts = df_stat['status'].value_counts().reset_index()
                        status_counts.columns = ['Status', 'Jumlah']
                        fig_pie = px.pie(status_counts, values='Jumlah', names='Status', hole=0.5, color_discrete_sequence=['#0ea5e9', '#3b82f6', '#f59e0b', '#10b981'])
                        fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                        fig_pie.update_layout(**layout_style, showlegend=False)
                        st.plotly_chart(fig_pie, use_container_width=True)
                with c_graf2:
                    with st.container(border=True):
                        st.markdown("**Top 5 Kategori Laporan**")
                        df_kat = df_stat[(df_stat['jenis_laporan'] != '') & (df_stat['jenis_laporan'] != '-')]
                        kat_counts = df_kat['jenis_laporan'].value_counts().head(5).reset_index()
                        kat_counts.columns = ['Kategori', 'Jumlah']
                        fig_bar = px.bar(kat_counts, x='Jumlah', y='Kategori', orientation='h', color_discrete_sequence=['#0284c7'])
                        fig_bar.update_layout(**layout_style, yaxis={'categoryorder':'total ascending'})
                        st.plotly_chart(fig_bar, use_container_width=True)
            except: st.warning("Pastikan `plotly` terinstal di requirements.txt")

    # ---------------------------------------------------------
    # TAB 4: INPUT MANUAL STAF
    # ---------------------------------------------------------
    with tab4:
        with st.form("form_manual_staf"):
            st.info("Form untuk menambah pasien baru yang belum ada di sistem.")
            m_nama = st.text_input("👤 Nama Pasien *")
            col_m1, col_m2 = st.columns(2)
            with col_m1: m_bpjs = st.text_input("🏥 Nomor BPJS/NIK")
            with col_m2: m_telp = st.text_input("📞 Nomor WA/HP")
            m_jenis = st.text_input("Kategori Laporan", placeholder="Bebas diisi apa saja...")
            m_keluhan = st.text_area("💬 Detail Keluhan *")
            m_solusi = st.text_area("✅ Tindakan/Solusi")
            
            if st.form_submit_button("Simpan Laporan Baru"):
                if m_nama and m_keluhan:
                    wkt_skrg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    conn.execute('''INSERT INTO pipp (waktu_input, waktu_selesai, no_bpjs, nama_pasien, alamat, no_telp, jenis_laporan, keluhan, solusi, link_bukti, is_real, status)
                                    VALUES (?,?,?,?,'-',?,?,?,?,'-',1, 'Selesai')''',
                                 (wkt_skrg, wkt_skrg, m_bpjs, m_nama, m_telp, m_jenis, m_keluhan, m_solusi))
                    conn.commit(); st.success("Data berhasil ditambahkan!")
                else: st.error("Nama dan Keluhan wajib diisi.")

    # ---------------------------------------------------------
    # TAB 5: MANAJEMEN EXCEL PROFESIONAL
    # ---------------------------------------------------------
    with tab5:
        st.markdown("#### 📑 Pusat Manajemen Data Excel")
        c_ex1, c_ex2 = st.columns(2)
        
        with c_ex1:
            st.markdown("##### 📤 Ekspor (Download) Data Sistem")
            st.info("Unduh seluruh laporan pasien ke format Excel dengan tabel rapi untuk keperluan laporan/arsip klinik.")
            try:
                df_arsip = pd.read_sql_query("SELECT * FROM pipp ORDER BY id DESC", conn)
                excel_file = create_excel(df_arsip, "Data_Klinik")
                st.download_button("⬇️ Download Excel Sekarang", excel_file, f"Arsip_Klinik_{datetime.now().strftime('%Y%m%d')}.xlsx", type="primary", use_container_width=True)
            except Exception as e: st.error(f"Gagal memproses file: {e}")
            
        with c_ex2:
            st.markdown("##### 📥 Impor (Upload) Data Buku Log")
            st.warning("Unggah file Excel (format .xlsx) untuk memasukkan data pasien masal ke sistem.")
            
            # Tombol untuk mengunduh template Excel yang sudah di-styling
            template_file = create_template_excel()
            st.download_button("⬇️ Download Template Excel Kosong", template_file, "Template_Upload_Klinik.xlsx", help="Gunakan file ini untuk mengetik data masal sebelum di-upload.")
            
            uploaded_file = st.file_uploader("Pilih file Excel yang sudah diisi dari komputer Anda:", type=['xlsx'])
            
            if uploaded_file is not None:
                if st.button("🚀 Proses & Masukkan ke Database", use_container_width=True):
                    try:
                        df_import = pd.read_excel(uploaded_file)
                        # Validasi apakah menggunakan template yang benar
                        if 'Nama Pasien' in df_import.columns and 'Keluhan Awal' in df_import.columns:
                            for index, row in df_import.iterrows():
                                # Abaikan baris pertama jika itu baris contoh (contoh nama)
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
                        else:
                            st.error("Gagal! Pastikan Anda menggunakan file Template yang disediakan di tombol atas.")
                    except Exception as e:
                        st.error(f"Terjadi kesalahan membaca Excel: {e}")

    # ---------------------------------------------------------
    # TAB 6: FITUR LAMA
    # ---------------------------------------------------------
    with tab6:
        try: st.dataframe(pd.read_sql_query("SELECT * FROM contacts ORDER BY last_interaction_at DESC", conn), use_container_width=True)
        except: pass
        try: st.dataframe(pd.read_sql_query("SELECT * FROM main_categories", conn), use_container_width=True)
        except: pass
