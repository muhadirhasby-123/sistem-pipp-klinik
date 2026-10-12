import streamlit as st
import pandas as pd
import sqlite3
import io
import requests
from datetime import datetime

# =====================================================================
# PENGATURAN & KONEKSI (GUNAKAN DATA RAHASIA ASLI ANDA)
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
    conn.commit()

init_db()
upgrade_db()

def create_excel(df, title):
    output = io.BytesIO()
    export_df = df.copy()
    export_df.insert(0, 'No', range(1, 1 + len(export_df)))
    if 'link_bukti' in export_df.columns:
        export_df['Link Lampiran'] = export_df['link_bukti'].apply(lambda x: f"{APP_URL}/?file_id={x}" if pd.notna(x) and x != "-" else "Tidak ada")
    export_df['Durasi Penanganan'] = export_df.apply(lambda row: str(pd.to_datetime(row['waktu_selesai']) - pd.to_datetime(row['waktu_input'])).split('.')[0] if pd.notna(row['waktu_selesai']) and row['waktu_selesai'] != '' else "Belum Selesai", axis=1)
    
    cols = ['No', 'waktu_input', 'no_bpjs', 'nama_pasien', 'no_telp', 'jenis_laporan', 'subkategori', 'keluhan', 'solusi', 'petugas', 'status', 'Durasi Penanganan']
    if 'Link Lampiran' in export_df.columns: cols.append('Link Lampiran')
    
    cols_exist = [c for c in cols if c in export_df.columns]
    export_df = export_df[cols_exist]
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name=title)
    return output.getvalue()

# =====================================================================
# AKSES BUKTI LAMPIRAN
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
st.set_page_config(page_title="Dashboard Admin Papan Mas", layout="wide", initial_sidebar_state="collapsed")

# --- UI/UX CUSTOM CSS (Desain Interaktif & Modern) ---
st.markdown("""
<style>
    /* Transisi lembut untuk seluruh elemen */
    * { transition: all 0.3s ease-in-out; }
    
    /* Tabel dengan Batas Tinggi agar tidak memanjang 1000 baris ke bawah */
    div[data-testid="stDataFrame"] > div {
        max-height: 400px !important; 
        overflow-y: auto !important;
        border-radius: 12px;
        border: 1px solid #334155;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
    }
    
    /* Hover effect pada input form */
    div[data-baseweb="input"]:hover, div[data-baseweb="select"]:hover {
        border-color: #0ea5e9 !important;
        box-shadow: 0 0 0 1px #0ea5e9 !important;
    }
    
    /* Tombol Interaktif */
    div[data-testid="stFormSubmitButton"] > button {
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%) !important;
        color: white !important;
        border-radius: 8px !important;
        font-weight: bold !important;
        border: none !important;
        width: 100% !important;
        padding: 10px !important;
    }
    div[data-testid="stFormSubmitButton"] > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 15px rgba(2, 132, 199, 0.4);
    }
    
    /* Kotak Form Detail yang elegan */
    div[data-testid="stForm"] {
        background-color: transparent !important;
        border: 1px solid #334155 !important;
        border-radius: 12px !important;
        padding: 20px !important;
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
        if st.button("Masuk / Login", type="primary", use_container_width=True):
            if pin_input == PIN_ADMIN_RAHASIA: st.session_state.admin_logged_in = True; st.rerun()
            else: st.error("PIN Salah!")
else:
    conn = get_conn()
    
    c_head1, c_head2 = st.columns([6, 1])
    with c_head1: 
        st.markdown("<h3 style='margin-bottom: 0px;'>🏥 Dashboard Layanan Klinik Pratama Papan Mas</h3>", unsafe_allow_html=True)
    with c_head2:
        if st.button("🚪 Logout", use_container_width=True): 
            st.session_state.admin_logged_in = False
            st.rerun()

    st.markdown("<hr style='border-top: 1px solid #334155; margin: 15px 0px;'>", unsafe_allow_html=True)

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🎫 Manajemen Laporan", 
        "👥 Interaksi", 
        "✍️ Input Manual", 
        "📑 Arsip Lengkap", 
        "⚙️ Kategori"
    ])

    with tab1:
        try:
            df_all = pd.read_sql_query("SELECT * FROM pipp ORDER BY id DESC", conn)
            df_all = df_all.fillna('')
            df_all['waktu_input_dt'] = pd.to_datetime(df_all['waktu_input'], errors='coerce')
            
            # --- FILTER EKSKLUSIF & RAMPING (Hanya Nama, BPJS & Tanggal) ---
            st.markdown("<p style='font-size: 15px; font-weight: bold; margin-bottom: 5px;'>🔍 Filter Data</p>", unsafe_allow_html=True)
            f_col1, f_col2 = st.columns([2, 1])
            
            with f_col1:
                search_query = st.text_input("Pencarian:", placeholder="Ketik Nama, Nomor BPJS/NIK, atau ID...", label_visibility="collapsed")
            with f_col2:
                tanggal_filter = st.date_input("Rentang Tanggal:", value=(), label_visibility="collapsed")

            # --- PROSES FILTER DATA ---
            df_filtered = df_all.copy()
            
            if search_query:
                sq = search_query.lower()
                df_filtered = df_filtered[
                    df_filtered['id'].astype(str).str.contains(sq) |
                    df_filtered['nama_pasien'].str.lower().str.contains(sq) |
                    df_filtered['no_bpjs'].str.lower().str.contains(sq)
                ]
            
            if len(tanggal_filter) == 2:
                df_filtered = df_filtered[(df_filtered['waktu_input_dt'].dt.date >= tanggal_filter[0]) & (df_filtered['waktu_input_dt'].dt.date <= tanggal_filter[1])]
            elif len(tanggal_filter) == 1:
                df_filtered = df_filtered[df_filtered['waktu_input_dt'].dt.date == tanggal_filter[0]]
                
            st.markdown("<br>", unsafe_allow_html=True)
            
            # --- TABEL RINGKAS (Tinggi Dikunci, bisa di-scroll internal) ---
            df_display = df_filtered[['id', 'waktu_input', 'no_bpjs', 'nama_pasien', 'keluhan', 'jenis_laporan', 'status']].copy()
            df_display['jenis_laporan'] = df_display['jenis_laporan'].apply(lambda x: 'Belum diklasifikasi' if x in ['', '-'] else x)
            df_display['nama_pasien'] = df_display['nama_pasien'].apply(lambda x: 'Belum tercatat' if x in ['', '-'] else x)
            df_display['no_bpjs'] = df_display['no_bpjs'].apply(lambda x: 'Belum tercatat' if x in ['', '-'] else x)
            
            df_display = df_display.rename(columns={
                'id': 'ID', 'waktu_input': 'Tanggal Input', 'no_bpjs': 'No. Identitas', 'nama_pasien': 'Nama Pasien',
                'keluhan': 'Keluhan', 'jenis_laporan': 'Kategori', 'status': 'Status'
            })
            
            # Dataframe sekarang akan otomatis memunculkan scrollbar jika lebih dari ~400px (diatur via CSS)
            st.dataframe(df_display, use_container_width=True, hide_index=True)
            
            # --- AKSI DETAIL LAPORAN ---
            st.markdown("<hr style='border-top: 1px solid #334155; margin: 25px 0px;'>", unsafe_allow_html=True)
            
            sel_col1, sel_col2 = st.columns([1, 4])
            with sel_col1:
                st.markdown("<p style='font-size: 15px; font-weight: bold; margin-top: 8px;'>📝 Kelola Laporan (ID):</p>", unsafe_allow_html=True)
            with sel_col2:
                selected_id = st.selectbox("ID Laporan", [None] + df_filtered['id'].tolist(), format_func=lambda x: f"Pilih ID: {x}" if x else "--- Klik di sini untuk memilih ID Laporan ---", label_visibility="collapsed")
            
            if selected_id:
                data_terpilih = df_all[df_all['id'] == selected_id].iloc[0]
                
                # Susun Informasi dan Form secara rapi dan profesional
                c_info, c_form = st.columns([1, 1.5], gap="large")
                
                with c_info:
                    st.markdown(f"#### Detail Pasien (ID: {selected_id})")
                    st.markdown(f"""
                    <div style='background-color: transparent; border-left: 4px solid #0ea5e9; padding-left: 15px; margin-top: 15px;'>
                        <p style='margin-bottom: 5px; color: #94a3b8; font-size: 13px;'>NAMA PASIEN</p>
                        <p style='font-weight: bold; font-size: 16px; margin-top: 0px;'>{data_terpilih['nama_pasien'] if data_terpilih['nama_pasien'] not in ['', '-'] else '*Belum tercatat*'}</p>
                        
                        <p style='margin-bottom: 5px; color: #94a3b8; font-size: 13px;'>IDENTITAS (BPJS/NIK)</p>
                        <p style='font-weight: bold; font-size: 16px; margin-top: 0px;'>{data_terpilih['no_bpjs'] if data_terpilih['no_bpjs'] not in ['', '-'] else '*Belum tercatat*'}</p>
                        
                        <p style='margin-bottom: 5px; color: #94a3b8; font-size: 13px;'>KELUHAN AWAL</p>
                        <p style='font-weight: normal; font-size: 15px; margin-top: 0px;'>{data_terpilih['keluhan'] if data_terpilih['keluhan'] not in ['', '-'] else '*Belum tercatat*'}</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with c_form:
                    with st.form(f"form_update_{selected_id}"):
                        st.markdown("#### Form Tindakan")
                        fc1, fc2 = st.columns(2)
                        with fc1:
                            val_kat = "" if data_terpilih['jenis_laporan'] == '-' else data_terpilih['jenis_laporan']
                            kategori_baru = st.text_input("Kategori Utama", value=val_kat)
                            
                            val_petugas = "" if data_terpilih['petugas'] == '-' else data_terpilih['petugas']
                            petugas_baru = st.text_input("Nama Petugas", value=val_petugas)
                        with fc2:
                            val_subkat = "" if data_terpilih['subkategori'] == '-' else data_terpilih['subkategori']
                            subkategori_baru = st.text_input("Subkategori", value=val_subkat)
                            
                            status_saat_ini = data_terpilih['status']
                            list_status = ["Baru", "Ditinjau", "Diproses", "Menunggu Pasien", "Selesai", "Ditutup"]
                            idx_status = list_status.index(status_saat_ini) if status_saat_ini in list_status else 0
                            status_baru = st.selectbox("Status", list_status, index=idx_status)
                        
                        val_solusi = "" if data_terpilih['solusi'] == '-' else data_terpilih['solusi']
                        solusi_baru = st.text_area("Tindakan / Solusi yang diberikan", value=val_solusi, height=100)
                        
                        st.markdown("<br>", unsafe_allow_html=True)
                        if st.form_submit_button("Simpan & Update Laporan"):
                            waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            waktu_selesai = data_terpilih['waktu_selesai']
                            
                            if status_baru in ["Selesai", "Ditutup"] and waktu_selesai == '':
                                waktu_selesai = waktu_sekarang
                            
                            conn.execute("""UPDATE pipp SET 
                                         jenis_laporan = ?, subkategori = ?, petugas = ?, solusi = ?, status = ?, waktu_selesai = ?
                                         WHERE id = ?""", 
                                         (kategori_baru, subkategori_baru, petugas_baru, solusi_baru, status_baru, waktu_selesai, selected_id))
                            
                            conn.execute("INSERT INTO ticket_logs (ticket_id, actor, action, note, is_internal, created_at) VALUES (?,?,?,?,?,?)",
                                         (selected_id, "Admin System", f"Update Status: {status_baru}", "Pembaruan data", 1, waktu_sekarang))
                            conn.commit()
                            st.success("✅ Laporan Berhasil Diperbarui!"); st.rerun()

        except Exception as e:
            st.error(f"Sistem sedang membaca database... ({e})")

    # ---------------------------------------------------------
    # TAB 2 - 5: (TETAP AMAN)
    # ---------------------------------------------------------
    with tab2:
        try:
            df_kontak = pd.read_sql_query("SELECT id, name, whatsapp, bpjs, consent_wa, last_interaction_at FROM contacts ORDER BY last_interaction_at DESC", conn)
            st.dataframe(df_kontak, use_container_width=True)
        except: st.info("Belum ada data kontak terekam.")

    with tab3:
        with st.form("form_manual_staf"):
            m_nama = st.text_input("👤 Nama Pasien Lengkap *")
            col_m1, col_m2 = st.columns(2)
            with col_m1: m_bpjs = st.text_input("🏥 Nomor BPJS (Opsional)")
            with col_m2: m_nik = st.text_input("💳 NIK KTP (Opsional)")
            m_jenis = st.selectbox("Kategori Layanan", ["PIPP BPJS", "Pengaduan Layanan/Fasilitas", "Layanan Medis Biasa"])
            m_keluhan = st.text_area("💬 Detail Keluhan *")
            m_solusi = st.text_area("✅ Solusi/Tindakan *")
            
            if st.form_submit_button("Simpan Laporan Baru"):
                if m_nama and m_keluhan and m_solusi:
                    identitas_m_gabungan = f"BPJS: {m_bpjs} / NIK: {m_nik}"
                    waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    conn.execute('''INSERT INTO pipp (waktu_input, waktu_selesai, no_bpjs, nama_pasien, alamat, no_telp, jenis_laporan, keluhan, solusi, link_bukti, is_real, status)
                                    VALUES (?,?,?,?,'-','-',?,?,?,'-',1, 'Selesai')''',
                                 (waktu_sekarang, waktu_sekarang, identitas_m_gabungan, m_nama, m_jenis, m_keluhan, m_solusi))
                    conn.commit(); st.success("✅ Data berhasil ditambahkan!")
                else: st.error("Lengkapi Nama, Keluhan, dan Solusi.")

    with tab4:
        try:
            df_arsip = pd.read_sql_query("SELECT * FROM pipp ORDER BY waktu_input DESC", conn)
            if not df_arsip.empty:
                st.dataframe(df_arsip.drop(columns=['id', 'is_real', 'link_bukti', 'alamat'], errors='ignore'), use_container_width=True)
                excel_data = create_excel(df_arsip, "Arsip Pasien")
                st.download_button(label="📥 Download Rekap Excel", data=excel_data, 
                                   file_name=f"Arsip_Klinik_{datetime.now().strftime('%Y%m%d')}.xlsx", 
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except: st.info("Belum ada arsip data.")

    with tab5:
        try:
            df_main = pd.read_sql_query("SELECT * FROM main_categories", conn)
            st.dataframe(df_main, use_container_width=True)
        except: pass
