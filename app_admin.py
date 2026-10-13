import streamlit as st
import pandas as pd
import sqlite3
import io
import requests
from datetime import datetime
import plotly.express as px

# =====================================================================
# PENGATURAN & KONEKSI (GUNAKAN DATA RAHASIA ASLI ANDA)
# =====================================================================
PIN_ADMIN_RAHASIA = "123456" # <-- PASTIKAN DIGANTI DENGAN PIN ASLI ANDA
TELEGRAM_BOT_TOKEN = "PASTE_TOKEN_DISINI" # <-- PASTIKAN DIGANTI DENGAN TOKEN BOT ANDA
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
    
    if 'petugas' not in columns: 
        cursor.execute("ALTER TABLE pipp ADD COLUMN petugas TEXT DEFAULT ''")
    if 'subkategori' not in columns: 
        cursor.execute("ALTER TABLE pipp ADD COLUMN subkategori TEXT DEFAULT ''")
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
    if 'Link Lampiran' in export_df.columns: 
        cols.append('Link Lampiran')
    
    cols_exist = [c for c in cols if c in export_df.columns]
    export_df = export_df[cols_exist]
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name=title)
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
st.set_page_config(page_title="Dashboard Admin Papan Mas", layout="wide", initial_sidebar_state="collapsed")

# --- UI/UX CUSTOM CSS ---
st.markdown("""
<style>
    .stApp { background-color: #f8fafc !important; font-family: 'Inter', sans-serif;}
    div[data-testid="metric-container"] {
        background-color: #ffffff; border: 1px solid #e2e8f0; padding: 12px 15px; border-radius: 8px; box-shadow: 0 1px 2px rgba(0,0,0,0.02);
    }
    div[data-testid="stDataFrame"] > div { max-height: 400px !important; overflow-y: auto !important; border-radius: 8px; border: 1px solid #e2e8f0; }
    div[data-testid="stFormSubmitButton"] > button {
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%) !important; color: white !important; border-radius: 6px !important; width: 100% !important;
    }
    .btn-logout > div > button {
        background-color: transparent !important; color: #ef4444 !important; border: 1px solid #ef4444 !important; border-radius: 6px !important; font-weight: 600 !important;
    }
</style>
""", unsafe_allow_html=True)

if "admin_logged_in" not in st.session_state: 
    st.session_state.admin_logged_in = False

if not st.session_state.admin_logged_in:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br><br>", unsafe_allow_html=True)
        st.info("🔐 Silakan masukkan PIN untuk mengakses Database Manajemen Klinik.")
        pin_input = st.text_input("PIN Keamanan:", type="password")
        if st.button("Masuk / Login", type="primary"):
            if pin_input == PIN_ADMIN_RAHASIA: 
                st.session_state.admin_logged_in = True
                st.rerun()
            else: 
                st.error("PIN Salah!")
else:
    conn = get_conn()
    
    # --- HEADER ---
    c_head1, c_head2 = st.columns([6, 1])
    with c_head1: 
        st.markdown("<h3 style='margin-bottom: 0px;'>🏥 Dashboard Layanan Klinik Pratama Papan Mas</h3>", unsafe_allow_html=True)
    with c_head2:
        st.markdown('<div class="btn-logout">', unsafe_allow_html=True)
        if st.button("🚪 Logout", use_container_width=True): 
            st.session_state.admin_logged_in = False
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
        
    st.markdown("<hr style='border-top: 1px solid #e2e8f0; margin: 10px 0px;'>", unsafe_allow_html=True)

    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "🎫 Laporan & Pencocokan", 
        "📝 Audit Logbook", 
        "📊 Analitik", 
        "👥 Interaksi", 
        "✍️ Input Manual", 
        "📑 Arsip", 
        "⚙️ Kategori"
    ])

    # ---------------------------------------------------------
    # TAB 1: MANAJEMEN ANTREAN & DATA HISTORIS
    # ---------------------------------------------------------
    with tab1:
        try:
            df_all = pd.read_sql_query("SELECT * FROM pipp ORDER BY id DESC", conn)
            df_all = df_all.fillna('')
            df_all['waktu_input_dt'] = pd.to_datetime(df_all['waktu_input'], errors='coerce')
            
            # --- PENCARIAN & FILTER ---
            st.markdown("<p style='font-size: 14px; color: #475569; font-weight: 600; margin-bottom: -10px;'>🔍 Filter Data</p>", unsafe_allow_html=True)
            f_col1, f_col2 = st.columns([2, 1])
            with f_col1: 
                search_query = st.text_input("Pencarian:", placeholder="Cari Nama, BPJS, NIK, atau Keluhan untuk dicocokkan...", label_visibility="collapsed")
            with f_col2: 
                tanggal_filter = st.date_input("Rentang Tanggal:", value=(), label_visibility="collapsed")

            df_filtered = df_all.copy()
            if search_query:
                sq = search_query.lower()
                df_filtered = df_filtered[
                    df_filtered['id'].astype(str).str.contains(sq) | 
                    df_filtered['nama_pasien'].str.lower().str.contains(sq) | 
                    df_filtered['no_bpjs'].str.lower().str.contains(sq) | 
                    df_filtered['keluhan'].str.lower().str.contains(sq)
                ]
                
            if len(tanggal_filter) == 2: 
                df_filtered = df_filtered[(df_filtered['waktu_input_dt'].dt.date >= tanggal_filter[0]) & (df_filtered['waktu_input_dt'].dt.date <= tanggal_filter[1])]
            elif len(tanggal_filter) == 1: 
                df_filtered = df_filtered[df_filtered['waktu_input_dt'].dt.date == tanggal_filter[0]]

            # --- TABEL RINGKAS (BERSIH TOTAL) ---
            st.markdown("<br>", unsafe_allow_html=True)
            df_display = df_filtered[['id', 'waktu_input', 'no_bpjs', 'nama_pasien', 'keluhan', 'jenis_laporan', 'status']].copy()
            
            # Mengganti string kosong dengan tanda strip agar bersih seperti buku fisik
            for col in ['jenis_laporan', 'nama_pasien', 'no_bpjs', 'keluhan']:
                df_display[col] = df_display[col].apply(lambda x: '-' if x in ['', '-'] else x)
                
            df_display = df_display.rename(columns={'id': 'ID', 'waktu_input': 'Tanggal Input', 'no_bpjs': 'Identitas (BPJS/NIK)', 'nama_pasien': 'Nama Pasien', 'keluhan': 'Keluhan', 'jenis_laporan': 'Kategori', 'status': 'Status'})
            
            st.dataframe(df_display, use_container_width=True, hide_index=True)
            
            # --- FORM UPDATE & PENCOCOKAN DATA LOGBOOK ---
            st.markdown("<hr style='border-top: 1px solid #e2e8f0; margin: 20px 0px;'>", unsafe_allow_html=True)
            
            sel_col1, sel_col2 = st.columns([1, 4])
            with sel_col1: 
                st.markdown("<p style='font-size: 15px; font-weight: bold; margin-top: 8px;'>📝 Kelola Data (ID):</p>", unsafe_allow_html=True)
            with sel_col2: 
                selected_id = st.selectbox("ID Laporan", [None] + df_filtered['id'].tolist(), format_func=lambda x: f"Pilih ID: {x}" if x else "--- Klik untuk memilih ID Laporan yang akan dikelola ---", label_visibility="collapsed")
            
            if selected_id:
                data_terpilih = df_all[df_all['id'] == selected_id].iloc[0]
                with st.container(border=True):
                    c_info, c_form = st.columns([1, 2], gap="large")
                    
                    with c_info:
                        st.markdown(f"#### Detail Awal Pasien")
                        st.write(f"**ID Laporan:** {selected_id}")
                        st.write(f"**Nama:** {data_terpilih['nama_pasien'] if data_terpilih['nama_pasien'] not in ['', '-'] else '-'}")
                        st.write(f"**Identitas:** {data_terpilih['no_bpjs'] if data_terpilih['no_bpjs'] not in ['', '-'] else '-'}")
                        st.write(f"**Kontak:** {data_terpilih['no_telp'] if data_terpilih['no_telp'] not in ['', '-'] else '-'}")
                        st.write(f"**Keluhan:** {data_terpilih['keluhan'] if data_terpilih['keluhan'] not in ['', '-'] else '-'}")
                    
                    with c_form:
                        with st.form(f"form_update_{selected_id}"):
                            st.markdown("#### Form Kelengkapan Data & Solusi")
                            st.info("💡 Anda dapat melengkapi Identitas dan Nomor WA di bawah ini berdasarkan buku log.")
                            
                            fc1, fc2 = st.columns(2)
                            with fc1:
                                bpjs_baru = st.text_input("No. Identitas (BPJS/NIK)", value="" if data_terpilih['no_bpjs'] == '-' else data_terpilih['no_bpjs'])
                                kat_baru = st.text_input("Kategori Laporan", value="" if data_terpilih['jenis_laporan'] == '-' else data_terpilih['jenis_laporan'])
                                pet_baru = st.text_input("Nama Petugas / Dokter", value="" if data_terpilih['petugas'] == '-' else data_terpilih['petugas'])
                            with fc2:
                                telp_baru = st.text_input("No. WhatsApp / HP", value="" if data_terpilih['no_telp'] == '-' else data_terpilih['no_telp'])
                                subkat_baru = st.text_input("Subkategori", value="" if data_terpilih['subkategori'] == '-' else data_terpilih['subkategori'])
                                
                                list_stat = ["Baru", "Ditinjau", "Diproses", "Menunggu Pasien", "Selesai", "🚩 Perlu Cek Ulang", "Ditutup"]
                                stat_saat_ini = data_terpilih['status']
                                idx_stat = list_stat.index(stat_saat_ini) if stat_saat_ini in list_stat else 0
                                stat_baru = st.selectbox("Status Penanganan", list_stat, index=idx_stat)
                            
                            sol_baru = st.text_area("Tindakan / Solusi yang diberikan", value="" if data_terpilih['solusi'] == '-' else data_terpilih['solusi'], height=70)
                            
                            if st.form_submit_button("Simpan & Update Data"):
                                wkt_skrng = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                wkt_selesai = data_terpilih['waktu_selesai']
                                
                                if stat_baru in ["Selesai", "Ditutup"] and wkt_selesai == '': 
                                    wkt_selesai = wkt_skrng
                                
                                conn.execute("""UPDATE pipp SET 
                                             no_bpjs = ?, no_telp = ?, jenis_laporan = ?, subkategori = ?, petugas = ?, solusi = ?, status = ?, waktu_selesai = ?
                                             WHERE id = ?""", 
                                             (bpjs_baru, telp_baru, kat_baru, subkat_baru, pet_baru, sol_baru, stat_baru, wkt_selesai, selected_id))
                                
                                conn.execute("INSERT INTO ticket_logs (ticket_id, actor, action, note, is_internal, created_at) VALUES (?,?,?,?,?,?)",
                                             (selected_id, "Admin System", f"Update Data (Status: {stat_baru})", "Pembaruan kelengkapan via dashboard", 1, wkt_skrng))
                                conn.commit()
                                st.success("✅ Data berhasil diperbarui!")
                                st.rerun()

        except Exception as e: 
            st.error(f"Sistem membaca database... ({e})")

    # ---------------------------------------------------------
    # TAB 2: AUDIT LOGBOOK
    # ---------------------------------------------------------
    with tab2:
        st.markdown("#### 📝 Mode Audit (Belakang Layar)")
        st.info("Daftar di bawah ini berisi laporan yang **Kategori, Solusi, atau Petugasnya** belum diisi. Gunakan ini untuk melacak pekerjaan admin yang belum selesai.")
        try:
            df_bolong = pd.read_sql_query("SELECT * FROM pipp WHERE (jenis_laporan = '' OR jenis_laporan = '-' OR jenis_laporan IS NULL) OR (solusi = '' OR solusi = '-' OR solusi IS NULL) ORDER BY id DESC", conn)
            df_bolong = df_bolong.fillna('')
            st.markdown(f"**Total laporan yang belum lengkap dicocokkan: {len(df_bolong)} data**")
            st.dataframe(df_bolong[['id', 'waktu_input', 'nama_pasien', 'keluhan']].rename(columns={'waktu_input': 'Tanggal'}), use_container_width=True, hide_index=True)
        except: 
            pass

    # ---------------------------------------------------------
    # TAB 3: ANALITIK (ON/OFF)
    # ---------------------------------------------------------
    with tab3:
        st.markdown("#### 📊 Visualisasi Analitik Klinik")
        tampilkan_grafik = st.toggle("Aktifkan Grafik (ON/OFF)", value=False, help="Nyalakan ini untuk melihat statistik data.")
        
        if tampilkan_grafik:
            try:
                df_stat = pd.read_sql_query("SELECT * FROM pipp", conn).fillna('')
                c_graf1, c_graf2 = st.columns(2)
                
                with c_graf1:
                    st.markdown("**Rasio Status Penanganan**")
                    status_counts = df_stat['status'].value_counts().reset_index()
                    status_counts.columns = ['Status', 'Jumlah']
                    st.plotly_chart(px.pie(status_counts, values='Jumlah', names='Status', hole=0.4), use_container_width=True)
                
                with c_graf2:
                    st.markdown("**Top Kategori Laporan**")
                    df_kat = df_stat[(df_stat['jenis_laporan'] != '') & (df_stat['jenis_laporan'] != '-')]
                    kat_counts = df_kat['jenis_laporan'].value_counts().head(5).reset_index()
                    kat_counts.columns = ['Kategori', 'Jumlah']
                    st.plotly_chart(px.bar(kat_counts, x='Kategori', y='Jumlah', color='Kategori'), use_container_width=True)
            except: 
                st.warning("Pastikan Anda sudah menambahkan plotly di requirements.txt")
        else:
            st.info("💡 Grafik dimatikan untuk menghemat beban sistem.")

    # ---------------------------------------------------------
    # TAB 4 - 7: FITUR LAMA
    # ---------------------------------------------------------
    with tab4:
        try: 
            df_kontak = pd.read_sql_query("SELECT * FROM contacts ORDER BY last_interaction_at DESC", conn)
            st.dataframe(df_kontak, use_container_width=True)
        except: 
            pass

    with tab5:
        with st.form("form_manual_staf"):
            st.info("Form Input Laporan Manual")
            m_nama = st.text_input("👤 Nama Pasien *")
            col_m1, col_m2 = st.columns(2)
            
            with col_m1: 
                m_bpjs = st.text_input("🏥 Nomor BPJS/NIK (Opsional)")
            with col_m2: 
                m_telp = st.text_input("📞 Nomor WA/HP (Opsional)")
                
            m_jenis = st.text_input("Kategori Laporan", placeholder="Contoh: PIPP BPJS")
            m_keluhan = st.text_area("💬 Detail Keluhan *")
            m_solusi = st.text_area("✅ Tindakan/Solusi *")
            
            if st.form_submit_button("Simpan Laporan Baru"):
                if m_nama and m_keluhan and m_solusi:
                    wkt_skrg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    conn.execute('''INSERT INTO pipp (waktu_input, waktu_selesai, no_bpjs, nama_pasien, alamat, no_telp, jenis_laporan, keluhan, solusi, link_bukti, is_real, status)
                                    VALUES (?,?,?,?,'-',?,?,?,?,'-',1, 'Selesai')''',
                                 (wkt_skrg, wkt_skrg, m_bpjs, m_nama, m_telp, m_jenis, m_keluhan, m_solusi))
                    conn.commit()
                    st.success("Data berhasil ditambahkan!")
                else: 
                    st.error("Lengkapi Nama, Keluhan, dan Solusi.")

    with tab6:
        try:
            df_arsip = pd.read_sql_query("SELECT * FROM pipp ORDER BY id DESC", conn)
            st.dataframe(df_arsip, use_container_width=True, hide_index=True)
            st.download_button("📥 Download Excel", create_excel(df_arsip, "Arsip"), f"Arsip_Klinik_{datetime.now().strftime('%Y%m%d')}.xlsx")
        except: 
            pass

    with tab7:
        try: 
            df_main = pd.read_sql_query("SELECT * FROM main_categories", conn)
            st.dataframe(df_main, use_container_width=True)
        except: 
            pass
