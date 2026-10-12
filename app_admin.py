import streamlit as st
import pandas as pd
import sqlite3
import io
import requests
from datetime import datetime

# =====================================================================
# PENGATURAN & KONEKSI
# =====================================================================
PIN_ADMIN_RAHASIA = "123456"
TELEGRAM_BOT_TOKEN = "PASTE_TOKEN_DISINI"
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

init_db()

def create_excel(df, title):
    output = io.BytesIO()
    export_df = df.copy()
    export_df.insert(0, 'No', range(1, 1 + len(export_df)))
    if 'link_bukti' in export_df.columns:
        export_df['Link Lampiran'] = export_df['link_bukti'].apply(lambda x: f"{APP_URL}/?file_id={x}" if pd.notna(x) and x != "-" else "Tidak ada")
    export_df['Durasi Penanganan'] = export_df.apply(lambda row: str(pd.to_datetime(row['waktu_selesai']) - pd.to_datetime(row['waktu_input'])).split('.')[0] if pd.notna(row['waktu_selesai']) else "Selesai Instan", axis=1)
    
    cols = ['No', 'waktu_input', 'no_bpjs', 'nama_pasien', 'no_telp', 'keluhan', 'solusi', 'Durasi Penanganan']
    if 'Link Lampiran' in export_df.columns: cols.append('Link Lampiran')
    
    export_df = export_df[cols]
    new_names = ['No', 'Waktu Masuk', 'No Identitas (BPJS/NIK)', 'Nama Pasien', 'Kontak', 'Keluhan', 'Solusi', 'Lama Penanganan']
    if 'Link Lampiran' in export_df.columns: new_names.append('Link Download Bukti')
    export_df.columns = new_names
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name=title)
    return output.getvalue()

# =====================================================================
# AKSES BUKTI LAMPIRAN (FILE ID VIA URL)
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

# --- UI/UX CUSTOM CSS ---
st.markdown("""
<style>
    /* Global Styling */
    .stApp { background-color: #f4f6f9 !important; font-family: 'Inter', 'Segoe UI', sans-serif;}
    
    /* Styling Kartu Metrik (Statistik) */
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        padding: 20px 24px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        border-left: 5px solid #0284c7;
    }
    div[data-testid="metric-container"] > div {
        color: #334155 !important;
    }
    
    /* Tombol Utama (Primary) */
    div[data-testid="stFormSubmitButton"] > button {
        background-color: #0284c7 !important;
        color: white !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        border: none !important;
        width: 100% !important;
        transition: all 0.3s ease;
    }
    div[data-testid="stFormSubmitButton"] > button:hover {
        background-color: #0369a1 !important;
        box-shadow: 0 4px 12px rgba(2, 132, 199, 0.3);
    }
    
    /* Tombol Logout (Secondary) */
    .btn-logout > div > button {
        background-color: transparent !important;
        color: #ef4444 !important;
        border: 1px solid #ef4444 !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
    }
    .btn-logout > div > button:hover {
        background-color: #fef2f2 !important;
    }
    
    /* Styling Tab */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #ffffff;
        border-radius: 8px 8px 0px 0px;
        padding: 0px 16px;
        box-shadow: 0 -2px 5px rgba(0,0,0,0.02);
    }
    .stTabs [aria-selected="true"] {
        background-color: #e0f2fe;
        color: #0369a1 !important;
        font-weight: 700;
        border-bottom: 3px solid #0284c7;
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
    
    # --- HEADER DASHBOARD ---
    c_head1, c_head2 = st.columns([5, 1])
    with c_head1: 
        st.markdown("<h2 style='color: #0f172a; margin-bottom: 0px;'>🏥 Dashboard Layanan Klinik Pratama Papan Mas</h2>", unsafe_allow_html=True)
        st.markdown("<p style='color: #64748b;'>Sistem Informasi Manajemen Antrean & Layanan Terpadu</p>", unsafe_allow_html=True)
    with c_head2:
        st.markdown("<div style='margin-top: 15px;'></div>", unsafe_allow_html=True)
        # Menempatkan tombol logout di dalam div custom agar terpengaruh CSS btn-logout
        st.markdown('<div class="btn-logout">', unsafe_allow_html=True)
        if st.button("🚪 Keluar (Logout)", use_container_width=True): 
            st.session_state.admin_logged_in = False
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # --- KARTU STATISTIK (METRICS) ---
    try:
        total_pending = pd.read_sql_query("SELECT COUNT(*) FROM pipp WHERE status IN ('Baru', 'Ditinjau', 'Diproses', 'Menunggu Pasien')", conn).iloc[0, 0]
        total_all = pd.read_sql_query("SELECT COUNT(*) FROM pipp", conn).iloc[0, 0]
    except:
        total_pending, total_all = 0, 0

    m1, m2, m3 = st.columns([1, 1, 2]) # Memberikan sisa ruang agar kartu tidak memanjang ekstrim
    with m1:
        st.metric("🔴 Tiket Belum Selesai", f"{total_pending} Laporan")
    with m2:
        st.metric("🟢 Total Keseluruhan", f"{total_all} Laporan")
    
    st.markdown("<hr style='border-top: 1px solid #e2e8f0; margin: 25px 0px;'>", unsafe_allow_html=True)

    # --- NAVIGASI TAB ---
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🎫 Antrean Tiket", 
        "👥 Kontak & Interaksi", 
        "✍️ Input Manual", 
        "📑 Arsip & Excel", 
        "⚙️ Form Builder"
    ])

    # --- TAB 1: ANTREAN TIKET ---
    with tab1:
        st.markdown("#### 🎫 Daftar Tiket Membutuhkan Tindakan")
        try:
            # Tetap gunakan df_pending original untuk logika pemrosesan
            df_pending = pd.read_sql_query("SELECT id, waktu_input, status, priority, jenis_laporan, nama_pasien, keluhan FROM pipp WHERE status NOT IN ('Selesai', 'Ditutup') ORDER BY id DESC", conn)
            
            if not df_pending.empty:
                # DataFrame Khusus Tampilan (Menyembunyikan nama kolom mentah)
                df_display_pending = df_pending.rename(columns={
                    'id': 'ID Tiket',
                    'waktu_input': 'Tanggal Input',
                    'status': 'Status',
                    'priority': 'Prioritas',
                    'jenis_laporan': 'Kategori Laporan',
                    'nama_pasien': 'Nama Pasien',
                    'keluhan': 'Keluhan'
                })
                # Tampilkan rapi dengan st.dataframe
                st.dataframe(df_display_pending, use_container_width=True, hide_index=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                with st.form("form_solusi"):
                    st.markdown("#### 📝 Update Status Tiket Pasien")
                    c_id, c_stat, c_pri = st.columns(3)
                    with c_id: id_laporan = st.selectbox("Pilih ID Tiket:", df_pending['id'].tolist())
                    with c_stat: status_baru = st.selectbox("Update Status ke:", ["Ditinjau", "Diproses", "Menunggu Pasien", "Menunggu Unit Lain", "Selesai", "Ditutup"])
                    with c_pri: priority_baru = st.selectbox("Tingkat Prioritas:", ["Normal", "Penting", "Mendesak"])
                    
                    st.markdown("---")
                    jenis_update = st.radio("Jenis Catatan Pengerjaan:", ["Catatan Internal (Hanya Admin)", "Pembaruan Pasien (Bisa Dilihat Pasien di Link Status)"])
                    solusi_admin = st.text_area("Tuliskan Catatan/Tindakan yang telah dilakukan:")
                    
                    if st.form_submit_button("Simpan Perubahan"):
                        waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        is_internal = 1 if "Internal" in jenis_update else 0
                        conn.execute("UPDATE pipp SET status=?, priority=? WHERE id=?", (status_baru, priority_baru, id_laporan))
                        if status_baru in ["Selesai", "Ditutup"]:
                            conn.execute("UPDATE pipp SET waktu_selesai=?, solusi=? WHERE id=?", (waktu_sekarang, solusi_admin, id_laporan))
                        conn.execute("INSERT INTO ticket_logs (ticket_id, actor, action, note, is_internal, created_at) VALUES (?,?,?,?,?,?)",
                                     (id_laporan, "Admin", f"Ubah Status ke {status_baru}", solusi_admin, is_internal, waktu_sekarang))
                        conn.commit()
                        st.success("✅ Tiket berhasil diupdate!"); st.rerun()
            else:
                st.success("🎉 Luar biasa! Seluruh antrean tiket sudah tertangani dengan baik.")
        except Exception as e:
            st.error("Sistem sedang memuat tabel dari database...")

    # --- TAB 2: KONTAK & INTERAKSI ---
    with tab2:
        st.markdown("#### 👥 Database Kontak Pasien")
        try:
            df_kontak = pd.read_sql_query("SELECT id, name, whatsapp, bpjs, consent_wa, last_interaction_at FROM contacts ORDER BY last_interaction_at DESC", conn)
            # Rapikan Kolom Tampilan
            df_display_kontak = df_kontak.rename(columns={
                'id': 'ID',
                'name': 'Nama Pasien',
                'whatsapp': 'No. WhatsApp',
                'bpjs': 'Nomor BPJS',
                'consent_wa': 'Izin WA Notif',
                'last_interaction_at': 'Interaksi Terakhir'
            })
            st.dataframe(df_display_kontak, use_container_width=True, hide_index=True)
        except: st.info("Belum ada data kontak terekam.")

    # --- TAB 3: INPUT MANUAL ---
    with tab3:
        st.markdown("#### ✍️ Form Input Data Manual Staf")
        with st.form("form_manual_staf"):
            m_nama = st.text_input("👤 Nama Pasien Lengkap *")
            col_m1, col_m2 = st.columns(2)
            with col_m1: m_bpjs = st.text_input("🏥 Nomor BPJS (Bila Ada)")
            with col_m2: m_nik = st.text_input("💳 NIK KTP (Bila Ada)")
            m_jenis = st.selectbox("Kategori Layanan", ["PIPP BPJS", "Pengaduan Layanan/Fasilitas", "Layanan Medis Biasa"])
            m_keluhan = st.text_area("💬 Detail Keluhan / Permintaan Pasien *")
            m_solusi = st.text_area("✅ Solusi/Tindakan Langsung dari Staf *")
            
            if st.form_submit_button("Simpan ke Database Utama"):
                if m_nama and m_keluhan and m_solusi:
                    identitas_m_gabungan = f"BPJS: {m_bpjs} / NIK: {m_nik}"
                    waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    conn.execute('''INSERT INTO pipp (waktu_input, waktu_selesai, no_bpjs, nama_pasien, alamat, no_telp, jenis_laporan, keluhan, solusi, link_bukti, is_real, status)
                                    VALUES (?,?,?,?,'-','-',?,?,?,'-',1, 'Selesai')''',
                                 (waktu_sekarang, waktu_sekarang, identitas_m_gabungan, m_nama, m_jenis, m_keluhan, m_solusi))
                    conn.commit(); st.success("✅ Data pasien manual berhasil direkam!")
                else: st.error("Mohon lengkapi bagian Nama, Keluhan, dan Solusi (berbintang wajib).")

    # --- TAB 4: ARSIP & EXCEL ---
    with tab4:
        st.markdown("#### 📑 Rekapitulasi Data & Export Excel")
        try:
            df_all = pd.read_sql_query("SELECT * FROM pipp ORDER BY waktu_input DESC", conn)
            if not df_all.empty:
                df_display_arsip = df_all.drop(columns=['id', 'is_real', 'link_bukti', 'alamat'], errors='ignore')
                # Rapikan Kolom Tampilan
                df_display_arsip = df_display_arsip.rename(columns={
                    'waktu_input': 'Waktu Input',
                    'waktu_selesai': 'Waktu Selesai',
                    'no_bpjs': 'No Identitas',
                    'nama_pasien': 'Nama Pasien',
                    'no_telp': 'Kontak',
                    'jenis_laporan': 'Kategori',
                    'keluhan': 'Keluhan',
                    'solusi': 'Tindakan/Solusi',
                    'contact_id': 'ID Kontak',
                    'status': 'Status',
                    'priority': 'Prioritas',
                    'short_link': 'Link Cek',
                    'consent_wa': 'Izin Notif WA'
                })
                st.dataframe(df_display_arsip, use_container_width=True, hide_index=True)
                
                st.markdown("<br>", unsafe_allow_html=True)
                excel_data = create_excel(df_all, "Arsip Pasien")
                st.download_button(
                    label="📥 Export Seluruh Data ke Excel (.xlsx)", 
                    data=excel_data, 
                    file_name=f"Rekap_Layanan_Klinik_{datetime.now().strftime('%Y%m%d')}.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        except: st.info("Arsip database belum memuat riwayat pasien.")

    # --- TAB 5: FORM BUILDER ---
    with tab5:
        st.markdown("#### ⚙️ Database Tabel Form Builder")
        st.info("Pemantauan arsitektur data kategori layanan. Penambahan tabel dilakukan melalui integrasi form admin (saat ini readonly).")
        try:
            df_main = pd.read_sql_query("SELECT * FROM main_categories", conn)
            st.dataframe(df_main.rename(columns={'nama': 'Nama Kategori', 'icon': 'Ikon', 'deskripsi': 'Deskripsi', 'tipe': 'Tipe Kategori', 'urutan': 'Urutan Tampil'}), use_container_width=True, hide_index=True)
        except: pass
