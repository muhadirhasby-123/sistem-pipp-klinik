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

# ---> PERBAIKAN: FUNGSI PEMBUAT TABEL DITAMBAHKAN KEMBALI <---
def init_db():
    conn = get_conn()
    conn.execute('''CREATE TABLE IF NOT EXISTS pipp (id INTEGER PRIMARY KEY AUTOINCREMENT, waktu_input DATETIME, waktu_selesai DATETIME, no_bpjs TEXT, nama_pasien TEXT, alamat TEXT, no_telp TEXT, jenis_laporan TEXT, keluhan TEXT, solusi TEXT, link_bukti TEXT, is_real INTEGER, contact_id INTEGER, status TEXT DEFAULT 'Baru', priority TEXT DEFAULT 'Normal', short_link TEXT, consent_wa INTEGER DEFAULT 0)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS contacts (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT, whatsapp TEXT UNIQUE, bpjs TEXT, nik TEXT, consent_wa INTEGER DEFAULT 0, first_interaction_at DATETIME, last_interaction_at DATETIME)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS interactions (id INTEGER PRIMARY KEY AUTOINCREMENT, contact_id INTEGER, module TEXT, details TEXT, is_ticketed INTEGER DEFAULT 0, created_at DATETIME)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS ticket_logs (id INTEGER PRIMARY KEY AUTOINCREMENT, ticket_id INTEGER, actor TEXT, action TEXT, note TEXT, is_internal INTEGER DEFAULT 1, created_at DATETIME)''')
    conn.execute('''CREATE TABLE IF NOT EXISTS main_categories (id INTEGER PRIMARY KEY AUTOINCREMENT, nama TEXT, icon TEXT, deskripsi TEXT, accent TEXT, light TEXT, tipe TEXT, urutan INTEGER)''')
    conn.commit()

# Panggil fungsi ini agar tabel selalu dicek/dibuat
init_db()
# -------------------------------------------------------------

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
st.set_page_config(page_title="Dashboard Admin", layout="wide")

st.markdown("""
<style>
    .stApp { background-color: #f8fafc !important; font-family: 'Inter', sans-serif;}
    div[data-testid="stButton"] > button { background-color: #16a34a !important; color: white !important; border-radius: 8px !important; font-weight: bold !important; width: 100% !important;}
</style>
""", unsafe_allow_html=True)

if "admin_logged_in" not in st.session_state: st.session_state.admin_logged_in = False

if not st.session_state.admin_logged_in:
    st.info("🔐 Silakan masukkan PIN untuk mengakses Database Manajemen Klinik.")
    pin_input = st.text_input("PIN Keamanan:", type="password")
    if st.button("Masuk / Login"):
        if pin_input == PIN_ADMIN_RAHASIA: st.session_state.admin_logged_in = True; st.rerun()
        else: st.error("PIN Salah!")
else:
    conn = get_conn()
    c1, c2 = st.columns([4, 1])
    with c1: st.title("📂 Dashboard Staf Klinik")
    with c2:
        if st.button("🚪 Logout"): st.session_state.admin_logged_in = False; st.rerun()

    try:
        total_pending = pd.read_sql_query("SELECT COUNT(*) FROM pipp WHERE status IN ('Baru', 'Ditinjau', 'Diproses', 'Menunggu Pasien')", conn).iloc[0, 0]
        total_all = pd.read_sql_query("SELECT COUNT(*) FROM pipp", conn).iloc[0, 0]
    except:
        total_pending, total_all = 0, 0

    m1, m2, m3 = st.columns(3)
    m1.metric("🔴 Tiket Belum Selesai", total_pending)
    m2.metric("🟢 Total Keseluruhan Laporan", total_all)
    st.markdown("---")

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["🔴 Antrean Tiket", "👥 Kontak & Interaksi", "✍️ Input Manual", "📑 Arsip & Excel", "⚙️ Database Form Builder"])

    with tab1:
        st.subheader("Daftar Tiket Membutuhkan Tindakan")
        try:
            df_pending = pd.read_sql_query("SELECT id, waktu_input, status, priority, jenis_laporan, nama_pasien, keluhan FROM pipp WHERE status NOT IN ('Selesai', 'Ditutup') ORDER BY id DESC", conn)
            if not df_pending.empty:
                st.dataframe(df_pending, use_container_width=True)
                with st.form("form_solusi"):
                    st.markdown("### Update Status Tiket")
                    c_id, c_stat, c_pri = st.columns(3)
                    with c_id: id_laporan = st.selectbox("ID Laporan:", df_pending['id'].tolist())
                    with c_stat: status_baru = st.selectbox("Update Status:", ["Ditinjau", "Diproses", "Menunggu Pasien", "Menunggu Unit Lain", "Selesai", "Ditutup"])
                    with c_pri: priority_baru = st.selectbox("Priority:", ["Normal", "Penting", "Mendesak"])
                    
                    st.markdown("---")
                    jenis_update = st.radio("Jenis Catatan:", ["Catatan Internal (Hanya Admin)", "Pembaruan Pasien (Bisa Dilihat Pasien di Link Status)"])
                    solusi_admin = st.text_area("Tuliskan Catatan/Pembaruan:")
                    
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
                st.success("🎉 Luar biasa! Semua tiket sudah ditangani.")
        except Exception as e:
            st.error("Sistem sedang menginisialisasi database...")

    with tab2:
        st.subheader("👥 Database Kontak Pasien")
        try:
            df_kontak = pd.read_sql_query("SELECT id, name, whatsapp, bpjs, consent_wa, last_interaction_at FROM contacts ORDER BY last_interaction_at DESC", conn)
            st.dataframe(df_kontak, use_container_width=True)
        except: st.info("Belum ada data.")

    with tab3:
        st.subheader("✍️ Input Data Pasien Manual")
        with st.form("form_manual_staf"):
            m_nama = st.text_input("👤 Nama Pasien *")
            col_m1, col_m2 = st.columns(2)
            with col_m1: m_bpjs = st.text_input("🏥 Nomor BPJS (Opsional)")
            with col_m2: m_nik = st.text_input("💳 NIK KTP (Opsional)")
            m_jenis = st.selectbox("Kategori Laporan", ["PIPP BPJS", "Pengaduan Layanan/Fasilitas", "Layanan Medis Biasa"])
            m_keluhan = st.text_area("💬 Detail Keluhan *")
            m_solusi = st.text_area("✅ Solusi yang Diberikan Staf *")
            if st.form_submit_button("Simpan ke Database"):
                if m_nama and m_keluhan and m_solusi:
                    identitas_m_gabungan = f"BPJS: {m_bpjs} / NIK: {m_nik}"
                    waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    conn.execute('''INSERT INTO pipp (waktu_input, waktu_selesai, no_bpjs, nama_pasien, alamat, no_telp, jenis_laporan, keluhan, solusi, link_bukti, is_real, status)
                                    VALUES (?,?,?,?,'-','-',?,?,?,'-',1, 'Selesai')''',
                                 (waktu_sekarang, waktu_sekarang, identitas_m_gabungan, m_nama, m_jenis, m_keluhan, m_solusi))
                    conn.commit(); st.success("✅ Data berhasil dimasukkan ke Database!")
                else: st.error("Nama, Keluhan, dan Solusi wajib diisi.")

    with tab4:
        st.subheader("📑 Arsip Data & Export Excel")
        try:
            df_all = pd.read_sql_query("SELECT * FROM pipp ORDER BY waktu_input DESC", conn)
            if not df_all.empty:
                st.dataframe(df_all.drop(columns=['id', 'is_real', 'link_bukti', 'alamat'], errors='ignore'), use_container_width=True)
                excel_data = create_excel(df_all, "Arsip Pasien")
                st.download_button(label="📥 Download Data Excel", data=excel_data, file_name=f"Arsip_Pasien_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except: st.info("Arsip kosong.")

    with tab5:
        st.subheader("⚙️ Database Tabel")
        st.info("Anda dapat memantau data mentah dari tabel pembangun formulir di sini.")
        try:
            df_main = pd.read_sql_query("SELECT * FROM main_categories", conn)
            st.dataframe(df_main)
        except: pass
