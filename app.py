import streamlit as st
import pandas as pd
import sqlite3
import io
import requests
from datetime import datetime

# --- 1. PENGATURAN ADMIN & NOTIFIKASI ---
# Ganti dengan PIN rahasia untuk Mas Hasby & Ibu
PIN_ADMIN_RAHASIA = "123456" 

# Setup Telegram (Nanti kita isi tokennya jika Mas sudah siap)
TELEGRAM_BOT_TOKEN = "8965561257:AAEpgrwaBQLpR114JITZZriF5BZhzZtNVnM" 
TELEGRAM_CHAT_ID = "2141502195" 

def kirim_notif_telegram(pesan_teks):
    """Fungsi mengirim notifikasi teks ke grup Telegram Admin"""
    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
        url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        payload = {"chat_id": TELEGRAM_CHAT_ID, "text": pesan_teks, "parse_mode": "HTML"}
        try:
            requests.post(url, json=payload)
        except Exception as e:
            pass # Abaikan jika gagal agar web pasien tidak error

# --- 2. DATABASE SETUP & MIGRATION ---
def init_db():
    conn = sqlite3.connect('klinik_pipp.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS pipp (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    waktu_input DATETIME,
                    no_bpjs TEXT,
                    nama_pasien TEXT,
                    alamat TEXT,
                    no_telp TEXT,
                    jenis_laporan TEXT,
                    keluhan TEXT,
                    solusi TEXT,
                    foto_bukti BLOB,
                    is_real INTEGER)''')
    
    # Auto-Migration jika pakai database lama
    try:
        c.execute("ALTER TABLE pipp ADD COLUMN alamat TEXT")
        c.execute("ALTER TABLE pipp ADD COLUMN waktu_input DATETIME")
    except sqlite3.OperationalError:
        pass 
        
    conn.commit()
    return conn

# --- 3. EXCEL GENERATOR ---
def create_excel(df):
    output = io.BytesIO()
    export_df = df.copy()
    export_df.insert(0, 'No', range(1, 1 + len(export_df)))
    export_df = export_df[['No', 'waktu_input', 'no_bpjs', 'nama_pasien', 'alamat', 'no_telp', 'jenis_laporan', 'keluhan', 'solusi']]
    export_df.columns = ['No', 'Waktu Masuk', 'No BPJS', 'Nama Pasien', 'Alamat', 'No HP/WA', 'Jenis', 'Keluhan', 'Solusi']

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name='Laporan')
        worksheet = writer.sheets['Laporan']
        from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
        
        header_fill = PatternFill(start_color="2C3E50", end_color="2C3E50", fill_type="solid")
        for cell in worksheet["1:1"]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center")
            cell.fill = header_fill

        thin = Side(style='thin')
        for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for cell in row:
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                
        widths = {'A': 6, 'B': 20, 'C': 18, 'D': 25, 'E': 30, 'F': 15, 'G': 20, 'H': 40, 'I': 40}
        for col, w in widths.items():
            worksheet.column_dimensions[col].width = w
    return output.getvalue()

# --- UI SETUP ---
st.set_page_config(page_title="Pusat Bantuan Klinik", layout="centered", initial_sidebar_state="collapsed")

# Pemilih Mode di atas (Bukan Sidebar agar lebih ramah HP)
mode_aplikasi = st.selectbox("Akses Sistem:", ["Lapor Keluhan (Pasien)", "Login Manajemen (Admin)"])
st.divider()

# =====================================================================
# MODE 1: FORM PASIEN (Sangat Ringan & Ramah HP)
# =====================================================================
if mode_aplikasi == "Lapor Keluhan (Pasien)":
    st.header("Layanan Pengaduan & Informasi")
    st.caption("Manajemen kami siap mendengar dan menindaklanjuti masukan Anda demi pelayanan yang lebih baik.")
    
    with st.form("form_pasien", clear_on_submit=True):
        nama_p = st.text_input("Nama Lengkap Anda *")
        bpjs_p = st.text_input("Nomor BPJS (13 Digit) *")
        telp_p = st.text_input("Nomor HP / WhatsApp *")
        alamat_p = st.text_area("Alamat Lengkap *")
        
        jenis = st.selectbox("Kategori Laporan *", [
            "PIPP (Informasi Medis, Antrean, Obat)", 
            "Pengaduan Layanan (Sikap Petugas, Fasilitas, dll)"
        ])
        
        keluhan_p = st.text_area("Detail Laporan / Keluhan Anda *")
        
        st.write("Lampirkan Foto Bukti (Opsional)")
        foto_p = st.file_uploader("Pilih Foto", type=['jpg', 'jpeg', 'png'])
        
        submit_pasien = st.form_submit_button("Kirim Laporan", use_container_width=True)
        
        if submit_pasien:
            if nama_p and bpjs_p and telp_p and alamat_p and keluhan_p:
                foto_bytes = foto_p.read() if foto_p else None
                waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                conn = init_db()
                conn.execute('''INSERT INTO pipp 
                                (waktu_input, no_bpjs, nama_pasien, alamat, no_telp, jenis_laporan, keluhan, solusi, foto_bukti, is_real) 
                                VALUES (?,?,?,?,?,?,?,'[MENUNGGU TINDAKAN]',?,1)''',
                             (waktu_sekarang, bpjs_p, nama_p, alamat_p, telp_p, jenis, keluhan_p, foto_bytes))
                conn.commit()
                
                # Format Pesan Notifikasi ke HP
                pesan_notif = f"""
🚨 <b>LAPORAN BARU MASUK!</b> 🚨
<b>Waktu:</b> {waktu_sekarang}
<b>Kategori:</b> {jenis}

<b>Nama:</b> {nama_p}
<b>BPJS:</b> {bpjs_p}
<b>Kontak:</b> {telp_p}

<b>Detail Keluhan:</b>
<i>"{keluhan_p}"</i>

Terdapat Foto: {'Ya' if foto_p else 'Tidak'}
Segera cek dashboard admin untuk menindaklanjuti!
"""
                kirim_notif_telegram(pesan_notif)
                
                st.success("✅ Terima kasih! Laporan Anda telah berhasil terkirim ke Manajemen.")
            else:
                st.error("Mohon lengkapi semua bidang yang bertanda Bintang (*).")

# =====================================================================
# MODE 2: DASHBOARD ADMIN KLINIK (Akses Terkunci)
# =====================================================================
elif mode_aplikasi == "Login Manajemen (Admin)":
    
    # SISTEM LOGIN PIN
    if "admin_logged_in" not in st.session_state:
        st.session_state.admin_logged_in = False

    if not st.session_state.admin_logged_in:
        pin_input = st.text_input("Masukkan PIN Keamanan:", type="password")
        if st.button("Masuk Dashboard"):
            if pin_input == PIN_ADMIN_RAHASIA:
                st.session_state.admin_logged_in = True
                st.rerun()
            else:
                st.error("PIN Salah!")
    
    # JIKA PIN BENAR, TAMPILKAN DASHBOARD
    else:
        st.success("Berhasil Login. Akses Diberikan.")
        if st.button("Keluar (Logout)"):
            st.session_state.admin_logged_in = False
            st.rerun()
            
        st.title("📂 Ruang Kendali Manajemen")
        conn = init_db()
        
        tab1, tab2 = st.tabs(["Laporan Masuk (Pending)", "Database Arsip Lengkap"])

        # TAB 1: TINDAK LANJUT
        with tab1:
            st.subheader("Menunggu Tindakan")
            df_pending = pd.read_sql_query("SELECT id, waktu_input, nama_pasien, jenis_laporan, keluhan FROM pipp WHERE solusi = '[MENUNGGU TINDAKAN]'", conn)
            
            if not df_pending.empty:
                st.dataframe(df_pending, use_container_width=True)
                with st.form("form_solusi"):
                    id_laporan = st.selectbox("Selesaikan Laporan ID:", df_pending['id'].tolist())
                    
                    c = conn.cursor()
                    c.execute("SELECT foto_bukti FROM pipp WHERE id=?", (id_laporan,))
                    foto_data = c.fetchone()[0]
                    if foto_data:
                        st.image(foto_data, caption="Foto Bukti Pelanggaran/Laporan", width=300)
                    else:
                        st.info("Tidak ada foto lampiran.")
                    
                    solusi_admin = st.text_area("Tuliskan Solusi / Penanganan yang telah dilakukan:")
                    if st.form_submit_button("Simpan & Tutup Laporan"):
                        conn.execute("UPDATE pipp SET solusi=? WHERE id=?", (solusi_admin, id_laporan))
                        conn.commit()
                        st.success("Laporan berhasil ditangani!")
                        st.rerun()
            else:
                st.info("Semua laporan sudah tertangani.")

        # TAB 2: ARSIP
        with tab2:
            st.subheader("Arsip Data Keseluruhan")
            df_all = pd.read_sql_query("SELECT * FROM pipp ORDER BY waktu_input DESC", conn)
            if not df_all.empty:
                df_display = df_all.drop(columns=['foto_bukti'])
                st.dataframe(df_display, use_container_width=True)
                st.download_button("Download Laporan Excel", create_excel(df_all), "Arsip_Klinik.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
