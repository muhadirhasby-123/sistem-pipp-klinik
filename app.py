import streamlit as st
import pandas as pd
import sqlite3
import io
import requests
from datetime import datetime

# --- PENGATURAN ADMIN & NOTIFIKASI ---
PIN_ADMIN_RAHASIA = "123456" 
TELEGRAM_BOT_TOKEN = "PASTE_TOKEN_DISINI" 
TELEGRAM_CHAT_ID = "PASTE_ID_DISINI"   

def kirim_notif_telegram_dengan_file(pesan_teks, uploaded_file):
    """Fungsi mengirim teks DAN meneruskan file ke Telegram"""
    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
        # 1. Kirim Teksnya dulu
        url_text = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        requests.post(url_text, json={"chat_id": TELEGRAM_CHAT_ID, "text": pesan_teks, "parse_mode": "HTML"})
        
        # 2. Jika ada lampiran, kirim sebagai Dokumen ke Telegram
        if uploaded_file is not None:
            url_doc = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
            file_bytes = uploaded_file.getvalue()
            # File dikirim langsung ke server Telegram
            files = {'document': (uploaded_file.name, file_bytes)}
            data = {'chat_id': TELEGRAM_CHAT_ID, 'caption': "Lampiran Bukti dari Pasien"}
            try:
                requests.post(url_doc, data=data, files=files)
            except Exception as e:
                pass 

# --- DATABASE SETUP (VERSI SUPER RINGAN) ---
def init_db():
    conn = sqlite3.connect('klinik_pipp.db')
    c = conn.cursor()
    # Kita hapus foto_bukti BLOB agar database tidak berat!
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
                    is_real INTEGER)''')
    conn.commit()
    return conn

# --- EXCEL GENERATOR (KHUSUS BPJS) ---
def create_excel_bpjs(df):
    output = io.BytesIO()
    export_df = df.copy()
    export_df.insert(0, 'No', range(1, 1 + len(export_df)))
    export_df = export_df[['No', 'waktu_input', 'no_bpjs', 'nama_pasien', 'no_telp', 'keluhan', 'solusi']]
    export_df.columns = ['No', 'Tanggal Laporan', 'No BPJS', 'Nama Pasien', 'No Kontak', 'Keluhan Administratif', 'Solusi / Edukasi yang Diberikan']

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name='Laporan PIPP BPJS')
        worksheet = writer.sheets['Laporan PIPP BPJS']
        from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
        
        header_fill = PatternFill(start_color="1B5E20", end_color="1B5E20", fill_type="solid")
        for cell in worksheet["1:1"]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center")
            cell.fill = header_fill

        thin = Side(style='thin')
        for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for cell in row:
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                
        widths = {'A': 5, 'B': 20, 'C': 18, 'D': 25, 'E': 15, 'F': 45, 'G': 45}
        for col, w in widths.items():
            worksheet.column_dimensions[col].width = w
    return output.getvalue()

# --- UI SETUP ---
st.set_page_config(page_title="Pusat Bantuan Klinik", layout="centered", initial_sidebar_state="collapsed")

mode_aplikasi = st.selectbox("Akses Sistem:", ["Lapor Keluhan (Pasien)", "Login Manajemen (Admin)"])
st.divider()

# =====================================================================
# MODE 1: FORM PASIEN
# =====================================================================
if mode_aplikasi == "Lapor Keluhan (Pasien)":
    st.header("Layanan Pengaduan & Informasi")
    st.caption("Arahkan kamera HP Anda ke Barcode di Meja Pendaftaran untuk membuka form ini.")
    
    with st.form("form_pasien", clear_on_submit=True):
        nama_p = st.text_input("Nama Lengkap Anda *")
        bpjs_p = st.text_input("Nomor BPJS (13 Digit) *")
        telp_p = st.text_input("Nomor HP / WhatsApp *")
        alamat_p = st.text_area("Alamat Lengkap *")
        
        jenis = st.selectbox("Kategori Laporan *", [
            "PIPP BPJS (Informasi Medis, Antrean, Obat)", 
            "Pengaduan Internal (Sikap Petugas, Fasilitas, dll)"
        ])
        
        keluhan_p = st.text_area("Detail Laporan / Keluhan Anda *")
        
        st.write("Lampirkan Foto/Video/Dokumen Bukti (Opsional, Maks 20MB)")
        # SEKARANG BISA TERIMA GAMBAR, VIDEO MP4, DAN PDF
        file_p = st.file_uploader("Pilih File", type=['jpg', 'jpeg', 'png', 'mp4', 'pdf'])
        
        submit_pasien = st.form_submit_button("Kirim Laporan", use_container_width=True)
        
        if submit_pasien:
            if nama_p and bpjs_p and telp_p and alamat_p and keluhan_p:
                waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                conn = init_db()
                # Menyimpan teks saja tanpa file agar database super ringan
                conn.execute('''INSERT INTO pipp 
                                (waktu_input, no_bpjs, nama_pasien, alamat, no_telp, jenis_laporan, keluhan, solusi, is_real) 
                                VALUES (?,?,?,?,?,?,?,?,1)''',
                             (waktu_sekarang, bpjs_p, nama_p, alamat_p, telp_p, jenis, keluhan_p, '[MENUNGGU TINDAKAN]'))
                conn.commit()
                
                pesan_notif = f"🚨 <b>LAPORAN MASUK!</b> 🚨\n<b>Kategori:</b> {jenis}\n<b>Nama:</b> {nama_p}\n<b>Keluhan:</b> {keluhan_p}"
                
                # Kirim ke Telegram beserta lampiran video/fotonya
                kirim_notif_telegram_dengan_file(pesan_notif, file_p)
                
                st.success("✅ Laporan beserta bukti berhasil terkirim. Manajemen akan segera menindaklanjutinya.")
            else:
                st.error("Mohon lengkapi semua bidang yang bertanda Bintang (*).")

# =====================================================================
# MODE 2: DASHBOARD ADMIN KLINIK
# =====================================================================
elif mode_aplikasi == "Login Manajemen (Admin)":
    
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
    
    else:
        if st.button("Keluar (Logout)"):
            st.session_state.admin_logged_in = False
            st.rerun()
            
        st.title("📂 Ruang Kendali Manajemen")
        conn = init_db()
        
        tab1, tab2 = st.tabs(["Menunggu Tindakan", "Database Arsip (Terpisah)"])

        with tab1:
            st.subheader("Daftar Antrean Solusi")
            df_pending = pd.read_sql_query("SELECT id, waktu_input, jenis_laporan, nama_pasien, keluhan FROM pipp WHERE solusi = '[MENUNGGU TINDAKAN]'", conn)
            
            if not df_pending.empty:
                st.dataframe(df_pending, use_container_width=True)
                with st.form("form_solusi"):
                    id_laporan = st.selectbox("Selesaikan Laporan ID:", df_pending['id'].tolist())
                    
                    st.info("💡 Buka aplikasi Telegram Anda untuk melihat kiriman Video/Foto/Dokumen dari pasien ini.")
                    
                    solusi_admin = st.text_area("Tuliskan Penanganan yang telah dilakukan:")
                    if st.form_submit_button("Simpan & Tutup Laporan"):
                        conn.execute("UPDATE pipp SET solusi=? WHERE id=?", (solusi_admin, id_laporan))
                        conn.commit()
                        st.success("Laporan berhasil ditangani!")
                        st.rerun()
            else:
                st.info("Semua laporan sudah tertangani.")

        with tab2:
            st.subheader("Pusat Arsip Data")
            
            # Tambahan pengaman jika database error karena struktur tabel berubah
            try:
                df_all = pd.read_sql_query("SELECT * FROM pipp ORDER BY waktu_input DESC", conn)
            except:
                df_all = pd.DataFrame()
                st.error("Database sedang menyesuaikan struktur baru. Laporan berikutnya akan normal.")
            
            if not df_all.empty:
                # Membelah DataFrame
                df_pipp_bpjs = df_all[df_all['jenis_laporan'].str.contains('PIPP BPJS', na=False)]
                df_internal = df_all[df_all['jenis_laporan'].str.contains('Pengaduan Internal', na=False)]
                
                st.write("---")
                st.markdown("### 🟢 ARSIP 1: Data Khusus PIPP (Untuk Laporan BPJS)")
                st.dataframe(df_pipp_bpjs.drop(columns=['id', 'is_real'], errors='ignore'), use_container_width=True)
                if not df_pipp_bpjs.empty:
                    st.download_button("📥 Download Excel PIPP BPJS (Resmi)", create_excel_bpjs(df_pipp_bpjs), "Laporan_PIPP_BPJS_Klinik.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                st.write("---")
                st.markdown("### 🔴 ARSIP 2: Data Pengaduan Internal (Hanya untuk Manajemen)")
                st.dataframe(df_internal.drop(columns=['id', 'is_real'], errors='ignore'), use_container_width=True)
            else:
                st.info("Belum ada arsip.")
