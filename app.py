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
# PENTING: Ganti dengan link Streamlit milik Mas Hasby!
APP_URL = "https://pipp-klinik.streamlit.app" 

# =====================================================================
# FITUR DOWNLOAD BUKTI DARI LINK EXCEL (Sistem Proxy Keamanan Tinggi)
# =====================================================================
if "file_id" in st.query_params:
    file_id = st.query_params["file_id"]
    st.title("📥 Akses Lampiran Pasien")
    st.warning("File ini bersifat rahasia. Silakan masukkan PIN untuk mengunduh.")
    
    pin = st.text_input("Masukkan PIN Keamanan Admin:", type="password")
    if pin == PIN_ADMIN_RAHASIA:
        with st.spinner("Membuka brankas Telegram untuk mengambil file..."):
            # Mengambil jalur file dari Telegram API
            file_info_url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/getFile?file_id={file_id}"
            res = requests.get(file_info_url).json()
            if res.get("ok"):
                file_path = res["result"]["file_path"]
                download_url = f"https://api.telegram.org/file/bot{TELEGRAM_BOT_TOKEN}/{file_path}"
                file_bytes = requests.get(download_url).content
                
                # Cek ekstensi file (apakah mp4, pdf, jpg, dll)
                ext = file_path.split(".")[-1] if "." in file_path else "file"
                st.success("File berhasil ditarik dari server. Silakan klik tombol di bawah.")
                st.download_button(label="⬇️ Download File Lampiran", data=file_bytes, file_name=f"Bukti_Laporan_Klinik.{ext}")
            else:
                st.error("Gagal! File tidak ditemukan, ukuran terlalu besar, atau sudah dihapus dari Telegram.")
    elif pin:
        st.error("PIN Salah!")
        
    st.stop() # Menghentikan web agar tidak menampilkan form pengaduan di halaman ini

# --- FUNGSI TELEGRAM (UPLOAD FILE & MINTA RESI) ---
def kirim_notif_telegram_dengan_file(pesan_teks, uploaded_file):
    file_id = None
    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
        url_text = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        requests.post(url_text, json={"chat_id": TELEGRAM_CHAT_ID, "text": pesan_teks, "parse_mode": "HTML"})
        
        if uploaded_file is not None:
            url_doc = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
            file_bytes = uploaded_file.getvalue()
            files = {'document': (uploaded_file.name, file_bytes)}
            data = {'chat_id': TELEGRAM_CHAT_ID, 'caption': "Lampiran Bukti dari Pasien"}
            try:
                res = requests.post(url_doc, data=data, files=files).json()
                if res.get("ok"):
                    # Simpan ID resi file dari Telegram
                    file_id = res["result"]["document"]["file_id"]
            except Exception as e:
                pass 
    return file_id

# --- DATABASE SETUP ---
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
                    link_bukti TEXT,
                    is_real INTEGER)''')
    
    # Update kolom baru untuk versi ini
    try:
        c.execute("ALTER TABLE pipp ADD COLUMN link_bukti TEXT")
    except sqlite3.OperationalError:
        pass 
        
    conn.commit()
    return conn

# --- EXCEL GENERATOR ---
def create_excel(df, title):
    output = io.BytesIO()
    export_df = df.copy()
    export_df.insert(0, 'No', range(1, 1 + len(export_df)))
    
    # Mengubah ID file menjadi link URL yang bisa diklik
    if 'link_bukti' in export_df.columns:
        export_df['Link Lampiran'] = export_df['link_bukti'].apply(
            lambda x: f"{APP_URL}/?file_id={x}" if pd.notna(x) and x != "-" else "Tidak ada lampiran"
        )
    
    cols = ['No', 'waktu_input', 'no_bpjs', 'nama_pasien', 'no_telp', 'keluhan', 'solusi']
    if 'Link Lampiran' in export_df.columns:
        cols.append('Link Lampiran')
        
    export_df = export_df[cols]
    
    new_names = ['No', 'Tanggal', 'No BPJS', 'Nama Pasien', 'Kontak', 'Keluhan', 'Solusi']
    if 'Link Lampiran' in export_df.columns:
        new_names.append('Link Download Bukti')
        
    export_df.columns = new_names

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name=title)
        worksheet = writer.sheets[title]
        from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
        
        # Warna Biru Elegan untuk header
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        for cell in worksheet["1:1"]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center")
            cell.fill = header_fill

        thin = Side(style='thin')
        for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for cell in row:
                cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                
        widths = {'A': 5, 'B': 20, 'C': 18, 'D': 25, 'E': 15, 'F': 40, 'G': 40, 'H': 50}
        for col_letter, w in zip(['A','B','C','D','E','F','G','H'], list(widths.values())[:len(new_names)]):
            worksheet.column_dimensions[col_letter].width = w
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
    st.caption("Arahkan kamera HP Anda ke Barcode di Klinik untuk membuka form ini.")
    
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
        
        # FITUR BARU: BISA UPLOAD VIDEO DAN PDF
        st.write("Lampirkan Foto/Video/Dokumen Bukti (Opsional - Maks 20MB)")
        file_p = st.file_uploader("Pilih File (Bisa Video MP4)", type=['jpg', 'jpeg', 'png', 'mp4', 'pdf'])
        
        submit_pasien = st.form_submit_button("Kirim Laporan", use_container_width=True)
        
        if submit_pasien:
            if nama_p and bpjs_p and telp_p and alamat_p and keluhan_p:
                waktu_sekarang = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                # Kirim media ke Telegram dulu untuk dapat Resi (ID)
                pesan_notif = f"🚨 <b>LAPORAN BARU!</b> 🚨\nKategori: {jenis}\nNama: {nama_p}\nKeluhan: {keluhan_p}"
                file_id = kirim_notif_telegram_dengan_file(pesan_notif, file_p)
                link_rahasia = file_id if file_id else "-"
                
                conn = init_db()
                conn.execute('''INSERT INTO pipp 
                                (waktu_input, no_bpjs, nama_pasien, alamat, no_telp, jenis_laporan, keluhan, solusi, link_bukti, is_real) 
                                VALUES (?,?,?,?,?,?,?,?,?,1)''',
                             (waktu_sekarang, bpjs_p, nama_p, alamat_p, telp_p, jenis, keluhan_p, '[MENUNGGU TINDAKAN]', link_rahasia))
                conn.commit()
                
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
        
        tab1, tab2 = st.tabs(["Antrean Solusi", "Database Arsip (Terpisah)"])

        with tab1:
            st.subheader("Daftar Laporan Menunggu")
            df_pending = pd.read_sql_query("SELECT id, waktu_input, jenis_laporan, nama_pasien, keluhan FROM pipp WHERE solusi = '[MENUNGGU TINDAKAN]'", conn)
            
            if not df_pending.empty:
                st.dataframe(df_pending, use_container_width=True)
                with st.form("form_solusi"):
                    id_laporan = st.selectbox("Selesaikan Laporan ID:", df_pending['id'].tolist())
                    st.info("💡 Lampiran foto/video otomatis terkirim ke Telegram Anda. Buka Telegram untuk melihatnya.")
                    
                    solusi_admin = st.text_area("Tuliskan Penanganan yang telah dilakukan:")
                    if st.form_submit_button("Simpan & Tutup Laporan"):
                        conn.execute("UPDATE pipp SET solusi=? WHERE id=?", (solusi_admin, id_laporan))
                        conn.commit()
                        st.success("Laporan berhasil ditangani!")
                        st.rerun()
            else:
                st.info("Semua laporan sudah tertangani.")

        with tab2:
            st.subheader("Pusat Arsip & Export Excel")
            
            try:
                df_all = pd.read_sql_query("SELECT * FROM pipp ORDER BY waktu_input DESC", conn)
            except:
                df_all = pd.DataFrame()
            
            if not df_all.empty:
                df_pipp_bpjs = df_all[df_all['jenis_laporan'].str.contains('PIPP', na=False)]
                df_internal = df_all[df_all['jenis_laporan'].str.contains('Internal', na=False)]
                
                st.write("---")
                st.markdown("### 🟢 ARSIP 1: Data Khusus PIPP BPJS")
                st.dataframe(df_pipp_bpjs.drop(columns=['id', 'is_real', 'foto_bukti'], errors='ignore'), use_container_width=True)
                if not df_pipp_bpjs.empty:
                    st.download_button("📥 Download Excel PIPP BPJS", create_excel(df_pipp_bpjs, 'PIPP BPJS'), "Laporan_PIPP_BPJS_Klinik.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                
                st.write("---")
                st.markdown("### 🔴 ARSIP 2: Data Pengaduan Internal")
                st.dataframe(df_internal.drop(columns=['id', 'is_real', 'foto_bukti'], errors='ignore'), use_container_width=True)
                if not df_internal.empty:
                    # Ditambahkan tombol download Excel untuk laporan internal
                    st.download_button("📥 Download Excel Pengaduan", create_excel(df_internal, 'Internal'), "Laporan_Pengaduan_Internal.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.info("Belum ada arsip.")
