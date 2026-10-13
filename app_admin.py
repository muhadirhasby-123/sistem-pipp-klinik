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
    
    # Tambah kolom audit untuk tracking pencocokan logbook
    if 'petugas' not in columns: cursor.execute("ALTER TABLE pipp ADD COLUMN petugas TEXT DEFAULT ''")
    if 'subkategori' not in columns: cursor.execute("ALTER TABLE pipp ADD COLUMN subkategori TEXT DEFAULT ''")
    if 'is_audited' not in columns: cursor.execute("ALTER TABLE pipp ADD COLUMN is_audited INTEGER DEFAULT 0")
    conn.commit()

init_db()
upgrade_db()

# =====================================================================
# FUNGSI EXCEL PROFESIONAL (STYLING OPENPYXL)
# =====================================================================
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
    
    # Ganti nama header menjadi lebih rapi
    header_mapping = {
        'waktu_input': 'Tanggal Input', 'no_bpjs': 'No. Identitas', 'nama_pasien': 'Nama Pasien', 
        'no_telp': 'Kontak (WA)', 'jenis_laporan': 'Kategori Utama', 'subkategori': 'Subkategori',
        'keluhan': 'Detail Keluhan', 'solusi': 'Tindakan/Solusi', 'petugas': 'Nama Petugas', 'status': 'Status Akhir'
    }
    export_df = export_df.rename(columns=header_mapping)
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name=title)
        workbook = writer.book
        worksheet = writer.sheets[title]
        
        # Styling Header
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        for col_num, column_title in enumerate(export_df.columns, 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = thin_border
            
            # Auto-fit Width
            column_letter = get_column_letter(col_num)
            column_length = max(export_df[column_title].astype(str).map(len).max(), len(str(column_title)))
            worksheet.column_dimensions[column_letter].width = column_length + 2

        # Styling Seluruh Data
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center", wrap_text=True)

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
st.set_page_config(page_title="Dashboard Admin Papan Mas", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
    .stApp { background-color: #f8fafc !important; font-family: 'Inter', sans-serif;}
    div[data-testid="stDataFrame"] > div { max-height: 450px !important; overflow-y: auto !important; border-radius: 8px; border: 1px solid #e2e8f0; }
    div[data-testid="stFormSubmitButton"] > button {
        background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%) !important; color: white !important; border-radius: 6px !important; width: 100% !important;
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
            
            st.markdown("<p style='font-size: 14px; color: #475569; font-weight: 600; margin-bottom: -10px;'>🔍 Filter Data</p>", unsafe_allow_html=True)
            f_col1, f_col2 = st.columns([2, 1])
            with f_col1: search_query = st.text_input("Pencarian:", placeholder="Cari Nama, BPJS, NIK, atau Keluhan untuk dicocokkan...", label_visibility="collapsed")
            with f_col2: tanggal_filter = st.date_input("Rentang Tanggal:", value=(), label_visibility="collapsed")

            df_filtered = df_all.copy()
            if search_query:
                sq = search_query.lower()
                df_filtered = df_filtered[
                    df_filtered['id'].astype(str).str.contains(sq) | df_filtered['nama_pasien'].str.lower().str.contains(sq) | df_filtered['no_bpjs'].str.lower().str.contains(sq) | df_filtered['keluhan'].str.lower().str.contains(sq)
                ]
                
            if len(tanggal_filter) == 2: df_filtered = df_filtered[(df_filtered['waktu_input_dt'].dt.date >= tanggal_filter[0]) & (df_filtered['waktu_input_dt'].dt.date <= tanggal_filter[1])]
            elif len(tanggal_filter) == 1: df_filtered = df_filtered[df_filtered['waktu_input_dt'].dt.date == tanggal_filter[0]]

            st.markdown("<br>", unsafe_allow_html=True)
            df_display = df_filtered[['id', 'waktu_input', 'no_bpjs', 'nama_pasien', 'keluhan', 'jenis_laporan', 'status']].copy()
            
            for col in ['jenis_laporan', 'nama_pasien', 'no_bpjs', 'keluhan']:
                df_display[col] = df_display[col].apply(lambda x: '-' if x in ['', '-'] else x)
                
            df_display = df_display.rename(columns={'id': 'ID', 'waktu_input': 'Tanggal Input', 'no_bpjs': 'Identitas (BPJS/NIK)', 'nama_pasien': 'Nama Pasien', 'keluhan': 'Keluhan', 'jenis_laporan': 'Kategori', 'status': 'Status'})
            
            st.dataframe(df_display, use_container_width=True, hide_index=True)
            
            st.markdown("<hr style='border-top: 1px solid #e2e8f0; margin: 20px 0px;'>", unsafe_allow_html=True)
            
            sel_col1, sel_col2 = st.columns([1, 4])
            with sel_col1: st.markdown("<p style='font-size: 15px; font-weight: bold; margin-top: 8px;'>📝 Kelola Data (ID):</p>", unsafe_allow_html=True)
            with sel_col2: selected_id = st.selectbox("ID Laporan", [None] + df_filtered['id'].tolist(), format_func=lambda x: f"Pilih ID: {x}" if x else "--- Klik untuk memilih ID Laporan yang akan dikelola ---", label_visibility="collapsed")
            
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
                                if stat_baru in ["Selesai", "Ditutup"] and wkt_selesai == '': wkt_selesai = wkt_skrng
                                
                                conn.execute("""UPDATE pipp SET 
                                             no_bpjs = ?, no_telp = ?, jenis_laporan = ?, subkategori = ?, petugas = ?, solusi = ?, status = ?, waktu_selesai = ?
                                             WHERE id = ?""", 
                                             (bpjs_baru, telp_baru, kat_baru, subkat_baru, pet_baru, sol_baru, stat_baru, wkt_selesai, selected_id))
                                conn.commit(); st.success("✅ Data berhasil diperbarui!"); st.rerun()
        except Exception as e: st.error(f"Error Database: {e}")

    # ---------------------------------------------------------
    # TAB 2: AUDIT LOGBOOK (PENCOCOKAN MASAL)
    # ---------------------------------------------------------
    with tab2:
        st.markdown("#### 📝 Validasi & Pencocokan Logbook Fisik")
        st.info("💡 **Petunjuk:** Halaman ini menampilkan SELURUH data historis. Gunakan halaman ini untuk menandai pasien yang datanya sudah Anda pastikan cocok/valid dengan buku log fisik Anda.")
        
        try:
            df_audit = pd.read_sql_query("SELECT id, waktu_input, nama_pasien, keluhan, is_audited FROM pipp ORDER BY id DESC", conn)
            df_audit = df_audit.fillna('')
            
            # Hitung Statistik Audit
            total_data = len(df_audit)
            sudah_dicek = len(df_audit[df_audit['is_audited'] == 1])
            belum_dicek = total_data - sudah_dicek
            
            c_met1, c_met2 = st.columns(2)
            c_met1.metric("Total Seluruh Data", total_data)
            c_met2.metric("Sisa Data Belum Divalidasi", belum_dicek, delta=f"{sudah_dicek} Sudah Valid", delta_color="normal")
            
            st.markdown("<hr style='margin: 10px 0;'>", unsafe_allow_html=True)
            
            # Tampilkan data yang BELUM diaudit
            df_pending_audit = df_audit[df_audit['is_audited'] != 1].copy()
            
            ca1, ca2 = st.columns([1.5, 1])
            with ca1:
                st.markdown("**Daftar Pasien Belum Divalidasi:**")
                st.dataframe(df_pending_audit[['id', 'waktu_input', 'nama_pasien', 'keluhan']].rename(columns={'waktu_input': 'Tanggal'}), use_container_width=True, hide_index=True)
                
            with ca2:
                st.markdown("**Validasi Pasien**")
                val_id = st.selectbox("Pilih ID untuk divalidasi:", [None] + df_pending_audit['id'].tolist(), format_func=lambda x: f"Validasi ID: {x}" if x else "--- Pilih ID ---")
                
                if val_id:
                    pasien_val = df_pending_audit[df_pending_audit['id'] == val_id].iloc[0]
                    st.success(f"**Nama:** {pasien_val['nama_pasien']}")
                    st.write(f"**Keluhan:** {pasien_val['keluhan']}")
                    
                    if st.button("✔ Tandai Sudah Sesuai (Valid)", type="primary", use_container_width=True):
                        conn.execute("UPDATE pipp SET is_audited = 1 WHERE id = ?", (val_id,))
                        conn.commit(); st.rerun()
        except Exception as e: st.error(f"Error Audit: {e}")

    # ---------------------------------------------------------
    # TAB 3: ANALITIK (GRAFIK PROFESIONAL)
    # ---------------------------------------------------------
    with tab3:
        st.markdown("#### 📊 Visualisasi Analitik Klinik")
        tampilkan_grafik = st.toggle("Aktifkan Grafik (ON/OFF)", value=False)
        
        if tampilkan_grafik:
            try:
                df_stat = pd.read_sql_query("SELECT * FROM pipp", conn).fillna('')
                c_graf1, c_graf2 = st.columns(2)
                
                # Pengaturan Layout Grafik Bawaan (Agar Transparan & Elegan)
                layout_style = dict(
                    paper_bgcolor='rgba(0,0,0,0)',
                    plot_bgcolor='rgba(0,0,0,0)',
                    font=dict(family="Inter, sans-serif", size=13, color="#334155"),
                    margin=dict(t=30, b=10, l=10, r=10)
                )
                
                with c_graf1:
                    with st.container(border=True):
                        st.markdown("**Rasio Status Penanganan**")
                        status_counts = df_stat['status'].value_counts().reset_index()
                        status_counts.columns = ['Status', 'Jumlah']
                        
                        fig_pie = px.pie(status_counts, values='Jumlah', names='Status', hole=0.5, 
                                         color_discrete_sequence=['#0ea5e9', '#3b82f6', '#f59e0b', '#10b981', '#64748b'])
                        fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                        fig_pie.update_layout(**layout_style, showlegend=False)
                        st.plotly_chart(fig_pie, use_container_width=True)
                
                with c_graf2:
                    with st.container(border=True):
                        st.markdown("**Top 5 Kategori Laporan**")
                        df_kat = df_stat[(df_stat['jenis_laporan'] != '') & (df_stat['jenis_laporan'] != '-')]
                        kat_counts = df_kat['jenis_laporan'].value_counts().head(5).reset_index()
                        kat_counts.columns = ['Kategori', 'Jumlah']
                        
                        fig_bar = px.bar(kat_counts, x='Jumlah', y='Kategori', orientation='h', 
                                         color_discrete_sequence=['#0284c7'])
                        fig_bar.update_layout(**layout_style, yaxis={'categoryorder':'total ascending'})
                        st.plotly_chart(fig_bar, use_container_width=True)
            except: 
                st.warning("Pastikan Anda sudah menambahkan `plotly` di requirements.txt")
        else:
            st.info("💡 Grafik dimatikan untuk menghemat beban sistem.")

    # ---------------------------------------------------------
    # TAB 4 - 7: FITUR LAMA 
    # ---------------------------------------------------------
    with tab4:
        try: st.dataframe(pd.read_sql_query("SELECT * FROM contacts ORDER BY last_interaction_at DESC", conn), use_container_width=True)
        except: pass

    with tab5:
        with st.form("form_manual_staf"):
            st.info("Form Input Laporan Manual")
            m_nama = st.text_input("👤 Nama Pasien *")
            col_m1, col_m2 = st.columns(2)
            with col_m1: m_bpjs = st.text_input("🏥 Nomor BPJS/NIK (Opsional)")
            with col_m2: m_telp = st.text_input("📞 Nomor WA/HP (Opsional)")
                
            m_jenis = st.text_input("Kategori Laporan", placeholder="Contoh: PIPP BPJS")
            m_keluhan = st.text_area("💬 Detail Keluhan *")
            m_solusi = st.text_area("✅ Tindakan/Solusi *")
            
            if st.form_submit_button("Simpan Laporan Baru"):
                if m_nama and m_keluhan and m_solusi:
                    wkt_skrg = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    conn.execute('''INSERT INTO pipp (waktu_input, waktu_selesai, no_bpjs, nama_pasien, alamat, no_telp, jenis_laporan, keluhan, solusi, link_bukti, is_real, status)
                                    VALUES (?,?,?,?,'-',?,?,?,?,'-',1, 'Selesai')''',
                                 (wkt_skrg, wkt_skrg, m_bpjs, m_nama, m_telp, m_jenis, m_keluhan, m_solusi))
                    conn.commit(); st.success("Data berhasil ditambahkan!")
                else: st.error("Lengkapi Nama, Keluhan, dan Solusi.")

    with tab6:
        try:
            df_arsip = pd.read_sql_query("SELECT * FROM pipp ORDER BY id DESC", conn)
            st.dataframe(df_arsip, use_container_width=True, hide_index=True)
            # Menggunakan Create Excel yang sudah di styling profesional
            st.download_button("📥 Download Excel", create_excel(df_arsip, "Arsip_Klinik"), f"Arsip_Klinik_{datetime.now().strftime('%Y%m%d')}.xlsx")
        except: pass

    with tab7:
        try: st.dataframe(pd.read_sql_query("SELECT * FROM main_categories", conn), use_container_width=True)
        except: pass
