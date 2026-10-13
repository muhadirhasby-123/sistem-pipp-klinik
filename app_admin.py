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
    from docx.shared import Pt
    from fpdf import FPDF
except ImportError:
    pass # Akan di-handle oleh st.warning di UI jika belum terinstall

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
    
    # Penambahan Parameter Baru Sesuai Request
    if 'petugas' not in columns: cursor.execute("ALTER TABLE pipp ADD COLUMN petugas TEXT DEFAULT ''")
    if 'subkategori' not in columns: cursor.execute("ALTER TABLE pipp ADD COLUMN subkategori TEXT DEFAULT ''")
    if 'is_audited' not in columns: cursor.execute("ALTER TABLE pipp ADD COLUMN is_audited INTEGER DEFAULT 0")
    if 'jenis_pasien' not in columns: cursor.execute("ALTER TABLE pipp ADD COLUMN jenis_pasien TEXT DEFAULT 'Umum'")
    conn.commit()

init_db()
upgrade_db()

# =====================================================================
# FUNGSI EXPORT (EXCEL, WORD, PDF) PROFESIONAL
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
        column_length = max(dataframe[column_title].astype(str).apply(lambda x: len(x) if x != 'nan' else 0).max(), len(str(column_title)))
        worksheet.column_dimensions[column_letter].width = min(column_length + 2, 45)

    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
        for cell in row:
            cell.border = thin_border
            cell.alignment = Alignment(vertical="top", wrap_text=True)

def create_excel(df, title):
    output = io.BytesIO()
    export_df = df.copy().astype(str).replace('nan', '')
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name=title)
        style_excel_sheet(writer.sheets[title], export_df)
    return output.getvalue()

def create_word(df, title):
    doc = Document()
    doc.add_heading(f'Laporan Klinik: {title}', 0)
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
    
    # Header
    pdf.set_fill_color(15, 23, 42)
    pdf.set_text_color(255, 255, 255)
    for i, col in enumerate(export_df.columns):
        pdf.cell(col_widths[i], 8, str(col)[:20], border=1, fill=True, align='C')
    pdf.ln()
    
    # Data
    pdf.set_font("helvetica", "", 8)
    pdf.set_text_color(0, 0, 0)
    for index, row in export_df.iterrows():
        for i, val in enumerate(row):
            pdf.cell(col_widths[i], 8, str(val)[:25], border=1)
        pdf.ln()
        
    return bytes(pdf.output())

# =====================================================================
# PROGRAM UTAMA ADMIN (DENGAN UI/UX MODERN)
# =====================================================================
st.set_page_config(page_title="Dashboard Admin", layout="wide", initial_sidebar_state="collapsed")

# --- UI/UX CUSTOM CSS (Desain Hidup, Modern, Animasi) ---
st.markdown("""
<style>
    .stApp { background-color: #f1f5f9 !important; font-family: 'Inter', 'Segoe UI', sans-serif;}
    
    /* Header Gradient & Animasi */
    .dashboard-header {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        padding: 20px; border-radius: 12px; color: white; margin-bottom: 20px;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
        display: flex; justify-content: space-between; align-items: center;
    }
    
    /* Tabel Modern */
    div[data-testid="stDataFrame"] > div { 
        max-height: 500px !important; overflow-y: auto !important; 
        border-radius: 12px; border: 1px solid #e2e8f0; 
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        background-color: white;
    }
    
    /* Kartu Form Update (Animasi Hover) */
    .update-card {
        background-color: white; padding: 20px; border-radius: 12px;
        border: 1px solid #e2e8f0; border-left: 5px solid #0ea5e9;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
        margin-bottom: 15px;
    }
    .update-card:hover {
        transform: translateY(-2px); box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }
    
    /* Tombol Super Interaktif */
    div[data-testid="stFormSubmitButton"] > button, .btn-primary > button {
        background: linear-gradient(135deg, #0ea5e9 0%, #0284c7 100%) !important;
        color: white !important; border-radius: 8px !important; width: 100% !important;
        font-weight: 700 !important; transition: all 0.3s ease !important; border: none !important;
    }
    div[data-testid="stFormSubmitButton"] > button:hover, .btn-primary > button:hover {
        box-shadow: 0 8px 15px rgba(2, 132, 199, 0.4) !important; transform: translateY(-2px) !important;
    }
</style>
""", unsafe_allow_html=True)

if "admin_logged_in" not in st.session_state: st.session_state.admin_logged_in = False

if not st.session_state.admin_logged_in:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br><br>", unsafe_allow_html=True)
        st.info("🔐 Silakan masukkan PIN untuk mengakses Database.")
        pin_input = st.text_input("PIN Keamanan:", type="password")
        if st.button("Masuk / Login", type="primary"):
            if pin_input == PIN_ADMIN_RAHASIA: st.session_state.admin_logged_in = True; st.rerun()
            else: st.error("PIN Salah!")
else:
    conn = get_conn()
    
    # --- HEADER MODERN ---
    st.markdown("""
        <div class="dashboard-header">
            <div>
                <h2 style='margin: 0; color: white;'>🏥 Dashboard Layanan Papan Mas</h2>
                <p style='margin: 0; color: #94a3b8;'>Pusat Kendali Data Pasien & Analitik</p>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    if st.button("🚪 Logout App", use_container_width=False): 
        st.session_state.admin_logged_in = False; st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🚀 Pencarian & Kelola Pasien", 
        "📑 Ekspor Data Spesifik", 
        "📝 Audit Logbook", 
        "📊 Visual Analitik", 
        "✍️ Input Data Baru"
    ])

    # ---------------------------------------------------------
    # TAB 1: PENCARIAN CERDAS & UPDATE (TIDAK PAKAI DROPDOWN KAKU)
    # ---------------------------------------------------------
    with tab1:
        df_all = pd.read_sql_query("SELECT * FROM pipp ORDER BY id DESC", conn).fillna('')
        df_all['waktu_input_dt'] = pd.to_datetime(df_all['waktu_input'], errors='coerce')
        
        col_kiri, col_kanan = st.columns([1.5, 1.2], gap="large")
        
        # --- BAGIAN KIRI: TABEL ---
        with col_kiri:
            st.markdown("<h5 style='color:#0f172a;'>📋 Data Pasien Terkini</h5>", unsafe_allow_html=True)
            df_display = df_all[['id', 'waktu_input', 'nama_pasien', 'no_bpjs', 'keluhan', 'status']].copy()
            
            # Pembersihan UI Tabel (Agar tidak ada tulisan default yang jelek)
            kata_buang = ['-', '', 'Belum Diklasifikasi', 'Belum Ditentukan', '[MENUNGGU TINDAKAN]']
            for col in ['nama_pasien', 'no_bpjs', 'keluhan']:
                df_display[col] = df_display[col].apply(lambda x: '-' if str(x).strip() in kata_buang else x)
                
            df_display = df_display.rename(columns={'id':'ID', 'waktu_input':'Waktu', 'nama_pasien':'Nama', 'no_bpjs':'Identitas', 'keluhan':'Keluhan', 'status':'Status'})
            
            # Highlight status (Opsional: Pandas Styler)
            st.dataframe(df_display, use_container_width=True, hide_index=True)

        # --- BAGIAN KANAN: SMART SEARCH & FORM UPDATE ---
        with col_kanan:
            st.markdown("<h5 style='color:#0f172a;'>🔍 Cari & Lengkapi Data</h5>", unsafe_allow_html=True)
            st.info("Ketik nama, ID, NIK, atau WA pasien di bawah ini untuk memunculkan form edit.")
            
            # SMART SEARCH INPUT
            cari_pasien = st.text_input("Pencarian Pasien:", placeholder="Cari nama pasien...", label_visibility="collapsed")
            
            if cari_pasien:
                # Filter data berdasarkan pencarian
                sq = cari_pasien.lower()
                hasil_cari = df_all[
                    df_all['id'].astype(str).str.contains(sq) | 
                    df_all['nama_pasien'].str.lower().str.contains(sq) | 
                    df_all['no_bpjs'].str.lower().str.contains(sq) | 
                    df_all['no_telp'].str.lower().str.contains(sq)
                ]
                
                if not hasil_cari.empty:
                    # Ambil hasil teratas (atau biarkan user pilih jika ada banyak, tp kita ambil 1 untuk kecepatan)
                    for index, row in hasil_cari.head(3).iterrows(): # Maksimal tampil 3 kartu form agar tidak berat
                        st.markdown(f"<div class='update-card'>", unsafe_allow_html=True)
                        st.markdown(f"**Mengelola Data: {row['nama_pasien']}** (ID: {row['id']})")
                        
                        with st.form(f"form_smart_{row['id']}"):
                            # Pembersihan Placeholder
                            v_bpjs = "" if str(row['no_bpjs']).strip() in kata_buang else row['no_bpjs']
                            v_wa = "" if str(row['no_telp']).strip() in kata_buang else row['no_telp']
                            v_kat = "" if str(row['jenis_laporan']).strip() in kata_buang else row['jenis_laporan']
                            v_pet = "" if str(row['petugas']).strip() in kata_buang else row['petugas']
                            v_sol = "" if str(row['solusi']).strip() in kata_buang else row['solusi']
                            v_alamat = "" if str(row['alamat']).strip() in kata_buang else row['alamat']
                            
                            c1, c2 = st.columns(2)
                            with c1:
                                nama_b = st.text_input("Nama Pasien", value=row['nama_pasien'])
                                bpjs_b = st.text_input("Identitas (NIK/BPJS)", value=v_bpjs)
                                jns_pasien = st.selectbox("Jenis Pasien", ["Umum", "BPJS", "Asuransi Lain", "VIP"], index=["Umum", "BPJS", "Asuransi Lain", "VIP"].index(row['jenis_pasien']) if row['jenis_pasien'] in ["Umum", "BPJS", "Asuransi Lain", "VIP"] else 0)
                                kat_b = st.text_input("Kategori Masalah", value=v_kat)
                                alamat_b = st.text_input("Alamat Asal", value=v_alamat)
                            with c2:
                                wa_b = st.text_input("WhatsApp / HP", value=v_wa)
                                pet_b = st.text_input("Dokter / Staf Penindak", value=v_pet)
                                prio_list = ["Normal", "Penting", "Mendesak"]
                                prio_b = st.selectbox("Prioritas", prio_list, index=prio_list.index(row['priority']) if row['priority'] in prio_list else 0)
                                
                                stat_list = ["Baru", "Diproses", "Perlu Cek Ulang", "Selesai", "Ditutup"]
                                stat_b = st.selectbox("Status Laporan", stat_list, index=stat_list.index(row['status']) if row['status'] in stat_list else 0)
                            
                            sol_b = st.text_area("Tindakan / Solusi dari Buku Log", value=v_sol, height=70)
                            
                            if st.form_submit_button("✔ Simpan Pembaruan"):
                                w_skrng = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                w_selesai = row['waktu_selesai']
                                if stat_b in ["Selesai", "Ditutup"] and w_selesai == '': w_selesai = w_skrng
                                
                                conn.execute("""UPDATE pipp SET 
                                             nama_pasien=?, no_bpjs=?, no_telp=?, alamat=?, jenis_pasien=?, priority=?, 
                                             jenis_laporan=?, petugas=?, solusi=?, status=?, waktu_selesai=? WHERE id=?""", 
                                             (nama_b, bpjs_b, wa_b, alamat_b, jns_pasien, prio_b, kat_b, pet_b, sol_b, stat_b, w_selesai, row['id']))
                                conn.commit(); st.success("✅ Data tersimpan!"); st.rerun()
                        st.markdown("</div>", unsafe_allow_html=True)
                else:
                    st.warning("⚠️ Pasien tidak ditemukan. Coba kata kunci lain atau gunakan Tab 'Input Data Baru'.")

    # ---------------------------------------------------------
    # TAB 2: EKSPOR SPESIFIK (EXCEL, WORD, PDF) DENGAN FILTER
    # ---------------------------------------------------------
    with tab2:
        st.markdown("<h4 style='color:#0f172a;'>📑 Generator Laporan Spesifik</h4>", unsafe_allow_html=True)
        st.info("Saring data pasien yang ingin Anda masukkan ke dalam laporan, lalu pilih format filenya (Excel, Word, atau PDF).")
        
        df_ekspor = df_all.copy()
        
        # 1. Panel Filter
        ef1, ef2, ef3 = st.columns(3)
        with ef1: e_tgl = st.date_input("1. Filter Tanggal Laporan:", value=())
        with ef2: e_stat = st.multiselect("2. Filter Status:", options=df_ekspor['status'].unique())
        with ef3: e_jns = st.multiselect("3. Filter Jenis Pasien:", options=["Umum", "BPJS", "Asuransi Lain", "VIP"])
        
        if len(e_tgl) == 2: df_ekspor = df_ekspor[(df_ekspor['waktu_input_dt'].dt.date >= e_tgl[0]) & (df_ekspor['waktu_input_dt'].dt.date <= e_tgl[1])]
        elif len(e_tgl) == 1: df_ekspor = df_ekspor[df_ekspor['waktu_input_dt'].dt.date == e_tgl[0]]
        if e_stat: df_ekspor = df_ekspor[df_ekspor['status'].isin(e_stat)]
        if e_jns: df_ekspor = df_ekspor[df_ekspor['jenis_pasien'].isin(e_jns)]
        
        # 2. Pemilihan Kolom
        st.markdown("**4. Pilih Kolom Data yang ingin di-print/ekspor:**")
        semua_kolom = ['waktu_input', 'nama_pasien', 'no_bpjs', 'no_telp', 'alamat', 'jenis_pasien', 'keluhan', 'jenis_laporan', 'solusi', 'petugas', 'status', 'priority']
        kolom_terpilih = st.multiselect("Kolom:", options=semua_kolom, default=['waktu_input', 'nama_pasien', 'keluhan', 'solusi', 'status'])
        
        st.markdown(f"**Pratinjau Data (Ditemukan {len(df_ekspor)} baris):**")
        df_final_ekspor = df_ekspor[kolom_terpilih] if kolom_terpilih else df_ekspor
        st.dataframe(df_final_ekspor.head(3), use_container_width=True) # Tampil 3 aja sbg preview
        
        # 3. Tombol Ekspor
        st.markdown("---")
        ex1, ex2, ex3 = st.columns(3)
        judul_file = f"Laporan_Klinik_{datetime.now().strftime('%Y%m%d')}"
        
        with ex1:
            if st.button("📊 Generate Excel (.xlsx)", use_container_width=True):
                st.download_button("⬇️ Download Excel Sekarang", create_excel(df_final_ekspor, "Data"), f"{judul_file}.xlsx", type="primary", use_container_width=True)
        with ex2:
            if st.button("📝 Generate Word (.docx)", use_container_width=True):
                try: st.download_button("⬇️ Download Word Sekarang", create_word(df_final_ekspor, "Data Spesifik"), f"{judul_file}.docx", type="primary", use_container_width=True)
                except Exception as e: st.error("Library python-docx belum terinstall.")
        with ex3:
            if st.button("📄 Generate PDF (.pdf)", use_container_width=True):
                try: st.download_button("⬇️ Download PDF Sekarang", create_pdf(df_final_ekspor, "Data Terfilter"), f"{judul_file}.pdf", mime="application/pdf", type="primary", use_container_width=True)
                except Exception as e: st.error("Library fpdf2 belum terinstall.")

    # ---------------------------------------------------------
    # TAB 3: AUDIT LOGBOOK 
    # ---------------------------------------------------------
    with tab3:
        st.markdown("#### 📝 Validasi & Pencocokan Logbook Fisik")
        df_audit = df_all.copy()
        sudah_dicek = len(df_audit[df_audit['is_audited'] == 1])
        c_met1, c_met2 = st.columns(2)
        c_met1.metric("Total Data", len(df_audit))
        c_met2.metric("Sisa Belum Divalidasi", len(df_audit) - sudah_dicek, delta=f"{sudah_dicek} Tervalidasi", delta_color="normal")
        
        df_pending = df_audit[df_audit['is_audited'] != 1]
        ca1, ca2 = st.columns([1.5, 1])
        with ca1: st.dataframe(df_pending[['id', 'waktu_input', 'nama_pasien', 'keluhan']].rename(columns={'waktu_input':'Tanggal'}), hide_index=True)
        with ca2:
            val_id = st.selectbox("Validasi ID:", [None] + df_pending['id'].tolist())
            if val_id:
                if st.button("✔ Tandai Sesuai Logbook (Valid)", type="primary"):
                    conn.execute("UPDATE pipp SET is_audited = 1 WHERE id = ?", (val_id,))
                    conn.commit(); st.rerun()

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
