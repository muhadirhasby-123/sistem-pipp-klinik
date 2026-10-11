import streamlit as st
import pandas as pd
import sqlite3
import io
import json
import requests
import re
import string
import random
from datetime import datetime
from PIL import Image, ImageEnhance, ImageOps
import numpy as np

# =====================================================================
# PENGATURAN & NOTIFIKASI
# =====================================================================
PIN_ADMIN_RAHASIA = "123456"
TELEGRAM_BOT_TOKEN = "PASTE_TOKEN_DISINI"
TELEGRAM_CHAT_ID = "PASTE_ID_DISINI"
APP_URL = "https://pipp-klinik.streamlit.app"
DB_PATH = "klinik_pipp.db"

# =====================================================================
# CEK URL PARAMS UNTUK SHORT LINK STATUS
# =====================================================================
if "status" in st.query_params:
    st.session_state["p_step"] = "CEK_STATUS"
    if "r_data" not in st.session_state: st.session_state.r_data = {}
    st.session_state.r_data["cek_tiket"] = st.query_params["status"]

# =====================================================================
# FUNGSI TELEGRAM (SUPPORT MULTI-FILE & VIDEO)
# =====================================================================
def kirim_notif_telegram_dengan_file(pesan_teks, uploaded_files):
    file_ids = []
    if TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID:
        url_text = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        requests.post(url_text, json={"chat_id": TELEGRAM_CHAT_ID, "text": pesan_teks, "parse_mode": "HTML"})
        
        if uploaded_files:
            if not isinstance(uploaded_files, list):
                uploaded_files = [uploaded_files]
            
            for uf in uploaded_files:
                ext = uf.name.split('.')[-1].lower()
                file_bytes = uf.getvalue()
                
                if ext in ['mp4', 'mov', 'avi']:
                    url_api = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendVideo"
                    files = {'video': (uf.name, file_bytes)}
                elif ext in ['jpg', 'jpeg', 'png']:
                    url_api = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendPhoto"
                    files = {'photo': (uf.name, file_bytes)}
                else:
                    url_api = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
                    files = {'document': (uf.name, file_bytes)}
                    
                data = {'chat_id': TELEGRAM_CHAT_ID}
                try:
                    res = requests.post(url_api, data=data, files=files).json()
                    if res.get("ok"):
                        if ext in ['mp4', 'mov', 'avi']:
                            file_ids.append(res["result"]["video"]["file_id"])
                        elif ext in ['jpg', 'jpeg', 'png']:
                            file_ids.append(res["result"]["photo"][-1]["file_id"])
                        else:
                            file_ids.append(res["result"]["document"]["file_id"])
                except Exception:
                    pass
    return ",".join(file_ids) if file_ids else "-"

# =====================================================================
# DATABASE UTAMA & MIGRASI AMAN
# =====================================================================
def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def init_db():
    conn = get_conn()
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS pipp (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    waktu_input DATETIME,
                    waktu_selesai DATETIME,
                    no_bpjs TEXT,
                    nama_pasien TEXT,
                    alamat TEXT,
                    no_telp TEXT,
                    jenis_laporan TEXT,
                    keluhan TEXT,
                    solusi TEXT,
                    link_bukti TEXT,
                    is_real INTEGER)''')
    
    columns_pipp = [col[1] for col in c.execute("PRAGMA table_info(pipp)").fetchall()]
    if "contact_id" not in columns_pipp:
        c.execute("ALTER TABLE pipp ADD COLUMN contact_id INTEGER")
    if "status" not in columns_pipp:
        c.execute("ALTER TABLE pipp ADD COLUMN status TEXT DEFAULT 'Baru'")
    if "priority" not in columns_pipp:
        c.execute("ALTER TABLE pipp ADD COLUMN priority TEXT DEFAULT 'Normal'")
    if "short_link" not in columns_pipp:
        c.execute("ALTER TABLE pipp ADD COLUMN short_link TEXT")
    if "consent_wa" not in columns_pipp:
        c.execute("ALTER TABLE pipp ADD COLUMN consent_wa INTEGER DEFAULT 0")

    c.execute('''CREATE TABLE IF NOT EXISTS contacts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT,
                    whatsapp TEXT UNIQUE,
                    bpjs TEXT,
                    nik TEXT,
                    consent_wa INTEGER DEFAULT 0,
                    first_interaction_at DATETIME,
                    last_interaction_at DATETIME)''')

    c.execute('''CREATE TABLE IF NOT EXISTS interactions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    contact_id INTEGER,
                    module TEXT,
                    details TEXT, 
                    is_ticketed INTEGER DEFAULT 0,
                    created_at DATETIME)''')

    c.execute('''CREATE TABLE IF NOT EXISTS ticket_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ticket_id INTEGER,
                    actor TEXT,
                    action TEXT,
                    note TEXT, 
                    is_internal INTEGER DEFAULT 1,
                    created_at DATETIME)''')

    conn.commit()
    return conn

# =====================================================================
# FUNGSI UTILITY ARSITEKTUR TIKET
# =====================================================================
def generate_ticket_id(prefix):
    today = datetime.now().strftime("%y%m%d")
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT COUNT(*) FROM pipp WHERE waktu_input LIKE ?", (f"{datetime.now().strftime('%Y-%m-%d')} %",))
    count = c.fetchone()[0] + 1
    conn.close()
    return f"{prefix}-{today}-{count:04d}"

def generate_short_link():
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))

def get_or_create_contact(nama, wa, bpjs, nik, consent):
    wa = str(wa).strip()
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id FROM contacts WHERE whatsapp=?", (wa,))
    res = c.fetchone()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if res:
        contact_id = res[0]
        c.execute("UPDATE contacts SET last_interaction_at=?, name=?, consent_wa=? WHERE id=?", (now, nama, consent, contact_id))
    else:
        c.execute("INSERT INTO contacts (name, whatsapp, bpjs, nik, consent_wa, first_interaction_at, last_interaction_at) VALUES (?,?,?,?,?,?,?)",
                  (nama, wa, bpjs, nik, consent, now, now))
        contact_id = c.lastrowid
    conn.commit()
    conn.close()
    return contact_id

# =====================================================================
# GAYA VISUAL PASIEN (PREMIUM UI/UX)
# =====================================================================
st.set_page_config(page_title="Pusat Layanan Pasien", layout="centered", initial_sidebar_state="collapsed")
init_db()

def ui_inject_patient_style():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700;800&family=Outfit:wght@500;700;800;900&display=swap');
        
        .stApp { 
            background-image: linear-gradient(rgba(240, 253, 244, 0.90), rgba(255, 255, 255, 0.98)), url("https://images.unsplash.com/photo-1519494026892-80bbd2d6fd0d?q=80&w=2000&auto=format&fit=crop");
            background-size: cover; background-attachment: fixed; background-position: center; font-family: 'Poppins', sans-serif; 
        }
        .stApp::before {
            content: ""; position: fixed; top: -10%; left: -10%; width: 400px; height: 400px;
            background: radial-gradient(circle, rgba(16,185,129,0.06) 0%, rgba(255,255,255,0) 70%); border-radius: 50%; z-index: -1; animation: floatBlob 10s ease-in-out infinite alternate; pointer-events: none;
        }
        @keyframes floatBlob { 0% { transform: translate(0, 0) scale(1); } 100% { transform: translate(30px, 30px) scale(1.1); } }
        
        div.block-container { padding-top: 2.5rem !important; padding-bottom: 4rem !important; max-width: 650px !important; z-index: 1;}
        
        @keyframes fadeUpSmooth { from { opacity: 0; transform: translateY(20px); } to { opacity: 1; transform: translateY(0); } }
        .anim-1 { animation: fadeUpSmooth 0.5s cubic-bezier(0.16, 1, 0.3, 1) forwards; animation-delay: 0.1s; opacity: 0; }
        .anim-2 { animation: fadeUpSmooth 0.5s cubic-bezier(0.16, 1, 0.3, 1) forwards; animation-delay: 0.2s; opacity: 0; }
        .anim-3 { animation: fadeUpSmooth 0.5s cubic-bezier(0.16, 1, 0.3, 1) forwards; animation-delay: 0.3s; opacity: 0; }
        .anim-4 { animation: fadeUpSmooth 0.5s cubic-bezier(0.16, 1, 0.3, 1) forwards; animation-delay: 0.4s; opacity: 0; }
        .page-trans { animation: fadeUpSmooth 0.3s ease-out forwards; }

        .gradient-text { background: linear-gradient(135deg, #059669 0%, #10b981 100%); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-family: 'Outfit', sans-serif; letter-spacing: -0.5px; text-shadow: 0 4px 15px rgba(16,185,129,0.1); }

        @keyframes pulseGlow { 0% { box-shadow: 0 0 0 0 rgba(16,185,129,0.3); } 70% { box-shadow: 0 0 0 12px rgba(16,185,129,0); } 100% { box-shadow: 0 0 0 0 rgba(16,185,129,0); } }
        .brand-logo-box { background: rgba(255, 255, 255, 0.4); backdrop-filter: blur(10px); border-radius: 16px; width: 50px; height: 50px; display: flex; align-items: center; justify-content: center; margin-right: 15px; flex-shrink: 0; animation: pulseGlow 2.5s infinite; border: 1px solid rgba(16,185,129,0.6); box-shadow: inset 0 0 15px rgba(16,185,129,0.4), 0 4px 10px rgba(0,0,0,0.05); }
        .brand-logo-icon { color: #059669; font-size: 28px; font-weight: 900; line-height: 1; }
        .brand-container { display: flex; align-items: center; margin-bottom: 2rem; margin-top: 0.5rem; margin-left: 2px; }

        div[data-testid="stButton"] > button[kind="secondary"] { background: rgba(255, 255, 255, 0.95) !important; backdrop-filter: blur(15px) !important; border: 1px solid rgba(226, 232, 240, 0.8) !important; border-radius: 16px !important; padding: 24px !important; box-shadow: 0 8px 20px rgba(0, 0, 0, 0.03), inset 0 2px 4px rgba(255,255,255,0.6) !important; text-align: left !important; height: auto !important; width: 100% !important; display: block !important; transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1) !important; margin-bottom: 15px !important; position: relative !important; overflow: hidden !important; z-index: 1 !important; }
        div[data-testid="stButton"] > button[kind="secondary"]:not([disabled]):hover { transform: translateY(-4px) scale(1.01) !important; border-color: #10b981 !important; box-shadow: 0 12px 25px rgba(16,185,129,0.12) !important; background: #ffffff !important; }
        div[data-testid="stButton"] > button[kind="secondary"] p { margin: 0 !important; color: #475569 !important; font-size: 14px !important; line-height: 1.5 !important; white-space: normal !important; position: relative; z-index: 2; font-weight: 500 !important; }
        div[data-testid="stButton"] > button[kind="secondary"] p strong { color: #0f172a !important; font-size: 17px !important; display: block !important; margin-bottom: 8px !important; font-family: 'Outfit', sans-serif !important; font-weight: 800 !important; }
        div[data-testid="stButton"] > button[kind="secondary"] p::first-letter { font-size: 32px !important; font-weight: 900 !important; color: #94a3b8 !important; float: left; line-height: 0.8 !important; margin-right: 6px; margin-top: 4px; padding-bottom: 4px; font-family: 'Outfit', sans-serif !important; }

        div.element-container:has(div.wm-bpjs) + div.element-container button[kind="secondary"]::after { content: "🪪"; position: absolute; right: -10px; bottom: -30px; font-size: 130px; opacity: 0.04; transform: rotate(-15deg); transition: 0.5s ease; z-index: 0;}
        div.element-container:has(div.wm-bpjs) + div.element-container button[kind="secondary"]:hover::after { transform: rotate(0deg) scale(1.1); opacity: 0.08; }
        div.element-container:has(div.wm-adu) + div.element-container button[kind="secondary"]::after { content: "⚠️"; position: absolute; right: -10px; bottom: -20px; font-size: 100px; opacity: 0.04; transform: rotate(10deg); transition: 0.5s ease; z-index: 0;}
        div.element-container:has(div.wm-adu) + div.element-container button[kind="secondary"]:hover::after { transform: rotate(0deg) scale(1.1); opacity: 0.08; }
        div.element-container:has(div.wm-saran) + div.element-container button[kind="secondary"]::after { content: "💡"; position: absolute; right: 0px; bottom: -20px; font-size: 100px; opacity: 0.04; transform: rotate(-10deg); transition: 0.5s ease; z-index: 0;}
        div.element-container:has(div.wm-saran) + div.element-container button[kind="secondary"]:hover::after { transform: rotate(0deg) scale(1.1); opacity: 0.08; }
        div.element-container:has(div.wm-apres) + div.element-container button[kind="secondary"]::after { content: "💜"; position: absolute; right: -5px; bottom: -15px; font-size: 100px; opacity: 0.04; transform: rotate(15deg); transition: 0.5s ease; z-index: 0;}
        div.element-container:has(div.wm-apres) + div.element-container button[kind="secondary"]:hover::after { transform: rotate(0deg) scale(1.1); opacity: 0.08; }

        .stRadio > div[role="radiogroup"] { gap: 14px; }
        .stRadio > div[role="radiogroup"] > label { background: rgba(255,255,255,0.9); border: 2px solid #e2e8f0; padding: 18px 22px; border-radius: 16px; cursor: pointer; transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1); box-shadow: 0 4px 10px rgba(0,0,0,0.03); font-weight: 600; color: #1e293b; display: flex; align-items: center; }
        .stRadio > div[role="radiogroup"] > label:hover { border-color: #10b981; background: #ecfdf5; transform: translateX(8px); box-shadow: 0 8px 20px rgba(16,185,129,0.15); }
        
        div.element-container:has(div.top-back-btn) + div.element-container button { background: transparent !important; border: none !important; color: #64748b !important; padding: 0 !important; width: auto !important; margin-bottom: 15px !important; display: inline-flex !important; font-size: 14px !important; box-shadow: none !important; }
        div.element-container:has(div.top-back-btn) + div.element-container button:hover { color: #10b981 !important; transform: translateX(-4px) !important; }

        .btn-outline > div > button { background-color: rgba(255,255,255,0.95) !important; border: 2px solid #cbd5e1 !important; color: #475569 !important; border-radius: 14px !important; padding: 16px !important; font-weight: 700 !important; font-size: 16px !important; transition: all 0.3s ease !important; width: 100% !important; display: flex !important; justify-content: center !important; backdrop-filter: blur(10px); box-shadow: 0 4px 6px rgba(0,0,0,0.02) !important; }
        .btn-outline > div > button:hover { background-color: #f1f5f9 !important; border-color: #94a3b8 !important; color: #0f172a !important; transform: translateY(-3px); box-shadow: 0 6px 12px rgba(0,0,0,0.05) !important;}
        
        .btn-primary > div > button { background: linear-gradient(135deg, #10b981 0%, #047857 100%) !important; color: #ffffff !important; border: none !important; border-radius: 14px !important; padding: 16px !important; font-weight: 700 !important; font-size: 16px !important; box-shadow: 0 6px 15px rgba(16,185,129,0.4) !important; transition: all 0.3s ease !important; width: 100% !important; display: block !important; }
        .btn-primary > div > button:hover { transform: translateY(-4px) scale(1.02); box-shadow: 0 10px 25px rgba(16,185,129,0.5) !important; }

        .glass-box { background: rgba(255,255,255,0.85); padding: 32px; border-radius: 24px; border: 1px solid rgba(255,255,255,1); margin-bottom: 24px; box-shadow: 0 15px 35px rgba(0,0,0,0.05), inset 0 2px 4px rgba(255,255,255,0.8); backdrop-filter: blur(20px); }
        div[data-baseweb="input"] > div, div[data-baseweb="textarea"] > div { border-radius: 12px !important; border: 2px solid #e2e8f0 !important; background: rgba(255,255,255,0.9) !important; box-shadow: inset 0 2px 4px rgba(0,0,0,0.02) !important; transition: 0.3s;}
        div[data-baseweb="input"] > div:focus-within, div[data-baseweb="textarea"] > div:focus-within { border-color: #10b981 !important; background: #ffffff !important; box-shadow: 0 0 0 4px rgba(16,185,129,0.15) !important;}
        header, footer { visibility: hidden; }
    </style>
    """, unsafe_allow_html=True)

def render_brand_header():
    st.markdown("""
    <div class="anim-1 brand-container">
        <div class="brand-logo-box">
            <!-- GANTI LOGO DISINI -->
            <span class="brand-logo-icon">✚</span>
        </div>
        <div>
            <h3 style="color: #0f172a; margin: 0; font-size: 19px; font-weight: 800; font-family: 'Outfit', sans-serif; letter-spacing: -0.5px;">Klinik Pratama Papan Mas</h3>
            <p style="color: #64748b; margin: 0; font-size: 13px; font-weight: 500;">PT. Tenri Sau Medika Indonesia</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

def top_back_button(target_step):
    st.markdown('<div class="top-back-btn"></div>', unsafe_allow_html=True)
    if st.button("⬅️ Kembali", key=f"top_back_{st.session_state.p_step}"):
        st.session_state.p_step = target_step
        st.rerun()

def get_val(key, default=""): return st.session_state.r_data.get(key, default)

# =====================================================================
# PROGRAM UTAMA: LAYANAN PASIEN
# =====================================================================
ui_inject_patient_style()

if "p_step" not in st.session_state: st.session_state.p_step = "LANDING"
if "r_data" not in st.session_state: st.session_state.r_data = {}

st.markdown('<div class="page-trans">', unsafe_allow_html=True)

if st.session_state.p_step == "LANDING":
    render_brand_header()
    st.markdown("<h1 class='anim-2 gradient-text' style='font-size: 34px; font-weight: 900; margin-bottom: 8px;'>👋 Halo, Ada yang bisa kami bantu?</h1><p class='anim-2' style='color: #475569; font-size: 16px; margin-bottom: 32px; font-weight:500;'>Sampaikan pertanyaan, keluhan, saran, atau apresiasi Anda kepada kami.</p>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="anim-3 wm-bpjs"></div>', unsafe_allow_html=True)
        if st.button("🪪 **Bantuan BPJS / JKN**\n\nPertanyaan atau kendala tentang Mobile JKN, rujukan, antrean, obat, kepesertaan, dan layanan JKN.", use_container_width=True, type="secondary"):
            st.session_state.p_step = "BPJS_MENU"; st.session_state.r_data = {}; st.rerun()

        st.markdown('<div class="anim-4 wm-saran"></div>', unsafe_allow_html=True)
        if st.button("💡 **Saran & Masukan**\n\nPunya ide agar pelayanan kami menjadi lebih baik?", use_container_width=True, type="secondary"):
            st.session_state.p_step = "SARAN_FORM"; st.session_state.r_data = {"active_module": "SARAN"}; st.rerun()
    with c2:
        st.markdown('<div class="anim-3 wm-adu"></div>', unsafe_allow_html=True)
        if st.button("⚠️ **Pengaduan Layanan**\n\nAda masalah dengan pelayanan, petugas, waktu tunggu, fasilitas, atau kenyamanan?", use_container_width=True, type="secondary"):
            st.session_state.p_step = "ADU_PILIH"; st.session_state.r_data = {"active_module": "PENGADUAN"}; st.rerun()

        st.markdown('<div class="anim-4 wm-apres"></div>', unsafe_allow_html=True)
        if st.button("💜 **Apresiasi**\n\nIngin memberikan apresiasi kepada petugas atau pelayanan kami?", use_container_width=True, type="secondary"):
            st.session_state.p_step = "APRES_FORM"; st.session_state.r_data = {"active_module": "APRESIASI"}; st.rerun()
    
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="anim-4 btn-outline">', unsafe_allow_html=True)
    if st.button("🔎 Cek Status Laporan", use_container_width=True): st.session_state.p_step = "CEK_STATUS"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('<p class="anim-4" style="text-align:center; color:#94a3b8; font-size:13px; margin-top:28px; font-weight:500;">🔒 Data yang Anda kirim digunakan untuk membantu kami menindaklanjuti kebutuhan Anda.</p>', unsafe_allow_html=True)

elif st.session_state.p_step == "CEK_STATUS":
    render_brand_header()
    top_back_button("LANDING")
    st.markdown("<h1 class='anim-2 gradient-text' style='font-size: 30px; font-weight: 800; margin-bottom: 8px;'>Cek Status Laporan</h1><p class='anim-2' style='color: #475569; font-size: 16px; margin-bottom: 32px;'>Masukkan nomor laporan atau Link ID Anda di bawah ini.</p>", unsafe_allow_html=True)
    
    st.markdown("<div class='anim-3 glass-box'>", unsafe_allow_html=True)
    tiket_input = st.text_input("Nomor Tiket / Link ID (Contoh: RUJ-260926-0001 atau AB12CD)", key="unik_cek_status", value=get_val("cek_tiket"))
    st.session_state.r_data["cek_tiket"] = tiket_input
    
    st.markdown('<div class="btn-primary" style="margin-top:24px;">', unsafe_allow_html=True)
    if st.button("Cek Sekarang", use_container_width=True):
        val_tiket = st.session_state.r_data.get("cek_tiket", "").strip()
        if val_tiket:
            conn = get_conn()
            df_cek = pd.read_sql_query(f"SELECT * FROM pipp WHERE keluhan LIKE '%{val_tiket}%' OR short_link = '{val_tiket}' LIMIT 1", conn)
            
            if not df_cek.empty:
                data_lapor = df_cek.iloc[0]
                status_db = data_lapor['status'] if 'status' in data_lapor and pd.notna(data_lapor['status']) else 'Baru'
                status_badge = "🟢 SELESAI / DITANGGAPI" if status_db in ["Selesai", "Ditutup"] else f"🟡 {status_db.upper()}"
                
                st.success("Laporan ditemukan!")
                st.markdown(f"**Tanggal Laporan:** {data_lapor['waktu_input']}")
                st.markdown(f"**Jenis:** {data_lapor['jenis_laporan']}")
                st.markdown(f"**Status Saat Ini:** {status_badge}")
                
                df_logs = pd.read_sql_query(f"SELECT created_at, action, note FROM ticket_logs WHERE ticket_id={data_lapor['id']} AND is_internal=0 ORDER BY id DESC", conn)
                if not df_logs.empty:
                    st.markdown("---")
                    st.markdown("### 📋 Riwayat Tindak Lanjut")
                    for _, row_log in df_logs.iterrows():
                        st.info(f"**{row_log['created_at']} - {row_log['action']}**\n\n{row_log['note']}")
                elif data_lapor['solusi'] and data_lapor['solusi'] != "[MENUNGGU TINDAKAN]":
                    st.markdown("---")
                    st.markdown(f"**Tanggapan/Solusi Petugas:**\n> {data_lapor['solusi']}")
            else:
                st.error("❌ Data tidak ditemukan. Pastikan format penulisan sudah benar.")
            conn.close()
        else:
            st.warning("⚠️ Masukkan data terlebih dahulu.")
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

elif st.session_state.p_step == "ADU_PILIH":
    render_brand_header(); top_back_button("LANDING")
    st.markdown("<h1 class='anim-2 gradient-text' style='font-size: 30px; font-weight: 800; margin-bottom: 8px;'>Ada kendala dalam pelayanan?</h1><p class='anim-2' style='color: #475569; font-size: 16px; margin-bottom: 32px;'>Ceritakan kepada kami.</p>", unsafe_allow_html=True)
    opsi = ["Sikap/komunikasi petugas", "Waktu tunggu", "Pelayanan dokter", "Pelayanan perawat", "Pelayanan bidan", "Administrasi", "Pendaftaran", "Farmasi/obat", "Laboratorium", "Fasilitas", "Kebersihan", "Privasi/kenyamanan", "Antrean", "BPJS/JKN", "Lainnya"]
    st.markdown('<div class="anim-3">', unsafe_allow_html=True)
    for op in opsi:
        if st.button(f"**{op}**", use_container_width=True, key=f"adu_btn_{op}"):
            st.session_state.r_data["active_masalah"] = op; st.session_state.p_step = "ADU_FORM"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "ADU_FORM":
    render_brand_header(); top_back_button("ADU_PILIH")
    masalah = get_val("active_masalah")
    st.markdown(f"<h1 class='anim-2 gradient-text' style='font-size: 24px; font-weight: 800; margin-bottom: 28px; line-height:1.4;'>Topik: {masalah}</h1>", unsafe_allow_html=True)
    st.markdown("<div class='anim-3 glass-box'>", unsafe_allow_html=True)

    if masalah == "Sikap/komunikasi petugas":
        opsi_sikap = ["Cara berbicara", "Kurang ramah", "Informasi tidak jelas", "Tidak mendapat penjelasan", "Petugas tidak responsif", "Merasa tidak dihargai", "Lainnya"]
        st.session_state.r_data["ADU_detail"] = st.radio("Masalah terkait apa?", opsi_sikap, index=opsi_sikap.index(get_val("ADU_detail")) if get_val("ADU_detail") in opsi_sikap else None, key="adu_detail_radio")
        opsi_unit = ["Pendaftaran", "Dokter", "Perawat", "Bidan", "Farmasi", "Laboratorium", "Tidak tahu", "Lainnya"]
        st.session_state.r_data["ADU_unit"] = st.radio("Siapa/unit yang terlibat?", opsi_unit, index=opsi_unit.index(get_val("ADU_unit")) if get_val("ADU_unit") in opsi_unit else None, key="adu_unit_radio")
    elif masalah == "Waktu tunggu":
        opsi_tunggu = ["Menunggu pendaftaran", "Menunggu dokter", "Menunggu obat", "Menunggu laboratorium", "Menunggu tindakan", "Menunggu administrasi", "Lainnya"]
        st.session_state.r_data["ADU_detail"] = st.radio("Menunggu apa?", opsi_tunggu, index=opsi_tunggu.index(get_val("ADU_detail")) if get_val("ADU_detail") in opsi_tunggu else None, key="adu_detail_radio")
        st.session_state.r_data["ADU_jam_dtg"] = st.text_input("Perkiraan jam datang", value=get_val("ADU_jam_dtg"), key="adu_jam_dtg")
        st.session_state.r_data["ADU_lama"] = st.text_input("Perkiraan waktu menunggu (mis: 2 jam)", value=get_val("ADU_lama"), key="adu_lama")
        st.session_state.r_data["ADU_no"] = st.text_input("Nomor antrean jika ada", value=get_val("ADU_no"), key="adu_no")
    elif masalah == "Pelayanan dokter":
        opsi_dok = ["Waktu konsultasi", "Komunikasi", "Penjelasan medis", "Sikap", "Proses pemeriksaan", "Lainnya"]
        st.session_state.r_data["ADU_detail"] = st.radio("Masalah terkait apa?", opsi_dok, index=opsi_dok.index(get_val("ADU_detail")) if get_val("ADU_detail") in opsi_dok else None, key="adu_detail_radio")
    elif masalah == "Fasilitas":
        opsi_fas = ["Ruang tunggu", "Kursi", "Toilet", "AC/kipas", "Kebersihan", "Parkir", "Ruang pelayanan", "Peralatan", "Lainnya"]
        st.session_state.r_data["ADU_detail"] = st.radio("Fasilitas apa?", opsi_fas, index=opsi_fas.index(get_val("ADU_detail")) if get_val("ADU_detail") in opsi_fas else None, key="adu_detail_radio")
    elif masalah == "Farmasi/obat":
        opsi_obat = ["Obat tidak tersedia", "Waktu tunggu", "Informasi obat", "Obat tidak sesuai", "Jumlah obat", "Lainnya"]
        st.session_state.r_data["ADU_detail"] = st.radio("Masalah terkait apa?", opsi_obat, index=opsi_obat.index(get_val("ADU_detail")) if get_val("ADU_detail") in opsi_obat else None, key="adu_detail_radio")

    st.markdown("<hr style='border:none; border-top:1px dashed #cbd5e1; margin:20px 0;'>", unsafe_allow_html=True)
    tgl = st.date_input("Tanggal Kejadian", key="adu_tgl_input")
    st.session_state.r_data["ADU_tgl"] = tgl.strftime("%d %B %Y")
    st.session_state.r_data["ADU_jam"] = st.text_input("Jam Kejadian", value=get_val("ADU_jam"), key="adu_jam_input")
    st.session_state.r_data["ADU_lokasi"] = st.text_input("Lokasi Kejadian", value=get_val("ADU_lokasi"), key="adu_lokasi_input")
    st.session_state.r_data["ADU_cerita"] = st.text_area("Ceritakan selengkapnya", height=120, value=get_val("ADU_cerita"), key="adu_cerita_input")
    st.session_state.r_data["ADU_harap"] = st.text_area("Apa harapan Anda? (Opsional)", value=get_val("ADU_harap"), key="adu_harap_input")
    
    f = st.file_uploader("Upload Foto/Video (Opsional)", type=['jpg','png','jpeg','pdf','mp4','mov','avi'], accept_multiple_files=True, key="adu_file_input")
    if f: st.session_state.r_data["file_bukti"] = f
    st.markdown("</div>", unsafe_allow_html=True)
    
    st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
    if st.button("Lanjut ➔", use_container_width=True): 
        for key in ["adu_detail_radio", "adu_unit_radio", "adu_jam_dtg", "adu_lama", "adu_no", "adu_jam_input", "adu_lokasi_input", "adu_cerita_input", "adu_harap_input"]:
            if key in st.session_state: st.session_state.r_data[key.replace("_input", "").replace("_radio", "")] = st.session_state[key]
        if "adu_file_input" in st.session_state and st.session_state["adu_file_input"]: st.session_state.r_data["file_bukti"] = st.session_state["adu_file_input"]
        st.session_state.p_step = "DATA_KONTAK"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "SARAN_FORM":
    render_brand_header(); top_back_button("LANDING")
    st.markdown("<h1 class='anim-2 gradient-text' style='font-size: 26px; font-weight: 800; margin-bottom: 6px;'>Punya ide untuk pelayanan lebih baik?</h1><p class='anim-2' style='color: #64748b; font-size: 15px; margin-bottom: 28px;'>Setiap saran Anda sangat berarti bagi kami.</p>", unsafe_allow_html=True)

    st.markdown("<div class='anim-3 glass-box'>", unsafe_allow_html=True)
    opsi = ["Pelayanan", "Fasilitas", "Administrasi", "BPJS/JKN", "Antrean", "Farmasi", "Laboratorium", "Kenyamanan", "Digital / website / WhatsApp", "Lainnya"]
    st.session_state.r_data["active_masalah"] = st.radio("Topik Saran:", opsi, index=opsi.index(get_val("active_masalah")) if get_val("active_masalah") in opsi else None, key="sar_opt_masalah")
    st.session_state.r_data["SAR_cerita"] = st.text_area("Ceritakan saran Anda", height=100, value=get_val("SAR_cerita"), key="sar_cerita_input")
    st.session_state.r_data["SAR_perbaiki"] = st.text_area("Apa yang sebaiknya diperbaiki? (Opsional)", height=80, value=get_val("SAR_perbaiki"), key="sar_perbaiki_input")
    
    opsi_ptg = ["Biasa", "Penting", "Sangat penting"]
    st.session_state.r_data["SAR_penting"] = st.radio("Tingkat Penting", opsi_ptg, index=opsi_ptg.index(get_val("SAR_penting")) if get_val("SAR_penting") in opsi_ptg else None, key="sar_penting_radio")

    f = st.file_uploader("Attachment (Opsional)", type=['jpg','png','jpeg','pdf','mp4','mov','avi'], accept_multiple_files=True, key="sar_file_input")
    if f: st.session_state.r_data["file_bukti"] = f
    st.markdown("</div>", unsafe_allow_html=True)
    
    st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
    if st.button("Lanjut ➔", use_container_width=True): 
        st.session_state.r_data["active_masalah"] = st.session_state["sar_opt_masalah"]
        st.session_state.r_data["SAR_cerita"] = st.session_state["sar_cerita_input"]
        st.session_state.r_data["SAR_perbaiki"] = st.session_state["sar_perbaiki_input"]
        st.session_state.r_data["SAR_penting"] = st.session_state["sar_penting_radio"]
        if st.session_state["sar_file_input"]: st.session_state.r_data["file_bukti"] = st.session_state["sar_file_input"]
        st.session_state.p_step = "DATA_KONTAK"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "APRES_FORM":
    render_brand_header(); top_back_button("LANDING")
    st.markdown("<h1 class='anim-2 gradient-text' style='font-size: 26px; font-weight: 800; margin-bottom: 6px;'>Terima kasih sudah berbagi pengalaman.</h1><p class='anim-2' style='color: #64748b; font-size: 15px; margin-bottom: 28px;'>Apresiasi Anda adalah semangat kami.</p>", unsafe_allow_html=True)

    st.markdown("<div class='anim-3 glass-box'>", unsafe_allow_html=True)
    opsi = ["Dokter", "Perawat", "Bidan", "Admin/Pendaftaran", "Farmasi", "Laboratorium", "Petugas lainnya", "Tim Klinik secara umum"]
    st.session_state.r_data["active_masalah"] = st.radio("Siapa/Layanan apa yang diapresiasi?", opsi, index=opsi.index(get_val("active_masalah")) if get_val("active_masalah") in opsi else None, key="apr_opt_masalah")
    st.session_state.r_data["APR_bantu"] = st.text_area("Apa yang membuat Anda merasa terbantu?", height=120, value=get_val("APR_bantu"), key="apr_bantu_input")
    st.session_state.r_data["APR_suka"] = st.text_area("Apa yang paling Anda sukai dari pelayanan kami? (Opsional)", height=80, value=get_val("APR_suka"), key="apr_suka_input")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
    if st.button("Lanjut ➔", use_container_width=True): 
        st.session_state.r_data["active_masalah"] = st.session_state["apr_opt_masalah"]
        st.session_state.r_data["APR_bantu"] = st.session_state["apr_bantu_input"]
        st.session_state.r_data["APR_suka"] = st.session_state["apr_suka_input"]
        st.session_state.p_step = "DATA_KONTAK"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "BPJS_MENU":
    render_brand_header(); top_back_button("LANDING")
    st.markdown("<h1 class='anim-2 gradient-text' style='font-size: 30px; font-weight: 800; margin-bottom: 8px;'>Bantuan BPJS / JKN</h1><p class='anim-2' style='color: #475569; font-size: 16px; margin-bottom: 32px;'>Silakan pilih kebutuhan Anda.</p>", unsafe_allow_html=True)

    st.markdown('<div class="anim-3">', unsafe_allow_html=True)
    if st.button("🩺 **Rujukan**\n\nSurat rujukan, status, prosedur", use_container_width=True, type="secondary"):
        st.session_state.r_data = {"active_module": "RUJUKAN"}; st.session_state.p_step = "RUJUKAN_PILIH"; st.rerun()
    if st.button("🎫 **Antrean**\n\nCara daftar, antrean klinik, MJKN", use_container_width=True, type="secondary"):
        st.session_state.r_data = {"active_module": "ANTREAN"}; st.session_state.p_step = "ANTREAN_PILIH"; st.rerun()
    if st.button("📱 **Mobile JKN**\n\nLogin, antrean, data, dll.", use_container_width=True, type="secondary"):
        st.session_state.r_data = {"active_module": "MJKN"}; st.session_state.p_step = "MJKN_PILIH"; st.rerun()
    if st.button("💊 **Obat**\n\nKetersediaan, resep, pengambilan", use_container_width=True, type="secondary"):
        st.session_state.r_data = {"active_module": "OBAT"}; st.session_state.p_step = "OBAT_PILIH"; st.rerun()
    if st.button("👤 **Kepesertaan**\n\nStatus, FKTP, perubahan data", use_container_width=True, type="secondary"):
        st.session_state.r_data = {"active_module": "KEPESERTAAN"}; st.session_state.p_step = "KEPESERTAAN_PILIH"; st.rerun()
    if st.button("🏥 **Layanan JKN**\n\nProsedur, hak peserta, biaya", use_container_width=True, type="secondary"):
        st.session_state.r_data = {"active_module": "LAYANAN"}; st.session_state.p_step = "LAYANAN_PILIH"; st.rerun()
    if st.button("❓ **Lainnya**\n\nJika tidak menemukan pilihan di atas", use_container_width=True, type="secondary"):
        st.session_state.r_data = {"active_module": "LAINNYA"}; st.session_state.p_step = "LAINNYA_PILIH"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------------------------------------------------
# FITUR RUJUKAN (PROTECTED - TIDAK DIUBAH SAMA SEKALI)
# ---------------------------------------------------------
elif st.session_state.p_step == "RUJUKAN_PILIH":
    render_brand_header(); top_back_button("BPJS_MENU")
    st.markdown("<h1 class='anim-2 gradient-text' style='font-size: 26px; font-weight: 800; margin-bottom: 6px;'>Bantuan Rujukan</h1><p class='anim-2' style='color: #64748b; font-size: 15px; margin-bottom: 28px;'>Pilih kendala yang paling sesuai dengan kondisi Anda.</p>", unsafe_allow_html=True)
    opsi = ["A. Saya ingin mendapatkan rujukan", "B. Rujukan saya belum bisa dibuat", "C. Rujukan saya ditolak / tidak bisa diproses", "D. Saya ingin mengetahui status rujukan", "E. Saya bingung mengenai rumah sakit tujuan", "F. Rujukan saya sudah ada tetapi bermasalah", "G. Masalah rujukan di Mobile JKN", "H. Rujukan untuk kontrol lanjutan", "I. Rujukan saya sudah kedaluwarsa / perlu diperbarui", "J. Lainnya"]
    
    st.markdown('<div class="anim-3">', unsafe_allow_html=True)
    for op in opsi:
        if st.button(f"**{op}**", use_container_width=True, key=f"btn_ruj_{op}"):
            st.session_state.r_data["active_masalah"] = op; st.session_state.p_step = "RUJUKAN_FORM"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "RUJUKAN_FORM":
    render_brand_header(); top_back_button("RUJUKAN_PILIH")
    masalah_full = get_val("active_masalah", "J. Lainnya")
    kode = masalah_full[0] if masalah_full else "J"
    masalah = masalah_full[3:] if masalah_full else "Lainnya"
    
    st.markdown(f"<h1 class='anim-2 gradient-text' style='font-size: 22px; font-weight: 800; margin-bottom: 24px; line-height:1.4;'>{masalah}</h1>", unsafe_allow_html=True)
    st.markdown("<div class='anim-3 glass-box'>", unsafe_allow_html=True)

    if kode == "A":
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px;'>Apakah Anda sudah diperiksa dokter di klinik ini?</p>", unsafe_allow_html=True)
        opsi_periksa = ["Ya", "Belum", "Tidak yakin"]
        idx = opsi_periksa.index(get_val('A_periksa')) if get_val('A_periksa') in opsi_periksa else None
        diperiksa = st.radio("Tanya Periksa", opsi_periksa, index=idx, key="unik_ruj_A_periksa", label_visibility="collapsed")
        if st.session_state.get("unik_ruj_A_periksa") == "Ya":
            st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Tanggal pemeriksaan</p>", unsafe_allow_html=True)
            tgl = st.date_input("Tanggal Periksa", key="unik_ruj_A_tgl", label_visibility="collapsed")
            st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Poli / layanan</p>", unsafe_allow_html=True)
            st.text_input("Poli Layanan", value=get_val("A_poli"), key="unik_ruj_A_poli", label_visibility="collapsed")
            st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Dokter (jika diketahui)</p>", unsafe_allow_html=True)
            st.text_input("Nama Dokter", value=get_val("A_dok"), key="unik_ruj_A_dok", label_visibility="collapsed")
            st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Apa tujuan rujukan Anda?</p>", unsafe_allow_html=True)
            opsi_tujuan = ["Dokter spesialis", "Pemeriksaan lanjutan", "Rumah sakit", "Pemeriksaan penunjang", "Lainnya"]
            idx_t = opsi_tujuan.index(get_val('A_tujuan')) if get_val('A_tujuan') in opsi_tujuan else None
            tujuan = st.radio("Tujuan Rujukan", opsi_tujuan, index=idx_t, key="unik_ruj_A_tujuan", label_visibility="collapsed")
            if st.session_state.get("unik_ruj_A_tujuan"):
                st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Apa yang ingin Anda tanyakan?</p>", unsafe_allow_html=True)
                opsi_tanya = ["Bagaimana proses rujukan?", "Kapan rujukan bisa dibuat?", "Rumah sakit tujuan", "Poli spesialis tujuan", "Persyaratan", "Lainnya"]
                idx_ta = opsi_tanya.index(get_val('A_tanya')) if get_val('A_tanya') in opsi_tanya else None
                tanya = st.radio("Tanya Rujukan", opsi_tanya, index=idx_ta, key="unik_ruj_A_tanya", label_visibility="collapsed")
                if st.session_state.get("unik_ruj_A_tanya") == "Lainnya":
                    st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Tuliskan pertanyaan Anda</p>", unsafe_allow_html=True)
                    st.text_area("Tulis Tanya Lainnya", value=get_val("A_tanya_lain"), key="unik_ruj_A_tanya_lain", label_visibility="collapsed")
    elif kode == "B":
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px;'>Apa yang terjadi?</p>", unsafe_allow_html=True)
        opsi = ["Dokter belum membuat rujukan", "Rujukan belum muncul", "Sistem tidak bisa membuat rujukan", "Data peserta bermasalah", "Rumah sakit tujuan tidak tersedia", "Saya tidak tahu penyebabnya", "Lainnya"]
        idx = opsi.index(get_val('B_kejadian')) if get_val('B_kejadian') in opsi else None
        st.radio("Apa Terjadi B", opsi, index=idx, key="unik_ruj_B_kejadian", label_visibility="collapsed")
        if st.session_state.get("unik_ruj_B_kejadian") == "Sistem tidak bisa membuat rujukan":
            st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Apa pesan/error yang muncul?</p>", unsafe_allow_html=True)
            st.text_area("Pesan Error B", value=get_val("B_error"), key="unik_ruj_B_error", label_visibility="collapsed")
            st.markdown("<p style='font-size:14px; margin-bottom:5px; margin-top:15px; font-weight:600;'>📱 Tambahkan screenshot (Optional)</p>", unsafe_allow_html=True)
            f = st.file_uploader("Upload B", type=['jpg','png','jpeg','pdf','mp4','mov','avi'], accept_multiple_files=True, key="unik_ruj_B_file", label_visibility="collapsed")
            if f: st.session_state.r_data["file_bukti"] = f
    elif kode == "C":
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px;'>Kapan rujukan dibuat?</p>", unsafe_allow_html=True)
        st.text_input("Kapan Dibuat C", value=get_val("C_kapan"), key="unik_ruj_C_kapan", label_visibility="collapsed")
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Di mana rujukan dibuat?</p>", unsafe_allow_html=True)
        st.text_input("Dimana Dibuat C", value=get_val("C_dimana"), key="unik_ruj_C_dimana", label_visibility="collapsed")
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Apa alasan penolakan yang Anda terima?</p>", unsafe_allow_html=True)
        st.text_area("Alasan Tolak C", value=get_val("C_alasan"), key="unik_ruj_C_alasan", label_visibility="collapsed")
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Apakah muncul pesan/error? (Tuliskan jika ada)</p>", unsafe_allow_html=True)
        st.text_input("Error C", value=get_val("C_error"), key="unik_ruj_C_error", label_visibility="collapsed")
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Apakah memiliki bukti?</p>", unsafe_allow_html=True)
        opsi_bukti = ["Ya", "Tidak"]
        idx = opsi_bukti.index(get_val('C_bukti')) if get_val('C_bukti') in opsi_bukti else None
        st.radio("Punya Bukti C", opsi_bukti, index=idx, key="unik_ruj_C_bukti", label_visibility="collapsed")
        if st.session_state.get("unik_ruj_C_bukti") == "Ya":
            st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Upload Screenshot/Foto/Dokumen</p>", unsafe_allow_html=True)
            f = st.file_uploader("Upload C", type=['jpg','png','pdf','mp4','mov','avi'], accept_multiple_files=True, key="unik_ruj_C_file", label_visibility="collapsed")
            if f: st.session_state.r_data["file_bukti"] = f
    elif kode == "D":
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px;'>Kapan rujukan dibuat?</p>", unsafe_allow_html=True)
        st.text_input("Kapan D", value=get_val("D_kapan"), key="unik_ruj_D_kapan", label_visibility="collapsed")
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Tujuan rumah sakit jika diketahui</p>", unsafe_allow_html=True)
        st.text_input("RS D", value=get_val("D_rs"), key="unik_ruj_D_rs", label_visibility="collapsed")
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Poli/spesialis jika diketahui</p>", unsafe_allow_html=True)
        st.text_input("Poli D", value=get_val("D_poli"), key="unik_ruj_D_poli", label_visibility="collapsed")
        if st.session_state.get("unik_ruj_D_kapan"):
            st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Apa yang ingin Anda ketahui?</p>", unsafe_allow_html=True)
            opsi = ["Apakah rujukan sudah aktif?", "Apakah rujukan sudah dapat digunakan?", "Bagaimana cara melihat rujukan?", "Rujukan belum muncul di Mobile JKN", "Lainnya"]
            idx = opsi.index(get_val('D_tanya')) if get_val('D_tanya') in opsi else None
            st.radio("Tanya D", opsi, index=idx, key="unik_ruj_D_tanya", label_visibility="collapsed")
    elif kode == "E":
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px;'>Spesialisasi apa yang Anda perlukan?</p>", unsafe_allow_html=True)
        st.text_input("Spes E", value=get_val("E_spes"), key="unik_ruj_E_spes", label_visibility="collapsed")
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Apakah sudah ada arahan dokter?</p>", unsafe_allow_html=True)
        opsi_arah = ["Ya", "Belum", "Tidak yakin"]
        idx_a = opsi_arah.index(get_val('E_arah')) if get_val('E_arah') in opsi_arah else None
        st.radio("Arahan E", opsi_arah, index=idx_a, key="unik_ruj_E_arah", label_visibility="collapsed")
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Apakah sudah memiliki rumah sakit tujuan?</p>", unsafe_allow_html=True)
        opsi_rs = ["Ya", "Belum"]
        idx_rs = opsi_rs.index(get_val('E_rs')) if get_val('E_rs') in opsi_rs else None
        st.radio("Punya RS E", opsi_rs, index=idx_rs, key="unik_ruj_E_rs", label_visibility="collapsed")
        if st.session_state.get("unik_ruj_E_rs") == "Belum":
            st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Apa yang ingin dibantu?</p>", unsafe_allow_html=True)
            st.text_area("Bantu E", value=get_val("E_bantu"), key="unik_ruj_E_bantu", label_visibility="collapsed")
    elif kode == "F":
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px;'>Pilih masalah rujukan Anda:</p>", unsafe_allow_html=True)
        opsi = ["Data salah", "Tujuan salah", "Tanggal salah", "Tidak muncul di Mobile JKN", "Tidak dapat digunakan", "Rumah sakit tidak menerima", "Masa berlaku bermasalah", "Lainnya"]
        idx = opsi.index(get_val('F_masalah')) if get_val('F_masalah') in opsi else None
        st.radio("Masalah F", opsi, index=idx, key="unik_ruj_F_masalah", label_visibility="collapsed")
        masalah_f = st.session_state.get("unik_ruj_F_masalah")
        if masalah_f == "Data salah":
            st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Data apa yang salah dan apa yang seharusnya?</p>", unsafe_allow_html=True)
            st.text_area("Detail F Salah", value=get_val("F_detail"), key="unik_ruj_F_detail1", label_visibility="collapsed")
        elif masalah_f == "Tujuan salah":
            st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Apa nama rumah sakit tujuan yang seharusnya?</p>", unsafe_allow_html=True)
            st.text_input("Detail F RS", value=get_val("F_detail"), key="unik_ruj_F_detail2", label_visibility="collapsed")
        elif masalah_f == "Tanggal salah":
            st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Tanggal berapa rujukan seharusnya?</p>", unsafe_allow_html=True)
            st.date_input("Tanggal F", key="unik_ruj_F_tgl", label_visibility="collapsed")
        elif masalah_f:
            st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Jelaskan detail masalahnya</p>", unsafe_allow_html=True)
            st.text_area("Detail F Umum", value=get_val("F_detail"), key="unik_ruj_F_detail3", label_visibility="collapsed")
    elif kode == "G":
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px;'>Apakah rujukan terlihat di Mobile JKN?</p>", unsafe_allow_html=True)
        opsi = ["Ya", "Tidak", "Tidak tahu"]
        idx = opsi.index(get_val('G_lihat')) if get_val('G_lihat') in opsi else None
        st.radio("Lihat G", opsi, index=idx, key="unik_ruj_G_lihat", label_visibility="collapsed")
        if st.session_state.get("unik_ruj_G_lihat") == "Tidak":
            st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Apa yang muncul di aplikasi?</p>", unsafe_allow_html=True)
            st.text_area("Muncul G", value=get_val("G_muncul"), key="unik_ruj_G_muncul", label_visibility="collapsed")
            st.markdown("<p style='font-size:14px; margin-bottom:5px; margin-top:15px; font-weight:600;'>📱 Tambahkan screenshot (Opsional)</p>", unsafe_allow_html=True)
            f = st.file_uploader("Upload G", type=['jpg','png','jpeg','mp4','mov','avi','pdf'], accept_multiple_files=True, key="unik_ruj_G_file", label_visibility="collapsed")
            if f: st.session_state.r_data["file_bukti"] = f
    elif kode == "H":
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px;'>Rumah sakit tujuan</p>", unsafe_allow_html=True)
        st.text_input("RS H", value=get_val("H_rs"), key="unik_ruj_H_rs", label_visibility="collapsed")
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Poli</p>", unsafe_allow_html=True)
        st.text_input("Poli H", value=get_val("H_poli"), key="unik_ruj_H_poli", label_visibility="collapsed")
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Tanggal kontrol jika diketahui</p>", unsafe_allow_html=True)
        st.date_input("Tanggal H", key="unik_ruj_H_tgl", label_visibility="collapsed")
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Apakah memiliki surat/keterangan kontrol?</p>", unsafe_allow_html=True)
        opsi = ["Ya", "Tidak", "Tidak yakin"]
        idx = opsi.index(get_val('H_surat')) if get_val('H_surat') in opsi else None
        st.radio("Surat H", opsi, index=idx, key="unik_ruj_H_surat", label_visibility="collapsed")
        if st.session_state.get("unik_ruj_H_surat") == "Ya":
            st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Apa kendalanya?</p>", unsafe_allow_html=True)
            st.text_area("Kendala H", value=get_val("H_kendala"), key="unik_ruj_H_kendala", label_visibility="collapsed")
    elif kode == "I":
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px;'>Kapan rujukan dibuat?</p>", unsafe_allow_html=True)
        st.text_input("Kapan I", value=get_val("I_kapan"), key="unik_ruj_I_kapan", label_visibility="collapsed")
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Apakah sudah pernah digunakan?</p>", unsafe_allow_html=True)
        opsi = ["Ya", "Belum"]
        idx = opsi.index(get_val('I_pakai')) if get_val('I_pakai') in opsi else None
        st.radio("Pakai I", opsi, index=idx, key="unik_ruj_I_pakai", label_visibility="collapsed")
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Kapan rencana kunjungan?</p>", unsafe_allow_html=True)
        st.text_input("Rencana I", value=get_val("I_rencana"), key="unik_ruj_I_rencana", label_visibility="collapsed")
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px; margin-top:20px;'>Apa yang ingin dibantu?</p>", unsafe_allow_html=True)
        st.text_area("Bantu I", value=get_val("I_bantu"), key="unik_ruj_I_bantu", label_visibility="collapsed")
    elif kode == "J":
        st.markdown("<p style='font-size:15px; font-weight:700; color:#0f172a; margin-bottom:10px;'>Ceritakan masalah atau pertanyaan Anda</p>", unsafe_allow_html=True)
        st.text_area("Cerita J", placeholder="Contoh:\nSaya ingin menanyakan...\nSaya mengalami kendala...", height=150, value=get_val("J_cerita"), key="unik_ruj_J_cerita", label_visibility="collapsed")

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
    if st.button("Lanjut ➔", use_container_width=True, key="btn_ruj_form_next"): 
        for k in list(st.session_state.keys()):
            if k.startswith("unik_ruj_"):
                val = st.session_state[k]
                if "tgl" in k and val: val = val.strftime("%d %B %Y")
                st.session_state.r_data[k.replace("unik_ruj_", "")] = val
        st.session_state.p_step = "DATA_KONTAK"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------------------------------------------------
# SISA MENU: ANTREAN, MJKN, OBAT, KEPESERTAAN, LAYANAN, LAINNYA
# ---------------------------------------------------------
elif st.session_state.p_step == "ANTREAN_PILIH":
    render_brand_header(); top_back_button("BPJS_MENU")
    st.markdown("<h1 class='anim-2 gradient-text' style='font-size: 26px; font-weight: 800; margin-bottom: 4px;'>Antrean BPJS/JKN</h1><p class='anim-2' style='color: #64748b; font-size: 15px; margin-bottom: 28px;'>Ada kendala atau pertanyaan terkait antrean?</p>", unsafe_allow_html=True)
    opsi = ["A. Saya tidak tahu cara mengambil antrean", "B. Saya tidak bisa mengambil antrean", "C. Antrean penuh", "D. Nomor antrean tidak muncul", "E. Nomor antrean hilang / berubah", "F. Masalah antrean Mobile JKN", "G. Saya bingung dengan jadwal", "H. Saya sudah mengambil antrean tetapi mengalami masalah", "I. Masalah antrean di klinik", "J. Lainnya"]
    st.markdown('<div class="anim-3">', unsafe_allow_html=True)
    for op in opsi:
        if st.button(f"**{op}**", use_container_width=True, key=f"btn_ant_{op}"):
            st.session_state.r_data["active_masalah"] = op
            if "Mobile JKN" in op:
                st.session_state.r_data["active_masalah"] = "E. Tidak bisa mengambil antrean"; st.session_state.r_data["active_module"] = "MJKN"; st.session_state.p_step = "MJKN_FORM"
            else: st.session_state.p_step = "ANTREAN_FORM"
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "ANTREAN_FORM":
    render_brand_header(); top_back_button("ANTREAN_PILIH")
    kode = get_val("active_masalah", "J")[0]; masalah = get_val("active_masalah", "J. Lainnya")[3:]
    st.markdown(f"<h1 class='anim-2 gradient-text' style='font-size: 20px; font-weight: 700; margin-bottom: 24px; line-height:1.4;'>{masalah}</h1>", unsafe_allow_html=True)
    st.markdown("<div class='anim-3 glass-box'>", unsafe_allow_html=True)
    if kode == "A":
        opsi = ["Berobat", "Kontrol", "Poli tertentu", "Lainnya"]
        st.radio("Perlu Antrean A", opsi, index=opsi.index(get_val('A_perlu')) if get_val('A_perlu') in opsi else None, key="unik_ant_A_perlu")
        if st.session_state.get("unik_ant_A_perlu"):
            opsi_mjkn = ["Ya", "Tidak", "Tidak tahu"]
            mjkn = st.radio("Pakai MJKN A", opsi_mjkn, index=opsi_mjkn.index(get_val('A_mjkn')) if get_val('A_mjkn') in opsi_mjkn else None, key="unik_ant_A_mjkn")
            if mjkn == "Ya": st.info("💡 Buka aplikasi Mobile JKN > Pendaftaran Pelayanan > faskes Klinik Pratama Papan Mas > Poli.")
            elif mjkn == "Tidak": st.info("💡 Anda dapat datang langsung ke mesin antrean klinik.")
            if mjkn: st.text_area("Bantuan A", value=get_val("A_bantu"), key="unik_ant_A_bantu")
    elif kode == "B":
        st.radio("MJKN B", ["Ya", "Tidak"], index=["Ya", "Tidak"].index(get_val('B_mjkn')) if get_val('B_mjkn') in ["Ya", "Tidak"] else None, key="unik_ant_B_mjkn")
        st.date_input("Tanggal B", key="unik_ant_B_tgl")
        st.text_input("Poli B", value=get_val("B_poli"), key="unik_ant_B_poli")
        st.text_area("Terjadi B", value=get_val("B_terjadi"), key="unik_ant_B_terjadi")
        st.text_input("Error B", value=get_val("B_err"), key="unik_ant_B_err")
        f = st.file_uploader("Upload B", type=['jpg','png','jpeg','pdf','mp4','mov','avi'], accept_multiple_files=True, key="unik_ant_B_file")
        if f: st.session_state.r_data["file_bukti"] = f
    elif kode == "C":
        st.date_input("Tanggal C", key="unik_ant_C_tgl")
        st.text_input("Poli C", value=get_val("C_poli"), key="unik_ant_C_poli")
        st.text_input("Jam C", value=get_val("C_jam"), key="unik_ant_C_jam")
        st.radio("Hari C", ["Hari ini", "Tanggal lain"], index=None, key="unik_ant_C_hari")
        f = st.file_uploader("Upload C", type=['jpg','png','jpeg','pdf','mp4','mov','avi'], accept_multiple_files=True, key="unik_ant_C_file")
        if f: st.session_state.r_data["file_bukti"] = f
    elif kode == "D":
        st.date_input("Tanggal D", key="unik_ant_D_tgl")
        st.text_input("Poli D", value=get_val("D_poli"), key="unik_ant_D_poli")
        st.text_input("Jam D", value=get_val("D_jam"), key="unik_ant_D_jam")
        st.radio("Sukses D", ["Ya", "Tidak", "Tidak yakin"], index=None, key="unik_ant_D_sukses")
        st.text_input("Sebelum D", value=get_val("D_sebelum"), key="unik_ant_D_sebelum")
    elif kode == "E":
        st.text_input("Lama E", value=get_val("E_no_lama"), key="unik_ant_E_no_lama")
        st.date_input("Tanggal E", key="unik_ant_E_tgl")
        st.text_input("Poli E", value=get_val("E_poli"), key="unik_ant_E_poli")
        st.text_input("Baru E", value=get_val("E_no_baru"), key="unik_ant_E_no_baru")
    elif kode == "G":
        opsi = ["Jadwal dokter", "Jadwal poli", "Jadwal berubah", "Jadwal tidak muncul", "Lainnya"]
        st.radio("Kendala G", opsi, index=opsi.index(get_val('G_kendala')) if get_val('G_kendala') in opsi else None, key="unik_ant_G_kendala")
        st.text_input("Poli G", value=get_val("G_poli"), key="unik_ant_G_poli")
        st.date_input("Tanggal G", key="unik_ant_G_tgl")
        st.text_input("Dokter G", value=get_val("G_dok"), key="unik_ant_G_dok")
    elif kode == "H":
        st.date_input("Tanggal H", key="unik_ant_H_tgl")
        st.text_input("Poli H", value=get_val("H_poli"), key="unik_ant_H_poli")
        st.text_input("Nomor H", value=get_val("H_no"), key="unik_ant_H_no")
        st.text_input("Jam H", value=get_val("H_jam"), key="unik_ant_H_jam")
        st.text_area("Masalah H", value=get_val("H_masalah"), key="unik_ant_H_masalah")
    elif kode == "I":
        st.text_input("Lokasi I", value=get_val("I_lok"), key="unik_ant_I_lok")
        st.date_input("Tanggal I", key="unik_ant_I_tgl")
        st.text_input("Jam I", value=get_val("I_jam"), key="unik_ant_I_jam")
        st.text_input("Nomor I", value=get_val("I_no"), key="unik_ant_I_no")
        st.text_area("Masalah I", value=get_val("I_masalah"), key="unik_ant_I_masalah")
    elif kode == "J":
        st.text_area("Cerita J", placeholder="Ceritakan...", value=get_val("J_cerita"), key="unik_ant_J_cerita")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
    if st.button("Lanjut ➔", use_container_width=True): 
        for k in list(st.session_state.keys()):
            if k.startswith("unik_ant_"):
                val = st.session_state[k]
                if "tgl" in k and val: val = val.strftime("%d %B %Y")
                st.session_state.r_data[k.replace("unik_ant_", "")] = val
        st.session_state.p_step = "DATA_KONTAK"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "MJKN_PILIH":
    render_brand_header(); top_back_button("BPJS_MENU")
    st.markdown("<h1 class='anim-2 gradient-text' style='font-size: 26px; font-weight: 800; margin-bottom: 4px;'>Bantuan Mobile JKN</h1><p class='anim-2' style='color: #64748b; font-size: 15px; margin-bottom: 28px;'>Pilih kendala yang Anda alami saat menggunakan Mobile JKN.</p>", unsafe_allow_html=True)
    opsi = ["A. Tidak bisa login", "B. Lupa password", "C. Akun bermasalah", "D. Data peserta tidak sesuai", "E. Tidak bisa mengambil antrean", "F. Antrean tidak muncul", "G. Rujukan tidak muncul", "H. Status kepesertaan", "I. FKTP", "J. Fitur lain di Mobile JKN", "K. Saya hanya ingin tahu cara menggunakan Mobile JKN", "L. Lainnya"]
    st.markdown('<div class="anim-3">', unsafe_allow_html=True)
    for op in opsi:
        if st.button(f"**{op}**", use_container_width=True, key=f"btn_mjkn_{op}"):
            st.session_state.r_data["active_masalah"] = op
            if op.startswith("E.") or op.startswith("F."):
                st.session_state.r_data["active_masalah"] = "F. Masalah antrean Mobile JKN"; st.session_state.r_data["active_module"] = "ANTREAN"; st.session_state.p_step = "ANTREAN_FORM"
            elif op.startswith("G."):
                st.session_state.r_data["active_masalah"] = "G. Masalah rujukan di Mobile JKN"; st.session_state.r_data["active_module"] = "RUJUKAN"; st.session_state.p_step = "RUJUKAN_FORM"
            else: st.session_state.p_step = "MJKN_FORM"
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "MJKN_FORM":
    render_brand_header(); top_back_button("MJKN_PILIH")
    kode = get_val("active_masalah", "L")[0]; masalah = get_val("active_masalah", "L. Lainnya")[3:]
    st.markdown(f"<h1 class='anim-2 gradient-text' style='font-size: 22px; font-weight: 800; margin-bottom: 24px; line-height:1.4;'>{masalah}</h1>", unsafe_allow_html=True)
    st.markdown("<div class='anim-3 glass-box'>", unsafe_allow_html=True)
    if kode == "A":
        st.radio("Buka A", ["Ya", "Tidak"], index=None, key="unik_mkn_A_buka")
        st.radio("User A", ["Ya", "Tidak"], index=None, key="unik_mkn_A_user")
        st.text_area("Jeda A", value=get_val("A_jeda"), key="unik_mkn_A_jeda")
        st.text_input("Err A", value=get_val("A_err"), key="unik_mkn_A_err")
        f = st.file_uploader("Upload A", type=['jpg','png','jpeg','pdf','mp4','mov','avi'], accept_multiple_files=True, key="unik_mkn_A_file")
        if f: st.session_state.r_data["file_bukti"] = f
    elif kode == "B":
        st.radio("Coba B", ["Ya", "Belum"], index=None, key="unik_mkn_B_coba")
        st.text_area("Jadi B", value=get_val("B_jadi"), key="unik_mkn_B_jadi")
        st.radio("Aktif B", ["Ya", "Tidak", "Tidak Tahu"], index=None, key="unik_mkn_B_aktif")
        st.text_input("Err B", value=get_val("B_err"), key="unik_mkn_B_err")
    elif kode == "C":
        opsi = ["Akun terkunci", "Nomor HP berubah", "Data tidak sesuai", "Tidak dapat masuk", "Masalah lain"]
        akun = st.radio("Akun C", opsi, index=opsi.index(get_val('C_akun')) if get_val('C_akun') in opsi else None, key="unik_mkn_C_akun")
        if st.session_state.get("unik_mkn_C_akun") == "Masalah lain": st.text_area("Lain C", value=get_val("C_lain"), key="unik_mkn_C_lain")
    elif kode == "D":
        opsi = ["Nama", "NIK", "Tanggal lahir", "Anggota keluarga", "FKTP", "Data lainnya"]
        data = st.radio("Data D", opsi, index=opsi.index(get_val('D_data')) if get_val('D_data') in opsi else None, key="unik_mkn_D_data")
        if st.session_state.get("unik_mkn_D_data"): st.text_area("Detail D", value=get_val("D_detail"), key="unik_mkn_D_detail")
    elif kode == "H":
        st.text_input("Status H", value=get_val("H_status"), key="unik_mkn_H_status")
        st.text_input("Seharusnya H", value=get_val("H_seharusnya"), key="unik_mkn_H_seharusnya")
        f = st.file_uploader("Upload H", type=['jpg','png','jpeg','pdf','mp4','mov','avi'], accept_multiple_files=True, key="unik_mkn_H_file")
        if f: st.session_state.r_data["file_bukti"] = f
    elif kode == "I":
        opsi = ["FKTP tidak sesuai", "Ingin mengetahui FKTP", "Ingin mengetahui prosedur perubahan FKTP", "Masalah lain"]
        st.radio("FKTP I", opsi, index=opsi.index(get_val('I_fktp')) if get_val('I_fktp') in opsi else None, key="unik_mkn_I_fktp")
        if st.session_state.get("unik_mkn_I_fktp") == "Ingin mengetahui prosedur perubahan FKTP": st.info("💡 Perubahan FKTP dapat dilakukan di aplikasi Mobile JKN menu 'Perubahan Data Peserta'.")
        elif st.session_state.get("unik_mkn_I_fktp") == "Masalah lain": st.text_area("Lain I", value=get_val("I_lain"), key="unik_mkn_I_lain")
    elif kode == "J":
        st.text_area("Cerita J", value=get_val("J_cerita"), key="unik_mkn_J_cerita")
    elif kode == "K":
        opsi = ["Melihat kepesertaan", "Mengambil antrean", "Melihat rujukan", "Melihat FKTP", "Fitur lainnya"]
        st.radio("Tanya K", opsi, index=opsi.index(get_val('K_tanya')) if get_val('K_tanya') in opsi else None, key="unik_mkn_K_tanya")
        if st.session_state.get("unik_mkn_K_tanya") == "Mengambil antrean": st.info("💡 Buka aplikasi > Pendaftaran Pelayanan > Pilih Faskes > Poli.")
        elif st.session_state.get("unik_mkn_K_tanya") == "Melihat rujukan": st.info("💡 Buka aplikasi > Info Riwayat Pelayanan > Rujukan.")
        if st.session_state.get("unik_mkn_K_tanya"): st.text_area("Lain K", value=get_val("K_lain"), key="unik_mkn_K_lain")
    elif kode == "L":
        st.text_area("Cerita L", value=get_val("L_cerita"), key="unik_mkn_L_cerita")
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
    if st.button("Lanjut ➔", use_container_width=True): 
        for k in list(st.session_state.keys()):
            if k.startswith("unik_mkn_"): st.session_state.r_data[k.replace("unik_mkn_", "")] = st.session_state[k]
        st.session_state.p_step = "DATA_KONTAK"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "OBAT_PILIH":
    render_brand_header(); top_back_button("BPJS_MENU")
    st.markdown("<h1 class='anim-2 gradient-text' style='font-size: 26px; font-weight: 800; margin-bottom: 4px;'>Bantuan Obat</h1><p class='anim-2' style='color: #64748b; font-size: 15px; margin-bottom: 28px;'>Pilih kendala atau pertanyaan terkait obat.</p>", unsafe_allow_html=True)
    opsi = ["A. Obat tidak tersedia", "B. Obat tidak diberikan", "C. Jumlah obat kurang", "D. Obat berbeda dengan resep", "E. Obat berbeda dari biasanya", "F. Obat PRB", "G. Masalah resep", "H. Saya ingin mengetahui cara penggunaan obat", "I. Pertanyaan pengambilan obat", "J. Lainnya"]
    st.markdown('<div class="anim-3">', unsafe_allow_html=True)
    for op in opsi:
        if st.button(f"**{op}**", use_container_width=True, key=f"btn_obt_{op}"):
            st.session_state.r_data["active_masalah"] = op; st.session_state.p_step = "OBAT_FORM"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "OBAT_FORM":
    render_brand_header(); top_back_button("OBAT_PILIH")
    kode = get_val("active_masalah", "J")[0]; masalah = get_val("active_masalah", "J. Lainnya")[3:]
    st.markdown(f"<h1 class='anim-2 gradient-text' style='font-size: 22px; font-weight: 800; margin-bottom: 24px; line-height:1.4;'>{masalah}</h1>", unsafe_allow_html=True)
    st.markdown("<div class='anim-3 glass-box'>", unsafe_allow_html=True)
    if kode == "A":
        st.date_input("Tanggal A", key="unik_obt_A_tgl")
        st.text_input("Poli A", value=get_val("O_poli"), key="unik_obt_O_poli")
        st.text_input("Nama Obat A", value=get_val("O_nama"), key="unik_obt_O_nama")
        opsi = ["Ya", "Tidak", "Tidak tahu"]
        st.radio("Ganti A", opsi, index=opsi.index(get_val('O_ganti')) if get_val('O_ganti') in opsi else None, key="unik_obt_O_ganti")
        st.text_area("Info A", value=get_val("O_info"), key="unik_obt_O_info")
    elif kode == "B":
        st.date_input("Tanggal B", key="unik_obt_B_tgl")
        st.text_input("Poli B", value=get_val("O_poli"), key="unik_obt_O_poli")
        opsi = ["Ya", "Belum"]
        st.radio("Resep B", opsi, index=opsi.index(get_val('O_resep')) if get_val('O_resep') in opsi else None, key="unik_obt_O_resep")
        st.text_area("Kejadian B", value=get_val("O_kejadian"), key="unik_obt_O_kejadian")
    elif kode == "C":
        st.text_input("Nama C", value=get_val("O_nama"), key="unik_obt_O_nama")
        st.text_input("Jml Harus C", value=get_val("O_jml_harus"), key="unik_obt_O_jml_harus")
        st.text_input("Jml Terima C", value=get_val("O_jml_terima"), key="unik_obt_O_jml_terima")
        f = st.file_uploader("Upload C", type=['jpg','png','jpeg','pdf','mp4','mov','avi'], accept_multiple_files=True, key="unik_obt_C_file")
        if f: st.session_state.r_data["file_bukti"] = f
    elif kode == "D":
        st.text_input("Nama Resep D", value=get_val("O_nama_resep"), key="unik_obt_O_nama_resep")
        st.text_input("Nama Terima D", value=get_val("O_nama_terima"), key="unik_obt_O_nama_terima")
        f = st.file_uploader("Upload D", type=['jpg','png','jpeg','pdf','mp4','mov','avi'], accept_multiple_files=True, key="unik_obt_D_file")
        if f: st.session_state.r_data["file_bukti"] = f
    elif kode == "E":
        st.date_input("Tanggal E", key="unik_obt_E_tgl")
        st.text_input("Nama Obat E", value=get_val("O_nama"), key="unik_obt_O_nama")
        st.text_area("Beda E", value=get_val("O_beda"), key="unik_obt_O_beda")
    elif kode == "F":
        opsi = ["Ya", "Tidak", "Tidak tahu"]
        st.radio("PRB F", opsi, index=opsi.index(get_val('O_prb_cek')) if get_val('O_prb_cek') in opsi else None, key="unik_obt_O_prb_cek")
        st.date_input("Tanggal F", key="unik_obt_F_tgl")
        st.text_input("Nama Obat F", value=get_val("O_nama"), key="unik_obt_O_nama")
        st.text_area("Masalah F", value=get_val("O_masalah"), key="unik_obt_O_masalah")
    elif kode == "G":
        opsi = ["Resep tidak terbaca", "Resep belum masuk", "Resep tidak sesuai", "Bingung dengan resep", "Lainnya"]
        st.radio("Kendala G", opsi, index=opsi.index(get_val('O_resep_kendala')) if get_val('O_resep_kendala') in opsi else None, key="unik_obt_O_resep_kendala")
    elif kode == "H":
        st.text_input("Nama Obat H", value=get_val("O_nama"), key="unik_obt_O_nama")
        opsi = ["Dosis", "Waktu minum", "Frekuensi", "Lama penggunaan", "Lainnya"]
        st.radio("Tanya H", opsi, index=opsi.index(get_val('O_tanya')) if get_val('O_tanya') in opsi else None, key="unik_obt_O_tanya")
    elif kode == "I":
        st.date_input("Tanggal I", key="unik_obt_I_tgl")
        st.text_input("Poli I", value=get_val("O_poli"), key="unik_obt_O_poli")
        st.text_input("Nama Obat I", value=get_val("O_nama"), key="unik_obt_O_nama")
        st.text_area("Tanya Bebas I", value=get_val("O_tanya_bebas"), key="unik_obt_O_tanya_bebas")
    elif kode == "J":
        st.text_area("Cerita J", value=get_val("O_cerita"), key="unik_obt_O_cerita")
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
    if st.button("Lanjut ➔", use_container_width=True): 
        for k in list(st.session_state.keys()):
            if k.startswith("unik_obt_"):
                val = st.session_state[k]
                if "tgl" in k and val: val = val.strftime("%d %B %Y")
                st.session_state.r_data[k.replace("unik_obt_", "")] = val
        st.session_state.p_step = "DATA_KONTAK"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "KEPESERTAAN_PILIH":
    render_brand_header(); top_back_button("BPJS_MENU")
    st.markdown("<h1 class='anim-2 gradient-text' style='font-size: 26px; font-weight: 800; margin-bottom: 4px;'>Masalah Kepesertaan BPJS/JKN</h1><p class='anim-2' style='color: #64748b; font-size: 15px; margin-bottom: 28px;'>Pilih kendala yang paling sesuai dengan kondisi Anda.</p>", unsafe_allow_html=True)
    opsi = ["A. Kepesertaan tidak aktif", "B. Data peserta tidak sesuai", "C. NIK tidak sesuai", "D. Nama/tanggal lahir tidak sesuai", "E. Masalah anggota keluarga", "F. FKTP tidak sesuai", "G. Perubahan FKTP", "H. Kartu BPJS/JKN bermasalah", "I. Data Mobile JKN berbeda", "J. Tidak tahu status kepesertaan", "K. Lainnya"]
    st.markdown('<div class="anim-3">', unsafe_allow_html=True)
    for op in opsi:
        if st.button(f"**{op}**", use_container_width=True, key=f"btn_kps_{op}"):
            st.session_state.r_data["active_masalah"] = op; st.session_state.p_step = "KEPESERTAAN_FORM"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "KEPESERTAAN_FORM":
    render_brand_header(); top_back_button("KEPESERTAAN_PILIH")
    kode = get_val("active_masalah", "K")[0]; masalah = get_val("active_masalah", "K. Lainnya")[3:]
    st.markdown(f"<h1 class='anim-2 gradient-text' style='font-size: 20px; font-weight: 700; margin-bottom: 24px; line-height:1.4;'>{masalah}</h1>", unsafe_allow_html=True)
    st.markdown("<div class='anim-3 glass-box'>", unsafe_allow_html=True)
    if kode == "A":
        st.text_input("Status A", value=get_val("KPS_status"), key="unik_kps_KPS_status")
        st.text_input("Kapan A", value=get_val("KPS_kapan"), key="unik_kps_KPS_kapan")
        opsi = ["Ya", "Tidak", "Tidak tahu"]
        st.radio("Aktif A", opsi, index=opsi.index(get_val('KPS_aktif')) if get_val('KPS_aktif') in opsi else None, key="unik_kps_KPS_aktif")
        f = st.file_uploader("Upload A", type=['jpg','png','jpeg','pdf','mp4','mov','avi'], accept_multiple_files=True, key="unik_kps_A_file")
        if f: st.session_state.r_data["file_bukti"] = f
    elif kode == "B":
        opsi = ["Nama", "NIK", "Tanggal lahir", "Alamat", "Anggota keluarga", "Lainnya"]
        st.radio("Data B", opsi, index=opsi.index(get_val('KPS_data')) if get_val('KPS_data') in opsi else None, key="unik_kps_KPS_data")
    elif kode == "C":
        st.text_input("NIK Asli C", value=get_val("KPS_nik_asli"), key="unik_kps_KPS_nik_asli")
    elif kode == "D":
        st.text_input("Tampil D", value=get_val("KPS_tampil"), key="unik_kps_KPS_tampil")
        st.text_input("Harus D", value=get_val("KPS_seharusnya"), key="unik_kps_KPS_seharusnya")
        f = st.file_uploader("Upload D", type=['jpg','png','jpeg','pdf','mp4','mov','avi'], accept_multiple_files=True, key="unik_kps_D_file")
        if f: st.session_state.r_data["file_bukti"] = f
    elif kode == "E":
        st.text_input("Keluarga E", value=get_val("KPS_siapa"), key="unik_kps_KPS_siapa")
        st.text_area("Masalah E", value=get_val("KPS_masalah"), key="unik_kps_KPS_masalah")
        st.text_input("Status Fam E", value=get_val("KPS_status_fam"), key="unik_kps_KPS_status_fam")
    elif kode == "F":
        st.text_input("Tampil F", value=get_val("KPS_fktp_tampil"), key="unik_kps_KPS_fktp_tampil")
        st.text_input("Harap F", value=get_val("KPS_fktp_harap"), key="unik_kps_KPS_fktp_harap")
        st.text_area("Masalah F", value=get_val("KPS_masalah"), key="unik_kps_KPS_masalah_F")
    elif kode == "G":
        opsi = ["Ingin mengetahui prosedur", "Sudah mencoba tetapi gagal", "Tidak tahu harus mulai dari mana", "Lainnya"]
        st.radio("Ubah FKTP G", opsi, index=opsi.index(get_val('KPS_ubah_fktp')) if get_val('KPS_ubah_fktp') in opsi else None, key="unik_kps_KPS_ubah_fktp")
    elif kode == "H":
        opsi = ["Hilang", "Rusak", "Data tidak sesuai", "Tidak dapat digunakan", "Lainnya"]
        st.radio("Kartu H", opsi, index=opsi.index(get_val('KPS_kartu')) if get_val('KPS_kartu') in opsi else None, key="unik_kps_KPS_kartu")
    elif kode == "I":
        st.text_area("Beda I", value=get_val("KPS_beda"), key="unik_kps_KPS_beda")
        st.text_input("Tampil I", value=get_val("KPS_tampil"), key="unik_kps_KPS_tampil_I")
        st.text_input("Harus I", value=get_val("KPS_seharusnya"), key="unik_kps_KPS_seharusnya_I")
        f = st.file_uploader("Upload I", type=['jpg','png','jpeg','pdf','mp4','mov','avi'], accept_multiple_files=True, key="unik_kps_I_file")
        if f: st.session_state.r_data["file_bukti"] = f
    elif kode == "J":
        st.text_area("Tahu J", value=get_val("KPS_tahu"), key="unik_kps_KPS_tahu")
    elif kode == "K":
        st.text_area("Cerita K", value=get_val("KPS_cerita"), key="unik_kps_KPS_cerita")
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
    if st.button("Lanjut ➔", use_container_width=True): 
        for k in list(st.session_state.keys()):
            if k.startswith("unik_kps_"): st.session_state.r_data[k.replace("unik_kps_", "")] = st.session_state[k]
        st.session_state.p_step = "DATA_KONTAK"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "LAYANAN_PILIH":
    render_brand_header(); top_back_button("BPJS_MENU")
    st.markdown("<h1 class='anim-2 gradient-text' style='font-size: 26px; font-weight: 800; margin-bottom: 4px;'>Informasi Layanan JKN</h1><p class='anim-2' style='color: #64748b; font-size: 15px; margin-bottom: 28px;'>Pilih topik layanan yang ingin Anda tanyakan.</p>", unsafe_allow_html=True)
    opsi = ["Prosedur layanan", "Hak peserta", "Jadwal layanan", "Persyaratan", "Administrasi", "Pemeriksaan", "Dokter", "Laboratorium", "Farmasi/obat", "Biaya/ketentuan", "Informasi FKTP", "Lainnya"]
    st.markdown('<div class="anim-3">', unsafe_allow_html=True)
    for op in opsi:
        if st.button(f"**{op}**", use_container_width=True, key=f"btn_lyn_{op}"):
            st.session_state.r_data["active_masalah"] = op; st.session_state.p_step = "LAYANAN_FORM"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "LAYANAN_FORM":
    render_brand_header(); top_back_button("LAYANAN_PILIH")
    masalah = get_val("active_masalah", "Lainnya")
    st.markdown(f"<h1 class='anim-2 gradient-text' style='font-size: 22px; font-weight: 800; margin-bottom: 24px; line-height:1.4;'>{masalah}</h1>", unsafe_allow_html=True)
    st.markdown("<div class='anim-3 glass-box'>", unsafe_allow_html=True)
    if masalah == "Jadwal layanan":
        st.text_input("Layanan LYN", value=get_val("LYN_layanan"), key="unik_lyn_LYN_layanan")
        tgl = st.date_input("Tanggal LYN", key="unik_lyn_LYN_tgl"); st.session_state.r_data["LYN_tgl_str"] = tgl.strftime("%d %B %Y")
        st.text_input("Poli LYN", value=get_val("LYN_polidok"), key="unik_lyn_LYN_polidok")
    elif masalah in ["Persyaratan", "Biaya/ketentuan", "Pemeriksaan"]:
        st.text_input("Layanan LYN", value=get_val("LYN_layanan"), key="unik_lyn_LYN_layanan")
        st.text_area("Tanya LYN", value=get_val("LYN_tanya"), key="unik_lyn_LYN_tanya")
    elif masalah == "Administrasi":
        st.text_input("Bagian LYN", value=get_val("LYN_bagian"), key="unik_lyn_LYN_bagian")
        st.text_area("Masalah LYN", value=get_val("LYN_masalah"), key="unik_lyn_LYN_masalah")
    elif masalah == "Informasi FKTP":
        st.text_area("Cerita LYN", value=get_val("LYN_cerita"), key="unik_lyn_LYN_cerita")
    else: 
        if masalah in ["Dokter", "Laboratorium", "Farmasi/obat"]: st.text_input("Layanan LYN", value=get_val("LYN_layanan"), key="unik_lyn_LYN_layanan")
        st.text_area("Cerita LYN", value=get_val("LYN_cerita"), key="unik_lyn_LYN_cerita")
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
    if st.button("Lanjut ➔", use_container_width=True): 
        for k in list(st.session_state.keys()):
            if k.startswith("unik_lyn_"): st.session_state.r_data[k.replace("unik_lyn_", "")] = st.session_state[k]
        st.session_state.p_step = "DATA_KONTAK"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "LAINNYA_PILIH":
    render_brand_header(); top_back_button("BPJS_MENU")
    st.markdown("<h1 class='anim-2 gradient-text' style='font-size: 26px; font-weight: 800; margin-bottom: 4px;'>Tidak menemukan masalah yang sesuai?</h1><p class='anim-2' style='color: #64748b; font-size: 15px; margin-bottom: 28px;'>Ceritakan kebutuhan Anda, tim kami akan membantu.</p>", unsafe_allow_html=True)
    opsi = ["Pertanyaan umum BPJS/JKN", "Administrasi", "Pelayanan", "Sistem/aplikasi", "Dokumen", "Lainnya"]
    st.markdown('<div class="anim-3">', unsafe_allow_html=True)
    for op in opsi:
        if st.button(f"**{op}**", use_container_width=True, key=f"btn_lny_{op}"):
            st.session_state.r_data["active_masalah"] = op; st.session_state.p_step = "LAINNYA_FORM"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "LAINNYA_FORM":
    render_brand_header(); top_back_button("LAINNYA_PILIH")
    masalah = get_val("active_masalah")
    st.markdown(f"<h1 class='anim-2 gradient-text' style='font-size: 22px; font-weight: 800; margin-bottom: 24px; line-height:1.4;'>{masalah}</h1>", unsafe_allow_html=True)
    st.markdown("<div class='anim-3 glass-box'>", unsafe_allow_html=True)
    st.session_state.r_data["LNY_cerita"] = st.text_area("Cerita LNY", value=get_val("LNY_cerita"), key="unik_lny_LNY_cerita")
    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown('<div class="btn-primary">', unsafe_allow_html=True)
    if st.button("Lanjut ➔", use_container_width=True): st.session_state.p_step = "DATA_KONTAK"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------------------------------------------------
# LAYAR KONTAK & REVIEW TIKET/LOG
# ---------------------------------------------------------
elif st.session_state.p_step == "DATA_KONTAK":
    render_brand_header()
    module = get_val("active_module")
    if module == "RUJUKAN": top_back_button("RUJUKAN_FORM")
    elif module == "ANTREAN": top_back_button("ANTREAN_FORM")
    elif module == "MJKN": top_back_button("MJKN_FORM")
    elif module == "OBAT": top_back_button("OBAT_FORM")
    elif module == "KEPESERTAAN": top_back_button("KEPESERTAAN_FORM")
    elif module == "LAYANAN": top_back_button("LAYANAN_FORM")
    elif module == "LAINNYA": top_back_button("LAINNYA_FORM")
    elif module == "PENGADUAN": top_back_button("ADU_FORM")
    elif module == "SARAN": top_back_button("SARAN_FORM")
    elif module == "APRESIASI": top_back_button("APRES_FORM")
    else: top_back_button("LANDING")

    st.markdown("<h1 class='anim-1 gradient-text' style='font-size: 26px; font-weight: 800; margin-bottom: 24px;'>Bagaimana kami dapat menghubungi Anda?</h1>", unsafe_allow_html=True)
    st.markdown("<div class='anim-2 glass-box'>", unsafe_allow_html=True)
    
    nama = st.text_input("Nama Lengkap *", value=get_val("k_nama"), key="unik_kontak_nama")
    wa = st.text_input("WhatsApp *", value=get_val("k_wa"), key="unik_kontak_wa")
    
    if module in ["RUJUKAN", "KEPESERTAAN", "ANTREAN", "MJKN", "OBAT", "LAYANAN", "LAINNYA"]:
        st.markdown("<hr style='border:none; border-top:1px dashed #cbd5e1; margin:30px 0;'>", unsafe_allow_html=True)
        st.markdown("<p style='font-size:14.5px; font-weight:800; color:#475569; margin-bottom:16px;'>Data Identitas (Opsional)</p>", unsafe_allow_html=True)
        c_bpjs, c_nik = st.columns(2)
        with c_bpjs: st.text_input("Nomor BPJS", value=get_val("k_bpjs"), key="unik_kontak_bpjs")
        with c_nik: st.text_input("NIK KTP", value=get_val("k_nik"), key="unik_kontak_nik")
    
    st.checkbox("Saya bersedia dihubungi lewat WhatsApp untuk tindak lanjut.", value=get_val("k_izin", True), key="unik_kontak_izin")

    if "file_bukti" not in st.session_state.r_data:
        st.file_uploader("Upload Lampiran Tambahan (Jika ada)", type=['jpg','png','jpeg','pdf','mp4','mov','avi'], accept_multiple_files=True, key="unik_kontak_file")

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown('<div class="anim-3 btn-primary">', unsafe_allow_html=True)
    if st.button("Lanjut ke Review ➔", use_container_width=True):
        nama_val = st.session_state.get("unik_kontak_nama", "")
        wa_val = st.session_state.get("unik_kontak_wa", "")
        if module not in ["SARAN", "APRESIASI"] and (not nama_val or not wa_val):
            st.error("⚠️ Nama Lengkap dan Nomor WhatsApp wajib diisi.")
        else:
            st.session_state.r_data["k_nama"] = nama_val; st.session_state.r_data["k_wa"] = wa_val
            st.session_state.r_data["k_izin"] = st.session_state.get("unik_kontak_izin", False)
            if module not in ["PENGADUAN", "SARAN", "APRESIASI"]:
                st.session_state.r_data["k_bpjs"] = st.session_state.get("unik_kontak_bpjs", "")
                st.session_state.r_data["k_nik"] = st.session_state.get("unik_kontak_nik", "")
            if "unik_kontak_file" in st.session_state and st.session_state["unik_kontak_file"]:
                st.session_state.r_data["file_bukti"] = st.session_state["unik_kontak_file"]
            st.session_state.p_step = "REVIEW"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "REVIEW":
    render_brand_header(); top_back_button("DATA_KONTAK")
    st.markdown("<h1 class='anim-1 gradient-text' style='font-size: 26px; font-weight: 800; margin-bottom: 24px;'>Periksa Laporan Anda</h1>", unsafe_allow_html=True)
    st.markdown("<div class='anim-2 glass-box'>", unsafe_allow_html=True)
    
    module = get_val("active_module", "UMUM"); masalah = get_val("active_masalah") or get_val("rujukan_masalah")
    
    if module in ["PENGADUAN", "SARAN", "APRESIASI"]: kategori_string = module.title(); sub_kat = module.title()
    else: kategori_string = "Bantuan BPJS/JKN"; sub_kat = module.title() if module != "MJKN" else "Mobile JKN"; sub_kat = "Lainnya" if module == "LAINNYA" else sub_kat
    
    st.markdown(f"<p style='color:#64748b; font-size:14px; margin-bottom:4px;'>Kategori Laporan</p><p style='color:#0f172a; font-size:16px; font-weight:700; margin-bottom:16px;'>{kategori_string} — {sub_kat}</p>", unsafe_allow_html=True)
    if masalah: st.markdown(f"<p style='color:#64748b; font-size:14px; margin-bottom:4px;'>Topik / Kendala</p><p style='color:#0f172a; font-size:16px; font-weight:700; margin-bottom:16px;'>{masalah}</p>", unsafe_allow_html=True)
    
    detail_string = ""
    exclude_keys = ['active_module', 'active_masalah', 'rujukan_masalah', 'k_nama', 'k_wa', 'k_bpjs', 'k_nik', 'k_izin', 'file_bukti']
    for key, val in st.session_state.r_data.items():
        if key not in exclude_keys and val:
            label_bersih = key[2:] if len(key) > 2 and key[1] == '_' else key 
            label_bersih = key[4:] if len(key) > 4 and key[3] == '_' else label_bersih 
            label_bersih = label_bersih.replace("_", " ").title()
            st.markdown(f"<p style='color:#64748b; font-size:14px; margin-bottom:4px;'>{label_bersih}</p><p style='color:#0f172a; font-size:15px; font-weight:600; margin-bottom:16px;'>{val}</p>", unsafe_allow_html=True)
            detail_string += f"• {label_bersih}: {val}\n"
            
    st.markdown("<hr style='border:none; border-top:1px dashed #cbd5e1; margin:24px 0;'>", unsafe_allow_html=True)
    st.markdown(f"- **Nama:** {get_val('k_nama', 'Anonim')}\n- **WhatsApp:** {get_val('k_wa', 'Tidak diisi')}")
    
    file_count = len(get_val('file_bukti')) if isinstance(get_val('file_bukti'), list) else (1 if get_val('file_bukti') else 0)
    if file_count > 0: st.markdown(f"- **Lampiran:** {file_count} File")

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="anim-3 btn-primary">', unsafe_allow_html=True)
    if st.button("✓ Kirim Permintaan", use_container_width=True):
        with st.spinner("⏳ Mengirim data & melampirkan file ke server... Mohon tunggu sebentar."):
            nama_pasien = get_val('k_nama', 'Anonim'); wa_pasien = get_val('k_wa', '-'); bpjs_pasien = get_val('k_bpjs', '-'); nik_pasien = get_val('k_nik', '-')
            consent_wa = 1 if get_val('k_izin') else 0
            contact_id = get_or_create_contact(nama_pasien, wa_pasien, bpjs_pasien, nik_pasien, consent_wa)

            prefix = "RUJ" if module == "RUJUKAN" else "ADU" if module == "PENGADUAN" else "SAR" if module == "SARAN" else "APR" if module == "APRESIASI" else "PIP"
            ticket_id = generate_ticket_id(prefix); short_link = generate_short_link()
            st.session_state.r_data["ticket"] = ticket_id; st.session_state.r_data["short_link"] = short_link
            
            waktu = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            identitas = f"BPJS: {bpjs_pasien} / NIK: {nik_pasien}"
            jenis_lengkap = kategori_string if module in ["PENGADUAN", "SARAN", "APRESIASI"] else f"BPJS - {sub_kat}"
            keluhan_final = f"[TICKET: {ticket_id}]\nMasalah: {masalah}\n\nDetail:\n{detail_string}"
            
            pesan_notif = f"🚨 <b>TIKET BARU: {ticket_id}</b> 🚨\n\nJalur: {jenis_lengkap}\nNama: {nama_pasien}\nWA: {wa_pasien}\nLampiran: {file_count} file\n\n{keluhan_final}"
            file_ids_str = kirim_notif_telegram_dengan_file(pesan_notif, get_val("file_bukti"))

            conn = get_conn()
            status_awal = 'Selesai' if module in ["SARAN", "APRESIASI"] else 'Baru'
            conn.execute('''INSERT INTO pipp (waktu_input, contact_id, no_bpjs, nama_pasien, alamat, no_telp, jenis_laporan, keluhan, solusi, link_bukti, is_real, status, priority, short_link, consent_wa) VALUES (?,?,?,?,?,?,?,?,?,?,1, ?, 'Normal', ?, ?)''', (waktu, contact_id, identitas, nama_pasien, "-", wa_pasien, jenis_lengkap, keluhan_final, '[MENUNGGU TINDAKAN]', file_ids_str, status_awal, short_link, consent_wa))
            inserted_ticket_id = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
            
            is_ticketed = 0 if module in ["SARAN", "APRESIASI"] else 1
            conn.execute('''INSERT INTO interactions (contact_id, module, details, is_ticketed, created_at) VALUES (?,?,?,?,?)''', (contact_id, module, masalah, is_ticketed, waktu))
            conn.execute('''INSERT INTO ticket_logs (ticket_id, actor, action, note, is_internal, created_at) VALUES (?,?,?,?,?,?)''', (inserted_ticket_id, 'System', 'Dibuat', 'Pasien menyelesaikan form layanan mandiri.', 1, waktu))
            conn.commit(); conn.close()
            
        st.session_state.p_step = "SUCCESS"; st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

elif st.session_state.p_step == "SUCCESS":
    st.balloons()
    module = get_val("active_module", "UMUM")
    
    if module in ["SARAN", "APRESIASI"]:
        st.markdown(f"""
        <div class="page-trans" style="text-align:center; padding: 50px 20px; background:rgba(255,255,255,0.9); border-radius:24px; border: 1px solid rgba(255,255,255,1); margin-top:20px; box-shadow:0 10px 40px rgba(0,0,0,0.06); backdrop-filter:blur(15px);">
            <div style="font-size: 80px; color: #10b981; margin-bottom: 15px; line-height:1; text-shadow: 0 4px 20px rgba(16,185,129,0.3);">❤️</div>
            <h2 class="gradient-text" style="font-weight:900; margin-bottom: 10px; font-size:28px;">Terima Kasih Banyak!</h2>
            <p style="color:#64748b; font-size: 16px; margin: 15px auto; max-width:85%; line-height:1.6;">Suara Anda telah kami terima dan akan menjadi semangat kami untuk terus memperbaiki kualitas layanan di Klinik Pratama Papan Mas.</p>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="anim-2 btn-primary">', unsafe_allow_html=True)
        if st.button("🏠 Kembali ke Beranda", use_container_width=True): st.session_state.p_step = "LANDING"; st.session_state.r_data = {}; st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.markdown(f"""
        <div class="page-trans" style="text-align:center; padding: 50px 20px; background:rgba(255,255,255,0.9); border-radius:24px; border: 1px solid rgba(255,255,255,1); margin-top:20px; box-shadow:0 10px 40px rgba(0,0,0,0.06); backdrop-filter:blur(15px);">
            <div style="font-size: 80px; color: #10b981; margin-bottom: 15px; line-height:1; text-shadow: 0 4px 20px rgba(16,185,129,0.3);">✓</div>
            <h2 class="gradient-text" style="font-weight:900; margin-bottom: 10px; font-size:28px;">Laporan Berhasil Terkirim</h2>
            <div style="background: linear-gradient(135deg, #ecfdf5, #ffffff); border: 2px dashed #a7f3d0; padding: 25px 45px; border-radius: 18px; margin: 25px 0; display:inline-block; box-shadow: 0 6px 15px rgba(16,185,129,0.08);">
                <p style="margin:0; font-size:15px; color:#10b981; font-weight:700; margin-bottom:8px;">Nomor Tiket Anda:</p>
                <h3 style="margin:0; color:#047857; font-weight:900; letter-spacing:2px; font-size:28px; font-family: 'Outfit', sans-serif;">{get_val("ticket")}</h3>
            </div>
            <p style="color:#64748b; font-size: 16px;">Status Saat Ini: <b style="color:#0f172a; background:#f8fafc; padding: 6px 16px; border-radius: 8px; border: 1px solid #e2e8f0; box-shadow:0 2px 4px rgba(0,0,0,0.02);">BARU</b></p>
            <div style="background:#f1f5f9; padding: 15px; border-radius:12px; margin-top: 20px;">
                <p style="margin:0; font-size:14px; font-weight:700; color:#475569; margin-bottom:8px;">Pantau status laporan Anda melalui link berikut:</p>
                <p style="margin:0; font-size:15px; font-weight:800; color:#0f172a; word-break: break-all;">{APP_URL}/?status={get_val("short_link")}</p>
            </div>
            <p style="color:#94a3b8; font-size: 13px; max-width:90%; margin: 20px auto 10px auto; line-height:1.6;">Tim kami akan segera memproses laporan Anda. Notifikasi akan dikirim ke WhatsApp Anda (jika layanan tersedia).</p>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown('<div class="anim-2 btn-outline">', unsafe_allow_html=True)
            st.button("📋 Salin Nomor Tiket", use_container_width=True) 
            st.markdown('</div>', unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="anim-2 btn-primary">', unsafe_allow_html=True)
            if st.button("🏠 Kembali ke Beranda", use_container_width=True): 
                st.session_state.p_step = "LANDING"; st.session_state.r_data = {}; st.query_params.clear(); st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
        
st.markdown('</div>', unsafe_allow_html=True)