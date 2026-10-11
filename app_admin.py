import streamlit as st
from database import init_db, get_conn, PIN_ADMIN_RAHASIA
from modul_admin.tab_dasbor import render_tab_dasbor
from modul_admin.tab_importer import render_tab_importer
from modul_admin.tab_input import render_tab_input
from modul_admin.tab_form_builder import render_tab_form_builder # <--- TAMBAH INI

st.set_page_config(page_title="Admin - Database Klinik", layout="wide", initial_sidebar_state="collapsed")
init_db()
conn = get_conn()

st.markdown("""
<style>
    .stApp { background-color: #f8fafc !important; font-family: 'Inter', sans-serif;}
    .filter-chip { background-color: #e2e8f0; color: #0f172a; padding: 4px 12px; border-radius: 16px; font-size: 13px; font-weight: 600; margin-right: 8px; display: inline-block; margin-bottom: 8px;}
    .card-detail { background-color: #ffffff; padding: 20px; border-radius: 12px; border: 1px solid #e2e8f0; margin-top:10px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);}
    .section-header { font-size: 18px; font-weight: 700; color: #1e293b; margin-top: 20px; margin-bottom: 10px; border-bottom: 2px solid #e2e8f0; padding-bottom: 5px;}
</style>
""", unsafe_allow_html=True)

if "admin_logged_in" not in st.session_state: st.session_state.admin_logged_in = False

if not st.session_state.admin_logged_in:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.info("🔐 Masukkan PIN Admin untuk mengakses Database.")
        pin_input = st.text_input("PIN Keamanan:", type="password")
        if st.button("Masuk"):
            if pin_input == PIN_ADMIN_RAHASIA:
                st.session_state.admin_logged_in = True; st.rerun()
            else:
                st.error("PIN Salah!")
else:
    c1, c2 = st.columns([4, 1])
    with c1: st.title("🗄 Database Utama Klinik")
    with c2: 
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚪 Logout", use_container_width=True): 
            st.session_state.admin_logged_in = False; st.rerun()

    # EXTEND: MENAMBAHKAN TAB FORM BUILDER
    tab_db, tab_import, tab_input, tab_builder = st.tabs(["📊 Update Penanganan (Dasbor)", "📥 Import Data", "✍ Input Pasien", "📋 Form Builder"])

    with tab_db: render_tab_dasbor(conn)
    with tab_import: render_tab_importer(conn)
    with tab_input: render_tab_input(conn)
    with tab_builder: render_tab_form_builder(conn)