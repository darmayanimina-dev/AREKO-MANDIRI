import io
import streamlit as st
import pandas as pd
import numpy as np
from parser_engine import (
    parse_rekening_universal,
    generate_form_pdf,
    generate_form_excel,
    sort_resume_chronological,
    MONTH_NAMES_ID
)

# ==============================================================================
# 1. KONFIGURASI HALAMAN & TEMA LIGHT MODE
# ==============================================================================
st.set_page_config(
    page_title="AREKO - Let’s Make Recap Less Recap-y",
    page_icon="💳",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Inisialisasi Session State
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0
if "resume_list" not in st.session_state:
    st.session_state["resume_list"] = None

def reset_all():
    st.session_state["resume_list"] = None
    st.session_state["uploader_key"] += 1
    st.session_state["inp_cabang"] = ""
    st.session_state["inp_cust"] = ""
    st.session_state["inp_rek"] = ""
    st.session_state["inp_bank"] = "Mandiri"
    st.session_state["inp_pemegang"] = ""
    st.session_state["inp_so"] = ""
    st.session_state["inp_oh"] = ""
    st.session_state["inp_note"] = "-"
    st.rerun()

# Custom Light Theme CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

    html, body, [class*="css"], [data-testid="stAppViewContainer"], [data-testid="stApp"] {
        font-family: 'Plus Jakarta Sans', sans-serif !important;
        background-color: #ffffff !important;
        color: #1e293b !important;
    }

    [data-testid="stSidebar"] {
        background-color: #f8fafc !important;
        border-right: 1px solid #e2e8f0;
    }

    .sidebar-section-title {
        font-size: 11px;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #64748b;
        margin: 16px 0 8px 0;
        padding-bottom: 4px;
        border-bottom: 1px solid #e2e8f0;
    }

    /* Empty State Box */
    .empty-state-box {
        background-color: #ffffff;
        border: 1px dashed #cbd5e1;
        border-radius: 12px;
        padding: 44px 20px;
        text-align: center;
        margin-top: 20px;
    }
    .empty-state-icon {
        font-size: 2.6rem;
        margin-bottom: 10px;
    }
    .empty-state-title {
        font-size: 1.1rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 6px;
    }
    .empty-state-desc {
        font-size: 0.88rem;
        color: #64748b;
        max-width: 520px;
        margin: 0 auto;
        line-height: 1.5;
    }

    /* Download Cards */
    .dl-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 18px 20px;
        margin-bottom: 12px;
        min-height: 90px;
    }
    .dl-card-title {
        font-size: 15px;
        font-weight: 600;
        color: #0f172a;
        margin-bottom: 4px;
    }
    .dl-card-desc {
        font-size: 12.5px;
        color: #64748b;
        line-height: 1.4;
    }

    /* Clean Buttons */
    div.stDownloadButton > button:first-child {
        border-radius: 8px !important;
        font-weight: 500 !important;
        border: 1px solid #cbd5e1 !important;
        background-color: #ffffff !important;
        color: #1e293b !important;
        width: 100% !important;
        padding: 10px 16px !important;
        transition: all 0.15s ease-in-out !important;
    }
    div.stDownloadButton > button:first-child:hover {
        background-color: #f8fafc !important;
        border-color: #94a3b8 !important;
        color: #0f172a !important;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. SIDEBAR - FORM INPUT DATA NASABAH
# ==============================================================================
with st.sidebar:
    st.markdown("### ⚙️ Parameter Dokumen")

    st.markdown('<div class="sidebar-section-title">Informasi Nasabah</div>', unsafe_allow_html=True)
    cabang = st.text_input("Cabang", value=st.session_state.get("inp_cabang", ""), key="inp_cabang", placeholder="Contoh: Jakarta Pusat")
    nama_cust = st.text_input("Nama Cust", value=st.session_state.get("inp_cust", ""), key="inp_cust", placeholder="Nama lengkap customer")
    no_rek = st.text_input("Nomor Rekening", value=st.session_state.get("inp_rek", ""), key="inp_rek", placeholder="Contoh: 1230009876543")

    bank_list = ["Mandiri", "BCA", "BNI", "BRI", "BSI", "BJB", "Permata", "Nobu"]
    current_bank = st.session_state.get("inp_bank", "Mandiri")
    bank_idx = bank_list.index(current_bank) if current_bank in bank_list else 0
    nama_bank = st.selectbox("Nama Bank", bank_list, index=bank_idx, key="inp_bank")

    pemegang_rek = st.text_input("Nama Pemegang Rekening", value=st.session_state.get("inp_pemegang", ""), key="inp_pemegang", placeholder="Sesuai buku tabungan")

    st.markdown('<div class="sidebar-section-title">Pengesahan & Catatan</div>', unsafe_allow_html=True)
    nama_so = st.text_input("Nama SO (Sales Officer)", value=st.session_state.get("inp_so", ""), key="inp_so", placeholder="Nama Sales Officer")
    nama_oh = st.text_input("Operation Head", value=st.session_state.get("inp_oh", ""), key="inp_oh", placeholder="Nama Operation Head")
    note_oh = st.text_area("Note OH", value=st.session_state.get("inp_note", "-"), key="inp_note", placeholder="Catatan dari Operation Head...")

    st.markdown("---")
    if st.button("🔄 Reset Form & Data", use_container_width=True):
        reset_all()

    header_input = {
        "cabang": cabang,
        "nama_cust": nama_cust,
        "no_rekening": no_rek,
        "nama_bank": nama_bank,
        "nama_pemegang_rek": pemegang_rek,
        "nama_so": nama_so,
        "nama_oh": nama_oh,
    }

# ==============================================================================
# 3. KONTEN UTAMA: HEADER & FILE UPLOADER
# ==============================================================================
st.markdown("""
<div style="margin-bottom: 24px;">
    <h1 style="font-size: 2.2rem; font-weight: 800; color: #0f172a; margin: 0; letter-spacing: -0.02em;">AREKO</h1>
    <p style="font-size: 1.05rem; font-weight: 500; color: #64748b; margin-top: 4px; margin-bottom: 0;">Let’s Make Recap Less Recap-y</p>
</div>
""", unsafe_allow_html=True)

st.markdown("#### 📂 Berkas Rekening Koran")
uploaded_files = st.file_uploader(
    "Pilih 1 hingga 3 Berkas Rekening Koran (Format PDF)",
    type=["pdf"],
    accept_multiple_files=True,
    key=f"uploader_{st.session_state['uploader_key']}",
    help="Unggah dokumen PDF mutasi rekening untuk dianalisis dan direkapitulasi secara otomatis."
)

if uploaded_files:
    col_btn, col_rst = st.columns([2, 1])
    with col_btn:
        proses_clicked = st.button("🚀 Ekstraksi & Rekap Data", type="primary", use_container_width=True)
    with col_rst:
        if st.button("🔄 Reset", use_container_width=True):
            reset_all()

    if proses_clicked:
        resume_list = []
        progress_bar = st.progress(0, text="Memulai ekstraksi...")

        for i, file_obj in enumerate(uploaded_files):
            progress_bar.progress(
                (i + 1) / len(uploaded_files),
                text=f"Membaca berkas: {file_obj.name}...",
            )
            extracted = parse_rekening_universal(file_obj)
            resume_list.append(extracted)

        st.session_state["resume_list"] = resume_list
        progress_bar.empty()
        st.toast("Ekstraksi data mutasi berhasil selesai!", icon="✅")
        st.rerun()

# ==============================================================================
# 4. TABEL & UNDUH DOKUMEN (HANYA MUNCUL JIKA SUDAH ADA DATA / EKSTRAKSI)
# ==============================================================================
if st.session_state.get("resume_list"):
    res_list = st.session_state["resume_list"]
    sorted_data = sort_resume_chronological(res_list)

    st.markdown("<br>", unsafe_allow_html=True)

    preview_rows = []
    for r in sorted_data:
        preview_rows.append({
            "Bulan": r.get("bulan", "-"),
            "Bank": r.get("bank", nama_bank),
            "Freq Debet": r.get("freq_db", 0),
            "Freq Kredit": r.get("freq_cr", 0),
            "Total Debet (Rp)": f"{r.get('mutasi_db', 0):,.2f}",
            "Total Kredit (Rp)": f"{r.get('mutasi_cr', 0):,.2f}",
            "Saldo Tertinggi (Rp)": f"{r.get('saldo_max', 0):,.2f}",
            "Saldo Rata-Rata (Rp)": f"{r.get('saldo_avg', 0):,.2f}",
            "Saldo Terendah (Rp)": f"{r.get('saldo_min', 0):,.2f}",
        })

    # 1. Tabel Native Streamlit
    st.dataframe(preview_rows, use_container_width=True, hide_index=True)

    # 2. Section Unduh Hasil
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("#### 📥 Unduh Dokumen Form Validasi")
    col_dl1, col_dl2 = st.columns(2)

    # Generate Buffer File
    pdf_bytes = generate_form_pdf(
        None,
        header_input,
        res_list,
        note_oh=note_oh,
        nama_so=nama_so,
        nama_oh=nama_oh,
    )

    xlsx_bytes = generate_form_excel(
        None,
        header_input,
        res_list,
        note_oh=note_oh,
        nama_so=nama_so,
        nama_oh=nama_oh,
    )

    with col_dl1:
        st.markdown(
            """
            <div class="dl-card">
                <div class="dl-card-title">📄 Dokumen PDF Resmi</div>
                <div class="dl-card-desc">Format standar A4 siap cetak dengan tanda tangan SO & Operation Head.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.download_button(
            label="Unduh PDF Form Validasi",
            data=pdf_bytes,
            file_name="FORM_VALIDASI_MUTASI_REKENING.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    with col_dl2:
        st.markdown(
            """
            <div class="dl-card">
                <div class="dl-card-title">📊 Dokumen Excel Spreadsheet</div>
                <div class="dl-card-desc">Format lembar kerja dinamis lengkap dengan rincian transaksi per bulan.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.download_button(
            label="Unduh Excel Form Validasi (.xlsx)",
            data=xlsx_bytes,
            file_name="FORM_VALIDASI_MUTASI_REKENING.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

else:
    # KONDISI AWAL (EMPTY STATE)
    st.markdown("""
    <div class="empty-state-box">
        <div class="empty-state-icon">📄</div>
        <div class="empty-state-title">Belum Ada Rekening Koran yang Diproses</div>
        <p class="empty-state-desc">
            Silakan unggah berkas PDF rekening koran di atas, lengkapi parameter dokumen di sidebar, lalu klik <b>Ekstraksi & Rekap Data</b> untuk menampilkan tabel ringkasan mutasi dan mengunduh formulir validasi.
        </p>
    </div>
    """, unsafe_allow_html=True)
