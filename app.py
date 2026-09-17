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
    page_title="Form Validasi Mutasi Rekening",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Light Theme CSS untuk tampilan persis seperti permintaan user
st.markdown("""
<style>
    /* Paksa Background Light Mode */
    html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
        background-color: #ffffff !important;
        color: #1e293b !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    }

    [data-testid="stSidebar"] {
        background-color: #f8fafc !important;
        border-right: 1px solid #e2e8f0;
    }

    /* Tabel Summary Bersih Sesuai Contoh */
    .summary-table-container {
        width: 100%;
        overflow-x: auto;
        margin-top: 10px;
        margin-bottom: 30px;
        border: 1px solid #e5e7eb;
        border-radius: 8px;
        background-color: #ffffff;
    }

    .clean-summary-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.90rem;
        background-color: #ffffff;
    }

    .clean-summary-table th {
        background-color: #ffffff;
        color: #64748b;
        font-weight: 600;
        text-align: left;
        padding: 12px 14px;
        border-bottom: 1px solid #e5e7eb;
        border-right: 1px solid #f1f5f9;
        white-space: nowrap;
    }

    .clean-summary-table td {
        padding: 13px 14px;
        border-bottom: 1px solid #f1f5f9;
        border-right: 1px solid #f8fafc;
        color: #1e293b;
        white-space: nowrap;
    }

    .clean-summary-table td.bold-text {
        font-weight: 700;
        color: #0f172a;
    }

    .clean-summary-table .text-right {
        text-align: right;
    }

    .clean-summary-table .text-center {
        text-align: center;
    }

    .clean-summary-table tr:last-child td {
        border-bottom: none;
    }

    .clean-summary-table tr:hover {
        background-color: #f8fafc;
    }

    /* Section Title */
    .section-download-title {
        font-size: 1.35rem;
        font-weight: 700;
        color: #1e293b;
        display: flex;
        align-items: center;
        gap: 10px;
        margin-top: 10px;
        margin-bottom: 18px;
    }

    /* Kartu Box Unduh */
    .download-card {
        background-color: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 20px 22px;
        min-height: 105px;
        margin-bottom: 14px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }

    .download-card-title {
        font-size: 1rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        gap: 8px;
    }

    .download-card-desc {
        font-size: 0.85rem;
        color: #64748b;
        line-height: 1.4;
        margin: 0;
    }

    /* Tombol Unduh Clean Outline */
    div.stDownloadButton > button {
        background-color: #ffffff !important;
        color: #1e293b !important;
        border: 1px solid #d1d5db !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
        font-size: 0.90rem !important;
        padding: 10px 16px !important;
        width: 100% !important;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04) !important;
        transition: all 0.15s ease !important;
    }

    div.stDownloadButton > button:hover {
        background-color: #f8fafc !important;
        border-color: #94a3b8 !important;
        color: #0f172a !important;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# Format Currency Angka Sesuai Contoh (134,181,200.00)
def format_num_curr(val):
    if val is None or pd.isna(val):
        return "0.00"
    return f"{val:,.2f}"

def format_count(val):
    if val is None or pd.isna(val):
        return "0"
    return f"{int(val)}"

# ==============================================================================
# 2. SIDEBAR - FORM INPUT DATA NASABAH
# ==============================================================================
with st.sidebar:
    st.markdown("### 📋 Input Data Nasabah")
    st.caption("Identitas nasabah untuk disematkan pada dokumen cetak PDF & Excel.")

    cabang = st.text_input("Cabang", value="", placeholder="Contoh: Jakarta Thamrin")
    nama_cust = st.text_input("Nama Cust", value="", placeholder="Nama debitur / nasabah")
    no_rekening = st.text_input("Nomor Rekening", value="", placeholder="Contoh: 123-00-1234567-8")

    bank_options = ["BANK MANDIRI", "BANK BCA", "BANK BRI", "BANK BNI", "BANK PERMATA", "BANK NOBU", "LAINNYA"]
    nama_bank_selected = st.selectbox("Nama Bank", options=bank_options, index=0)
    if nama_bank_selected == "LAINNYA":
        nama_bank = st.text_input("Sebutkan Nama Bank", value="").strip().upper()
    else:
        nama_bank = nama_bank_selected

    nama_pemegang_rek = st.text_input("Nama Pemegang Rekening", value="", placeholder="Sesuai buku tabungan")

    st.markdown("---")
    st.markdown("### ✍️ Petugas & Catatan")
    nama_so = st.text_input("Nama SO (Sales Officer)", value="", placeholder="Nama Sales Officer / AO")
    note_oh = st.text_area("Note OH (Operation Head)", value="-", height=75, placeholder="Catatan analisa...")

    header_input = {
        "cabang": cabang,
        "nama_cust": nama_cust,
        "no_rekening": no_rekening,
        "nama_bank": nama_bank,
        "nama_pemegang_rek": nama_pemegang_rek
    }

# ==============================================================================
# 3. KONTEN UTAMA: UPLOADER & PROSES
# ==============================================================================
st.title("🏦 Form Validasi Mutasi Rekening")
st.caption("Parser rekening koran otomatis multi-bank (Mandiri, BCA, BRI, BNI, Permata, Nobu)")

uploaded_files = st.file_uploader(
    "Unggah Rekening Koran PDF (Pilih 1 s/d 3 berkas):",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    col_btn, col_info = st.columns([2, 5])
    with col_btn:
        process_btn = st.button("⚡ Proses Rekening Koran", type="primary", use_container_width=True)
    with col_info:
        st.write(f"📁 **{len(uploaded_files)} berkas** siap diproses.")

    if process_btn:
        with st.spinner("Mengekstrak data mutasi rekening..."):
            resume_list = []
            for up_file in uploaded_files:
                try:
                    data_ext = parse_rekening_universal(up_file, filename=up_file.name)
                    resume_list.append(data_ext)
                except Exception as e:
                    st.error(f"Gagal memproses {up_file.name}: {str(e)}")

            if resume_list:
                sorted_resume = sort_resume_chronological(resume_list)
                st.session_state["resume_list"] = sorted_resume
                st.session_state["processed_header"] = header_input
                st.session_state["processed_nama_so"] = nama_so
                st.session_state["processed_note_oh"] = note_oh

# ==============================================================================
# 4. TAMPILAN HASIL PERSIS SEPERTI GAMBAR USER
# ==============================================================================
if "resume_list" in st.session_state and st.session_state["resume_list"]:
    resume_list = st.session_state["resume_list"]
    cur_header = st.session_state.get("processed_header", header_input)
    p_note_oh = st.session_state.get("processed_note_oh", note_oh)
    p_nama_so = st.session_state.get("processed_nama_so", nama_so)

    # 1. Tabel Utama Bersih Sesuai Screenshot
    table_rows_html = ""
    for it in resume_list:
        b_name = it.get("bulan", "-")
        b_bank = it.get("bank", "MANDIRI")
        f_db = format_count(it.get("freq_db", 0))
        f_cr = format_count(it.get("freq_cr", 0))
        m_db = format_num_curr(it.get("mutasi_db", 0.0))
        m_cr = format_num_curr(it.get("mutasi_cr", 0.0))
        s_max = format_num_curr(it.get("saldo_max", 0.0))
        s_avg = format_num_curr(it.get("saldo_avg", 0.0))
        s_min = format_num_curr(it.get("saldo_min", 0.0))

        table_rows_html += f"""
        <tr>
            <td class="bold-text">{b_name}</td>
            <td class="bold-text">{b_bank}</td>
            <td class="text-right">{f_db}</td>
            <td class="text-right">{f_cr}</td>
            <td class="text-right">{m_db}</td>
            <td class="text-right">{m_cr}</td>
            <td class="text-right">{s_max}</td>
            <td class="text-right">{s_avg}</td>
            <td class="text-right">{s_min}</td>
        </tr>
        """

    full_table_html = f"""
    <div class="summary-table-container">
        <table class="clean-summary-table">
            <thead>
                <tr>
                    <th style="width: 8%;">Bulan</th>
                    <th style="width: 9%;">Bank</th>
                    <th class="text-right" style="width: 8%;">Freq Debet</th>
                    <th class="text-right" style="width: 8%;">Freq Kredit</th>
                    <th class="text-right" style="width: 14%;">Total Debet (Rp)</th>
                    <th class="text-right" style="width: 14%;">Total Kredit (Rp)</th>
                    <th class="text-right" style="width: 13%;">Saldo Tertinggi (Rp)</th>
                    <th class="text-right" style="width: 13%;">Saldo Rata-Rata (Rp)</th>
                    <th class="text-right" style="width: 13%;">Saldo Terendah (Rp)</th>
                </tr>
            </thead>
            <tbody>
                {table_rows_html}
            </tbody>
        </table>
    </div>
    """
    st.markdown(full_table_html, unsafe_allow_html=True)

    # 2. Section Unduh Dokumen Form Validasi Sesuai Screenshot
    st.markdown("""
    <div class="section-download-title">
        <span>📩</span> Unduh Dokumen Form Validasi
    </div>
    """, unsafe_allow_html=True)

    # Generate File Buffers
    pdf_bytes = generate_form_pdf(
        None,
        cur_header,
        resume_list,
        note_oh=p_note_oh,
        nama_so=p_nama_so
    )

    excel_bytes = generate_form_excel(
        None,
        cur_header,
        resume_list,
        note_oh=p_note_oh,
        nama_so=p_nama_so
    )

    col_pdf, col_excel = st.columns(2)

    with col_pdf:
        st.markdown("""
        <div class="download-card">
            <div class="download-card-title">
                <span>📄</span> Dokumen PDF Resmi
            </div>
            <p class="download-card-desc">
                Format standar A4 siap cetak dengan tanda tangan SO & Operation Head.
            </p>
        </div>
        """, unsafe_allow_html=True)

        st.download_button(
            label="Unduh PDF Form Validasi",
            data=pdf_bytes,
            file_name="FORM_VALIDASI_MUTASI_REKENING.pdf",
            mime="application/pdf",
            use_container_width=True
        )

    with col_excel:
        st.markdown("""
        <div class="download-card">
            <div class="download-card-title">
                <span>📊</span> Dokumen Excel Spreadsheet
            </div>
            <p class="download-card-desc">
                Format lembar kerja dinamis lengkap dengan rincian transaksi per bulan.
            </p>
        </div>
        """, unsafe_allow_html=True)

        st.download_button(
            label="Unduh Excel Form Validasi (.xlsx)",
            data=excel_bytes,
            file_name="FORM_VALIDASI_MUTASI_REKENING.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # Expander Opsional untuk melihat transaksi jika dibutuhkan
    with st.expander("🔍 Lihat Rincian Transaksi Lengkap (Opsional)", expanded=False):
        tab_titles = [f"📅 {item.get('bulan', f'Bulan {i+1}')}" for i, item in enumerate(resume_list)]
        month_tabs = st.tabs(tab_titles)
        for i, tab in enumerate(month_tabs):
            with tab:
                cur_month_data = resume_list[i]
                records = cur_month_data.get("tx_records", [])
                if records:
                    df_tx = pd.DataFrame(records)
                    df_tx.columns = ["Tanggal", "Debet", "Kredit", "Saldo"]
                    df_tx["Debet"] = df_tx["Debet"].apply(lambda x: format_num_curr(x) if x > 0 else "-")
                    df_tx["Kredit"] = df_tx["Kredit"].apply(lambda x: format_num_curr(x) if x > 0 else "-")
                    df_tx["Saldo"] = df_tx["Saldo"].apply(lambda x: format_num_curr(x) if x is not None else "-")
                    st.dataframe(df_tx, use_container_width=True, hide_index=True)
                else:
                    st.info("Tidak ada data baris transaksi.")
