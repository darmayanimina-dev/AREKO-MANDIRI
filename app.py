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

# Custom Light Theme CSS untuk tampilan bersih, elegan, dan profesional
st.markdown("""
<style>
    /* Paksa Background Light Mode */
    html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"] {
        background-color: #f8fafc !important;
        color: #0f172a !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    }

    [data-testid="stSidebar"] {
        background-color: #ffffff !important;
        border-right: 1px solid #e2e8f0;
    }

    /* Card Box Tampilan Modern */
    .metric-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px 20px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
        transition: transform 0.15s ease, box-shadow 0.15s ease;
    }
    .metric-card:hover {
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
    }
    .metric-label {
        font-size: 0.8rem;
        font-weight: 600;
        text-transform: uppercase;
        color: #64748b;
        letter-spacing: 0.5px;
    }
    .metric-value {
        font-size: 1.45rem;
        font-weight: 700;
        color: #0f172a;
        margin-top: 4px;
    }
    .metric-sub {
        font-size: 0.78rem;
        color: #94a3b8;
        margin-top: 2px;
    }

    /* Badge Bank */
    .bank-pill {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 9999px;
        font-size: 0.75rem;
        font-weight: 700;
        letter-spacing: 0.3px;
        margin-right: 6px;
        margin-bottom: 6px;
    }
    .bank-mandiri { background-color: #003d79; color: #ffb703; }
    .bank-bca     { background-color: #0060af; color: #ffffff; }
    .bank-bri     { background-color: #00529c; color: #ff7700; }
    .bank-bni     { background-color: #005e6a; color: #f15a24; }
    .bank-permata { background-color: #2b7a4b; color: #ffffff; }
    .bank-nobu    { background-color: #5c2d91; color: #ffffff; }
    .bank-umum    { background-color: #475569; color: #ffffff; }

    /* Header Banner */
    .header-banner {
        background: linear-gradient(135deg, #ffffff 0%, #f1f5f9 100%);
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 22px 28px;
        margin-bottom: 20px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.03);
    }
    .header-title {
        font-size: 1.6rem;
        font-weight: 800;
        color: #0f172a;
        margin: 0;
        line-height: 1.2;
    }
    .header-desc {
        font-size: 0.92rem;
        color: #64748b;
        margin-top: 6px;
        margin-bottom: 12px;
    }

    /* Tabel Resmi Mutasi */
    .official-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.84rem;
        background-color: #ffffff;
        border-radius: 8px;
        overflow: hidden;
        border: 1px solid #cbd5e1;
    }
    .official-table th {
        background-color: #D9E1F2;
        color: #1e293b;
        font-weight: 700;
        text-align: center;
        padding: 9px 8px;
        border: 1px solid #94a3b8;
        font-size: 0.8rem;
    }
    .official-table td {
        padding: 7px 10px;
        border: 1px solid #cbd5e1;
        color: #1e293b;
    }
    .official-table .text-right {
        text-align: right;
    }
    .official-table .text-center {
        text-align: center;
    }
    .official-table tr.avg-row td {
        background-color: #E2EFDA !important;
        font-weight: 700;
        color: #0f172a;
        border-top: 2px solid #64748b;
    }
    .official-table tr:hover:not(.avg-row) td {
        background-color: #f8fafc;
    }

    /* Box Info Nasabah Resmi */
    .customer-info-box {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 16px;
    }
    .customer-info-table {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.85rem;
    }
    .customer-info-table td {
        padding: 5px 8px;
        vertical-align: top;
    }
    .customer-info-table .label {
        width: 170px;
        font-weight: 600;
        color: #475569;
    }
    .customer-info-table .value {
        background-color: #E2EFDA;
        font-weight: 600;
        color: #1e293b;
        border: 1px solid #94a3b8;
        border-radius: 4px;
        padding: 4px 10px;
    }

    /* Tombol Unduh Utama */
    div.stDownloadButton > button {
        background-color: #0052cc !important;
        color: #ffffff !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        border: none !important;
        padding: 10px 18px !important;
        box-shadow: 0 2px 4px rgba(0, 82, 204, 0.2) !important;
        transition: background-color 0.15s ease !important;
        width: 100%;
    }
    div.stDownloadButton > button:hover {
        background-color: #0747a6 !important;
    }

    /* Sembunyikan elemen streamlit yang tidak perlu */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# Format Rupiah Helper
def format_rp(val):
    if val is None or pd.isna(val):
        return "Rp 0,00"
    return f"Rp {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def format_num(val):
    if val is None or pd.isna(val):
        return "0"
    return f"{int(val):,}".replace(",", ".")

# Helper Badge Bank
def get_bank_badge(bank_name):
    b = bank_name.upper()
    if "MANDIRI" in b:
        return '<span class="bank-pill bank-mandiri">MANDIRI</span>'
    elif "BCA" in b:
        return '<span class="bank-pill bank-bca">BCA</span>'
    elif "BRI" in b:
        return '<span class="bank-pill bank-bri">BRI</span>'
    elif "BNI" in b:
        return '<span class="bank-pill bank-bni">BNI</span>'
    elif "PERMATA" in b:
        return '<span class="bank-pill bank-permata">PERMATA</span>'
    elif "NOBU" in b:
        return '<span class="bank-pill bank-nobu">NOBU</span>'
    else:
        return f'<span class="bank-pill bank-umum">{bank_name}</span>'

# ==============================================================================
# 2. SIDEBAR - FORM INPUT DATA NASABAH & DOKUMEN
# ==============================================================================
with st.sidebar:
    st.markdown("### 📋 Input Data Nasabah")
    st.caption("Lengkapi identitas nasabah untuk dicetak pada dokumen PDF & Excel.")

    cabang = st.text_input("Cabang", value="", placeholder="Contoh: Jakarta Thamrin")
    nama_cust = st.text_input("Nama Cust", value="", placeholder="Nama lengkap debitur/nasabah")
    no_rekening = st.text_input("Nomor Rekening", value="", placeholder="Contoh: 123-00-1234567-8")

    bank_options = ["BANK MANDIRI", "BANK BCA", "BANK BRI", "BANK BNI", "BANK PERMATA", "BANK NOBU", "LAINNYA"]
    nama_bank_selected = st.selectbox("Nama Bank", options=bank_options, index=0)
    if nama_bank_selected == "LAINNYA":
        nama_bank = st.text_input("Sebutkan Nama Bank", value="").strip().upper()
    else:
        nama_bank = nama_bank_selected

    nama_pemegang_rek = st.text_input("Nama Pemegang Rekening", value="", placeholder="Sesuai buku tabungan / e-statement")

    st.markdown("---")
    st.markdown("### ✍️ Petugas & Catatan")
    nama_so = st.text_input("Nama SO (Sales Officer)", value="", placeholder="Nama Sales Officer / AO")
    note_oh = st.text_area("Note OH (Operation Head)", value="-", height=75, placeholder="Catatan analisa mutasi...")

    header_input = {
        "cabang": cabang,
        "nama_cust": nama_cust,
        "no_rekening": no_rekening,
        "nama_bank": nama_bank,
        "nama_pemegang_rek": nama_pemegang_rek
    }

# ==============================================================================
# 3. KONTEN UTAMA: HEADER & FILE UPLOADER
# ==============================================================================
st.markdown("""
<div class="header-banner">
    <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
        <div>
            <h1 class="header-title">🏦 Form Validasi Mutasi Rekening</h1>
            <p class="header-desc">Ekstraksi mutasi multi-bank otomatis, verifikasi saldo, dan kalkulasi rasio mutasi debet/kredit resmi.</p>
        </div>
        <div>
            <span class="bank-pill bank-mandiri">MANDIRI</span>
            <span class="bank-pill bank-bca">BCA</span>
            <span class="bank-pill bank-bri">BRI</span>
            <span class="bank-pill bank-bni">BNI</span>
            <span class="bank-pill bank-permata">PERMATA</span>
            <span class="bank-pill bank-nobu">NOBU</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# Upload Area
st.markdown("#### 📂 Unggah Rekening Koran (PDF)")
uploaded_files = st.file_uploader(
    "Pilih 1 s/d 3 berkas Rekening Koran (PDF):",
    type=["pdf"],
    accept_multiple_files=True,
    help="Dukungan parser spesifik: Mandiri, BCA, BRI, BNI, Permata Bank, Nobu, dan Umum."
)

if uploaded_files:
    # Action Button
    col_btn, col_count = st.columns([2, 5])
    with col_btn:
        process_btn = st.button("⚡ Proses Analisis Rekening", type="primary", use_container_width=True)
    with col_count:
        st.info(f"📁 Terdeteksi **{len(uploaded_files)} berkas** siap dianalisis. Tekan tombol **Proses Analisis Rekening** di samping.")

    # Simpan hasil analisis di session state
    if process_btn:
        with st.spinner("Sedang membaca halaman PDF & mengekstrak data transaksi..."):
            resume_list = []
            for up_file in uploaded_files:
                try:
                    data_ext = parse_rekening_universal(up_file, filename=up_file.name)
                    resume_list.append(data_ext)
                except Exception as e:
                    st.error(f"Gagal memproses berkas {up_file.name}: {str(e)}")

            if resume_list:
                # Urutkan secara kronologis
                sorted_resume = sort_resume_chronological(resume_list)
                st.session_state["resume_list"] = sorted_resume
                st.session_state["processed_header"] = header_input
                st.session_state["processed_nama_so"] = nama_so
                st.session_state["processed_note_oh"] = note_oh
                st.success(f"Berhasil mengekstrak {len(sorted_resume)} periode rekening koran!")

# ==============================================================================
# 4. TAMPILAN HASIL ANALISIS & RESUME MUTASI
# ==============================================================================
if "resume_list" in st.session_state and st.session_state["resume_list"]:
    resume_list = st.session_state["resume_list"]
    n_period = len(resume_list)

    # Hitung total & rata-rata
    totals = {
        "f_db": sum(item.get("freq_db", 0) for item in resume_list),
        "f_cr": sum(item.get("freq_cr", 0) for item in resume_list),
        "m_db": sum(item.get("mutasi_db", 0.0) for item in resume_list),
        "m_cr": sum(item.get("mutasi_cr", 0.0) for item in resume_list),
        "s_max": sum(item.get("saldo_max", 0.0) for item in resume_list),
        "s_avg": sum(item.get("saldo_avg", 0.0) for item in resume_list),
        "s_min": sum(item.get("saldo_min", 0.0) for item in resume_list),
    }

    avg_f_db = round(totals["f_db"] / n_period) if n_period else 0
    avg_f_cr = round(totals["f_cr"] / n_period) if n_period else 0
    avg_m_db = totals["m_db"] / n_period if n_period else 0.0
    avg_m_cr = totals["m_cr"] / n_period if n_period else 0.0
    avg_s_max = totals["s_max"] / n_period if n_period else 0.0
    avg_s_avg = totals["s_avg"] / n_period if n_period else 0.0
    avg_s_min = totals["s_min"] / n_period if n_period else 0.0

    # Ringkasan KPI Cards
    st.markdown("### 📊 Ringkasan Mutasi Eksekutif")
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Total Mutasi Kredit (Masuk)</div>
            <div class="metric-value" style="color: #16a34a;">{format_rp(totals['m_cr'])}</div>
            <div class="metric-sub">Total {totals['f_cr']}x transaksi (Rata-rata {format_rp(avg_m_cr)}/bln)</div>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Total Mutasi Debet (Keluar)</div>
            <div class="metric-value" style="color: #dc2626;">{format_rp(totals['m_db'])}</div>
            <div class="metric-sub">Total {totals['f_db']}x transaksi (Rata-rata {format_rp(avg_m_db)}/bln)</div>
        </div>
        """, unsafe_allow_html=True)

    with c3:
        net_flow = totals['m_cr'] - totals['m_db']
        flow_color = "#16a34a" if net_flow >= 0 else "#dc2626"
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Net Cash Flow (Cr - Db)</div>
            <div class="metric-value" style="color: {flow_color};">{format_rp(net_flow)}</div>
            <div class="metric-sub">Rata-rata per bulan: {format_rp(net_flow / n_period if n_period else 0)}</div>
        </div>
        """, unsafe_allow_html=True)

    with c4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Rata-Rata Saldo Mengendap</div>
            <div class="metric-value" style="color: #0052cc;">{format_rp(avg_s_avg)}</div>
            <div class="metric-sub">Tertinggi: {format_rp(avg_s_max)} | Terendah: {format_rp(avg_s_min)}</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Identitas & Tabel Resmi Form Validasi Mutasi
    col_table, col_actions = st.columns([7, 3])

    with col_table:
        st.markdown("### 📑 Resume Mutasi Rekening Resmi")

        # Box Header Nasabah
        cur_header = st.session_state.get("processed_header", header_input)
        st.markdown(f"""
        <div class="customer-info-box">
            <table class="customer-info-table">
                <tr>
                    <td class="label">Cabang:</td>
                    <td class="value">{cur_header.get('cabang') or '-'}</td>
                    <td class="label" style="padding-left: 20px;">Nomor Rekening:</td>
                    <td class="value">{cur_header.get('no_rekening') or '-'}</td>
                </tr>
                <tr>
                    <td class="label">Nama Cust:</td>
                    <td class="value">{cur_header.get('nama_cust') or '-'}</td>
                    <td class="label" style="padding-left: 20px;">Nama Bank:</td>
                    <td class="value">{cur_header.get('nama_bank') or '-'}</td>
                </tr>
                <tr>
                    <td class="label">Nama Pemegang Rek.:</td>
                    <td class="value" colspan="3">{cur_header.get('nama_pemegang_rek') or '-'}</td>
                </tr>
            </table>
        </div>
        """, unsafe_allow_html=True)

        # Tabel Form Mutasi
        table_rows_html = ""
        for it in resume_list:
            b_name = it.get("bulan", "-")
            table_rows_html += f"""
            <tr>
                <td class="text-center"><b>{b_name}</b></td>
                <td class="text-right">{format_num(it.get('freq_db', 0))}</td>
                <td class="text-right">{format_num(it.get('freq_cr', 0))}</td>
                <td class="text-right">{format_rp(it.get('mutasi_db', 0.0))}</td>
                <td class="text-right">{format_rp(it.get('mutasi_cr', 0.0))}</td>
                <td class="text-right">{format_rp(it.get('saldo_max', 0.0))}</td>
                <td class="text-right">{format_rp(it.get('saldo_avg', 0.0))}</td>
                <td class="text-right">{format_rp(it.get('saldo_min', 0.0))}</td>
            </tr>
            """

        # Baris Rata-Rata
        avg_row_html = f"""
        <tr class="avg-row">
            <td class="text-center"><b>Rata-Rata</b></td>
            <td class="text-right"><b>{format_num(avg_f_db)}</b></td>
            <td class="text-right"><b>{format_num(avg_f_cr)}</b></td>
            <td class="text-right"><b>{format_rp(avg_m_db)}</b></td>
            <td class="text-right"><b>{format_rp(avg_m_cr)}</b></td>
            <td class="text-right"><b>{format_rp(avg_s_max)}</b></td>
            <td class="text-right"><b>{format_rp(avg_s_avg)}</b></td>
            <td class="text-right"><b>{format_rp(avg_s_min)}</b></td>
        </tr>
        """

        full_table_html = f"""
        <table class="official-table">
            <thead>
                <tr>
                    <th rowspan="2" style="vertical-align: middle; width: 12%;">Bulan</th>
                    <th colspan="2">Frekuensi</th>
                    <th colspan="2">Mutasi</th>
                    <th colspan="3">Saldo</th>
                </tr>
                <tr>
                    <th style="width: 8%;">Debet</th>
                    <th style="width: 8%;">Kredit</th>
                    <th style="width: 18%;">Debet</th>
                    <th style="width: 18%;">Kredit</th>
                    <th style="width: 12%;">Tertinggi</th>
                    <th style="width: 12%;">Rata-Rata</th>
                    <th style="width: 12%;">Terendah</th>
                </tr>
            </thead>
            <tbody>
                {table_rows_html}
                {avg_row_html}
            </tbody>
        </table>
        """
        st.markdown(full_table_html, unsafe_allow_html=True)

        # Note OH & SO Box
        st.markdown(f"""
        <div style="margin-top: 14px; padding: 12px 16px; background-color: #ffffff; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 0.85rem;">
            <b>Note OH:</b> {st.session_state.get('processed_note_oh', note_oh) or '-'}<br>
            <div style="margin-top: 10px; display: flex; justify-content: space-between;">
                <span><b>SO:</b> {st.session_state.get('processed_nama_so', nama_so) or '-'}</span>
                <span><b>Operation Head:</b> ____________________</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col_actions:
        st.markdown("### 📥 Unduh Dokumen")
        st.caption("Unduh dokumen resmi yang siap dicetak dan ditandatangani.")

        # Generate PDF
        pdf_bytes = generate_form_pdf(
            None,
            cur_header,
            resume_list,
            note_oh=st.session_state.get("processed_note_oh", note_oh),
            nama_so=st.session_state.get("processed_nama_so", nama_so)
        )

        st.download_button(
            label="📄 Unduh PDF Form Validasi",
            data=pdf_bytes,
            file_name="FORM_VALIDASI_MUTASI_REKENING.pdf",
            mime="application/pdf",
            use_container_width=True
        )

        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

        # Generate Excel
        excel_bytes = generate_form_excel(
            None,
            cur_header,
            resume_list,
            note_oh=st.session_state.get("processed_note_oh", note_oh),
            nama_so=st.session_state.get("processed_nama_so", nama_so)
        )

        st.download_button(
            label="📊 Unduh Excel Form Validasi",
            data=excel_bytes,
            file_name="FORM_VALIDASI_MUTASI_REKENING.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

        # Detail Ringkasan Berkas
        st.markdown("---")
        st.markdown("#### 🔍 Informasi Berkas")
        for item in resume_list:
            b_badge = get_bank_badge(item.get("bank", "UMUM"))
            st.markdown(f"""
            <div style="padding: 8px 12px; margin-bottom: 8px; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 6px; font-size: 0.8rem;">
                <b>{item.get('bulan')}</b> - {b_badge}<br>
                <span style="color: #64748b;">File: {item.get('filename')}</span><br>
                <span style="color: #475569;">Saldo Awal: <b>{format_rp(item.get('opening_bal'))}</b></span><br>
                <span style="color: #475569;">Total Baris: <b>{len(item.get('tx_records', []))} transaksi</b></span>
            </div>
            """, unsafe_allow_html=True)

    # ==============================================================================
    # 5. GRAFIK TREN MUTASI & SALDO
    # ==============================================================================
    st.markdown("---")
    st.markdown("### 📈 Visualisasi Dinamika Keuangan")

    chart_data = []
    for it in resume_list:
        chart_data.append({
            "Bulan": it.get("bulan", "-"),
            "Mutasi Debet (Keluar)": it.get("mutasi_db", 0.0),
            "Mutasi Kredit (Masuk)": it.get("mutasi_cr", 0.0),
            "Saldo Tertinggi": it.get("saldo_max", 0.0),
            "Saldo Rata-Rata": it.get("saldo_avg", 0.0),
            "Saldo Terendah": it.get("saldo_min", 0.0),
        })
    df_chart = pd.DataFrame(chart_data)

    col_c1, col_c2 = st.columns(2)
    with col_c1:
        st.markdown("##### Perbandingan Mutasi Kredit vs Debet")
        if not df_chart.empty:
            st.bar_chart(
                df_chart.set_index("Bulan")[["Mutasi Kredit (Masuk)", "Mutasi Debet (Keluar)"]],
                color=["#16a34a", "#dc2626"]
            )

    with col_c2:
        st.markdown("##### Pergerakan Saldo Bulanan")
        if not df_chart.empty:
            st.line_chart(
                df_chart.set_index("Bulan")[["Saldo Tertinggi", "Saldo Rata-Rata", "Saldo Terendah"]],
                color=["#0284c7", "#0052cc", "#64748b"]
            )

    # ==============================================================================
    # 6. RINCIAN TRANSAKSI PER BULAN
    # ==============================================================================
    st.markdown("---")
    st.markdown("### 🔍 Rincian Transaksi Per Bulan")
    st.caption("Pilih tab bulan di bawah ini untuk melihat detail buku transaksi mutasi secara transparan.")

    tab_titles = [f"📅 {item.get('bulan', f'Bulan {i+1}')}" for i, item in enumerate(resume_list)]
    month_tabs = st.tabs(tab_titles)

    for i, tab in enumerate(month_tabs):
        with tab:
            cur_month_data = resume_list[i]
            records = cur_month_data.get("tx_records", [])

            col_t1, col_t2, col_t3, col_t4 = st.columns(4)
            col_t1.metric("Bank", cur_month_data.get("bank", "-"))
            col_t2.metric("Saldo Awal", format_rp(cur_month_data.get("opening_bal")))
            col_t3.metric("Frekuensi Transaksi", f"{len(records)} baris")
            col_t4.metric("Saldo Rata-Rata", format_rp(cur_month_data.get("saldo_avg")))

            if records:
                df_tx = pd.DataFrame(records)
                # Rapikan tampilan tabel transaksi
                df_tx_display = df_tx.copy()
                df_tx_display.columns = ["Tanggal", "Debet (Rp)", "Kredit (Rp)", "Saldo (Rp)"]
                
                # Format Rupiah
                df_tx_display["Debet (Rp)"] = df_tx_display["Debet (Rp)"].apply(lambda x: format_rp(x) if x > 0 else "-")
                df_tx_display["Kredit (Rp)"] = df_tx_display["Kredit (Rp)"].apply(lambda x: format_rp(x) if x > 0 else "-")
                df_tx_display["Saldo (Rp)"] = df_tx_display["Saldo (Rp)"].apply(lambda x: format_rp(x) if x is not None else "-")

                st.dataframe(
                    df_tx_display,
                    use_container_width=True,
                    height=360,
                    hide_index=True
                )
            else:
                st.warning("Tidak ada data transaksi yang ditemukan pada bulan ini.")

else:
    # State Awal Saat Belum Ada Berkas yang Diunggah
    st.markdown("""
    <div style="background-color: #ffffff; border: 1px dashed #cbd5e1; border-radius: 12px; padding: 40px; text-align: center; margin-top: 20px;">
        <div style="font-size: 3rem; margin-bottom: 12px;">📊</div>
        <h3 style="color: #0f172a; margin-bottom: 8px;">Siap Memproses Rekening Koran</h3>
        <p style="color: #64748b; max-width: 620px; margin: 0 auto 20px auto; font-size: 0.95rem;">
            Unggah 1 s/d 3 berkas PDF rekening koran di atas (Mandiri, BCA, BRI, BNI, Permata, atau Nobu), 
            lengkapi data nasabah pada bilah samping (sidebar), lalu tekan tombol <b>Proses Analisis Rekening</b>.
        </p>
        <div style="display: flex; justify-content: center; gap: 20px; flex-wrap: wrap; margin-top: 15px;">
            <div style="text-align: left; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px 20px; width: 260px;">
                <b style="color: #0052cc;">✓ Deteksi Bank Otomatis</b><br>
                <span style="font-size: 0.82rem; color: #64748b;">Mengenali layout spesifik Bank Mandiri, BCA, BRI, BNI, Permata, dan Nobu.</span>
            </div>
            <div style="text-align: left; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px 20px; width: 260px;">
                <b style="color: #16a34a;">✓ Perhitungan Presisi</b><br>
                <span style="font-size: 0.82rem; color: #64748b;">Ekstraksi saldo awal, mutasi debet/kredit, dan rata-rata saldo mengendap.</span>
            </div>
            <div style="text-align: left; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 14px 20px; width: 260px;">
                <b style="color: #ea580c;">✓ Ekspor 1-Klik</b><br>
                <span style="font-size: 0.82rem; color: #64748b;">Unduh Form Validasi Mutasi Rekening resmi dalam format PDF dan Excel.</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
