# 🏦 AREKO-MANDIRI: Form Validasi Mutasi Rekening & Parser Multi-Bank

Aplikasi web berbasis **Streamlit** modern dengan desain **Light Mode** yang bersih dan profesional untuk melakukan otomatisasi ekstraksi mutasi rekening koran bank (*Bank Statement Parser*), kalkulasi saldo, dan pembuatan **Form Validasi Mutasi Rekening** resmi dalam format **PDF** dan **Excel (.xlsx)**.

Didesain khusus untuk kebutuhan operasional perbankan, analisis pembiayaan, dan komite kredit (*Credit Committee / Operation Head / Sales Officer*).

---

## 🌟 Fitur Utama

- **🎨 Pure Light Mode & Clean Aesthetic:**
  - Antarmuka terang yang elegan dengan palet warna korporat perbankan, kartu metrik KPI, dan tipografi modern.
  - Tabel resume resmi identik dengan standar form validasi operasional (header biru muda `#D9E1F2`, baris rata-rata hijau `#E2EFDA`).
- **🏛️ Dukungan Multi-Bank Otomatis:**
  - **Bank Mandiri** (ekstraksi presisi *opening balance*, mutasi awal, dan *sticky date*).
  - **Bank Central Asia (BCA)** (deteksi mutasi CR / DB, saldo awal, dan frekuensi).
  - **Bank Rakyat Indonesia (BRI)** (parsing format IBIZ / rekening koran finansial).
  - **Bank Negara Indonesia (BNI)** (ekstraksi koordinat word balance & ledger balance).
  - **Bank Permata** (penanganan transaksi reverse, debit/kredit resmi).
  - **Bank Nobu & Bank Lainnya**.
- **📊 Ringkasan Eksekutif & Visualisasi Interaktif:**
  - Kartu KPI: Total Mutasi Kredit (Masuk), Total Mutasi Debet (Keluar), Net Cash Flow, dan Rata-rata Saldo Mengendap.
  - Grafik perbandingan mutasi kredit vs debet bulanan dan grafik fluktuasi saldo (tertinggi, rata-rata, terendah).
- **📑 Rincian Transaksi Transparan:**
  - Tab per bulan dengan tabel mutasi rapi lengkap dengan format mata uang Rupiah (`Rp #.###.###,##`).
- **📥 Ekspor 1-Klik:**
  - Unduh **PDF Form Validasi Mutasi Rekening** siap cetak dan tanda tangan.
  - Unduh **Excel (.xlsx) Form Validasi Mutasi Rekening** lengkap dengan formula mutasi, styling warna, dan pembatas antar bulan.

---

## 📁 Struktur Proyek

```text
AREKO-MANDIRI/
├── .streamlit/
│   └── config.toml           # Konfigurasi tema Light Mode & server Streamlit
├── app.py                    # Aplikasi utama Streamlit
├── parser_engine.py          # Mesin ekstraksi PDF & generator PDF/Excel
├── requirements.txt          # Dependensi Python
├── .gitignore                # File pengecualian Git
└── README.md                 # Dokumentasi proyek
```

---

## 🚀 Panduan Menjalankan Secara Lokal

### 1. Prasyarat
Pastikan Python versi 3.9 atau lebih baru telah terpasang di komputer Anda.

### 2. Pasang Dependensi
Buka terminal di direktori proyek dan jalankan:
```bash
pip install -r requirements.txt
```

### 3. Jalankan Aplikasi Streamlit
```bash
streamlit run app.py
```
Aplikasi akan terbuka secara otomatis di peramban web pada alamat `http://localhost:8501`.

---

## 🌐 Panduan Upload ke GitHub & Publish di Streamlit Cloud

### Langkah 1: Push ke GitHub
Repository target: `https://github.com/darmayanimina-dev/AREKO-MANDIRI.git`

Jalankan perintah berikut di terminal:
```bash
# Inisialisasi git dan buat branch main
git init
git add .
git commit -m "feat: initial commit modern streamlit bank statement parser"
git branch -M main

# Tambahkan remote repository
git remote add origin https://github.com/darmayanimina-dev/AREKO-MANDIRI.git

# Push ke GitHub
git push -u origin main
```

> **Catatan:** Jika remote `origin` sudah pernah ada, cukup jalankan `git remote set-url origin https://github.com/darmayanimina-dev/AREKO-MANDIRI.git` sebelum `git push`.

---

### Langkah 2: Deploy di Streamlit Community Cloud

1. Buka [share.streamlit.io](https://share.streamlit.io/) dan masuk dengan akun GitHub Anda.
2. Klik tombol **"Create app"** (atau **"New app"**).
3. Pilih opsi **"I already have an app"**.
4. Isi data konfigurasi berikut:
   - **Repository:** `darmayanimina-dev/AREKO-MANDIRI`
   - **Branch:** `main`
   - **Main file path:** `app.py`
5. Klik **"Deploy!"**.
6. Aplikasi akan siap digunakan secara online dan memiliki URL publik!

---

## 🛠️ Dependensi

- `streamlit`: Kerangka kerja web UI interaktif
- `pdfplumber`: Ekstraksi teks dan koordinat tabel dari berkas PDF rekening koran
- `reportlab`: Pembuat dokumen PDF form validasi mutasi resmi
- `openpyxl`: Pembuat lembar kerja Excel (.xlsx) dengan formula dan styling warna
- `pandas` & `numpy`: Pengolahan data mutasi dan kalkulasi statistik saldo
