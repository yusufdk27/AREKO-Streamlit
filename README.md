# Form Validasi Mutasi Rekening (AREKO-Streamlit)

Aplikasi berbasis web modern dan minimalis untuk ekstraksi, rekonsiliasi, dan rekapitulasi mutasi rekening koran multi-bank secara otomatis, presisi, dan terstruktur. Menghasilkan dokumen resmi dalam format **PDF** dan spreadsheet interaktif **Excel (.xlsx)** lengkap dengan kolom otorisasi (Sales Officer & Operation Head).

---

## 🚀 Fitur Utama

- **Multi-Bank Parser Otomatis**: Mendukung format rekening koran PDF dari berbagai bank di Indonesia:
  - Bank Mandiri (termasuk Kopra)
  - BCA
  - BNI (termasuk BNI Direct)
  - BRI (termasuk iBiz)
  - BSI (Bank Syariah Indonesia)
  - Bank Permata
  - Bank Nobu
- **Antarmuka Modern & Minimalis**: Tipografi *Plus Jakarta Sans*, tata letak bersih, badge bank, kartu ringkasan KPI (Total Debet, Total Kredit, Rata-rata Saldo), dan tabel pratinjau yang responsif.
- **Dukungan Otorisasi Lengkap**: Input informasi Sales Officer (SO) dan Operation Head (OH) yang otomatis tercetak pada dokumen PDF dan Excel.
- **Export Dokumen Resmi**:
  - **PDF**: Format A4 standar perbankan dengan tanda tangan SO dan Operation Head.
  - **Excel (.xlsx)**: Spreadsheet dinamis dengan rincian transaksi debit/kredit dan formula saldo otomatis.

---

## 🛠️ Instalasi & Menjalankan Lokal

1. **Clone repositori**:
   ```bash
   git clone https://github.com/yusufdk27/AREKO-Streamlit.git
   cd AREKO-Streamlit
   ```

2. **Buat virtual environment (opsional tapi disarankan)**:
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # Di Linux/macOS
   # venv\Scripts\activate   # Di Windows
   ```

3. **Install dependensi**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Jalankan aplikasi Streamlit**:
   ```bash
   streamlit run app.py
   ```

---

## ☁️ Deployment ke Streamlit Community Cloud

1. Buat repositori baru di GitHub (misal: `AREKO-Streamlit`) dan push kode sumber.
2. Buka [share.streamlit.io](https://share.streamlit.io/) dan login dengan akun GitHub.
3. Klik **"New app"**, lalu pilih:
   - **Repository**: `yusufdk27/AREKO-Streamlit`
   - **Branch**: `main`
   - **Main file path**: `app.py`
4. Klik **"Deploy!"**.
