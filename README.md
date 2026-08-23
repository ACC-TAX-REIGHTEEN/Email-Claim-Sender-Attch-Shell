# 📧 Email Claim Sender — Shell Attachment

> **Kirim email klaim Shell satu per invoice secara otomatis — cukup siapkan file Excel dan PDF, jalankan skrip, semua email terkirim beserta lampirannya**

Skrip Python dua langkah yang membaca daftar invoice klaim dari file Excel macro (`Surat Saffiela - 130326.xlsm`), lalu mengirimkan satu email per invoice ke penerima yang dikonfigurasi via Gmail SMTP — masing-masing disertai file PDF lampiran yang disesuaikan secara otomatis berdasarkan nomor invoice.

---

## 📋 Daftar Isi

- [Gambaran Umum](#-gambaran-umum)
- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Struktur Folder & File](#-struktur-folder--file)
- [Cara Penggunaan](#-cara-penggunaan)
- [Alur Kerja Pipeline](#-alur-kerja-pipeline)
- [Konfigurasi `config.conf`](#-konfigurasi-configconf)
- [Format File Excel (`Surat Saffiela`)](#-format-file-excel-surat-saffiela)
- [Konvensi Penamaan File PDF](#-konvensi-penamaan-file-pdf)
- [Setup Gmail App Password](#-setup-gmail-app-password)
- [Penyesuaian untuk Pengguna Lain](#-penyesuaian-untuk-pengguna-lain)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## 🗂️ Gambaran Umum

Proses pengiriman klaim Shell melibatkan banyak invoice sekaligus — masing-masing membutuhkan email terpisah dengan lampiran PDF yang berbeda. Mengerjakan ini secara manual (buka email, tulis subject, pilih PDF, kirim, ulangi) memakan waktu dan rentan lewat.

Skrip ini mengotomasi seluruh proses:

```
Surat Saffiela - 130326.xlsm  +  INV0012026.pdf, INV0022026.pdf, ...
              ↓
   Baca daftar invoice dari sheet "Isian"
              ↓
   Per invoice: cari PDF → kirim email via Gmail
              ↓
   Semua email terkirim, file sementara dibersihkan
```

---

## ✨ Fitur Utama

- **Satu email per invoice** — Setiap baris yang memiliki `No Invoice Klaim` dan `Nama Program Klaim` menghasilkan satu email terpisah.
- **Subject otomatis** — Format subject: `INV PT. ABCXYZ {No Invoice Klaim}` dibentuk otomatis dari data Excel.
- **Body HTML dengan placeholder** — Template body di `config.conf` mendukung variabel `{nama_program}` yang diisi otomatis dari data tiap baris.
- **Attachment PDF otomatis** — Nama file PDF diturunkan dari nomor invoice (strip karakter `-` dan `/`). Jika file PDF tidak ditemukan, baris tersebut dilewati dengan pesan peringatan.
- **Satu koneksi SMTP untuk semua email** — Seluruh pengiriman dilakukan dalam satu sesi SMTP SSL (port 465), lebih efisien daripada membuka koneksi baru per email.
- **Auto-detect header** — Skrip mencari posisi kolom `No Invoice Klaim` dan `Nama Program Klaim` secara dinamis, tidak bergantung pada nomor baris atau kolom yang tetap.
- **Auto-cleanup** — Semua file `.xlsx`, `.xls`, dan `.pdf` di folder `Dapur/` serta semua `.pdf` di folder utama dihapus otomatis setelah proses selesai.
- **CC support** — Mendukung pengiriman dengan CC ke satu atau lebih alamat email.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan

```bash
pip install openpyxl
```

| Library | Kegunaan |
|---|---|
| `openpyxl` | Baca `.xlsm` (data_only) dan buat `isian_temp.xlsx` |
| `smtplib`, `ssl`, `email.message` | Kirim email via Gmail SMTP SSL (semua sudah di standard library) |
| `configparser`, `os`, `glob`, `shutil`, `subprocess` | Utilitas (standard library) |

### Akun Gmail dengan App Password
Akun Gmail pengirim harus menggunakan **App Password** (bukan password biasa). Lihat [Setup Gmail App Password](#-setup-gmail-app-password).

---

## 📁 Struktur Folder & File

```
📦 Email-Claim-Sender/
│
├── 📄 Jalankan Sender.py                  ← Titik masuk utama. Jalankan ini
│
├── 📄 Surat Saffiela - 130326.xlsm        ← [INPUT] File Excel klaim (nama harus persis)
├── 📄 INV0012026.pdf                      ← [INPUT] PDF lampiran per invoice
├── 📄 INV0022026.pdf                      ← [INPUT] PDF lampiran per invoice
├── 📄 (dst.)                              ← Satu file PDF per invoice
│
└── 📁 Dapur/                              ← Folder pipeline (jangan diubah)
    ├── 📄 __init__.py
    ├── 📄 1_EkstrakData.py               ← Ekstrak sheet "Isian" → isian_temp.xlsx
    ├── 📄 2_GmailSender.py               ← Kirim email via Gmail SMTP
    └── 📄 config.conf                    ← Konfigurasi email (SMTP, penerima, body)
```

> Semua file input (`*.xlsm` dan `*.pdf`) diletakkan di **folder utama** (sejajar dengan `Jalankan Sender.py`).

---

## 🚀 Cara Penggunaan

### Langkah 1 — Siapkan file Excel klaim

Pastikan file `Surat Saffiela - 130326.xlsm` ada di folder utama dengan sheet `Isian` yang sudah terisi data invoice. Lihat format yang dibutuhkan di [Format File Excel](#-format-file-excel-surat-saffiela).

### Langkah 2 — Siapkan file PDF lampiran

Letakkan semua file PDF di folder utama. Nama file PDF harus mengikuti konvensi penamaan dari nomor invoice. Lihat [Konvensi Penamaan File PDF](#-konvensi-penamaan-file-pdf).

### Langkah 3 — Isi `config.conf`

```ini
[SMTP]
sender_email = emailanda@gmail.com
sender_password = xxxx xxxx xxxx xxxx    ; App Password (bukan password Gmail biasa)

[RECIPIENT]
to_email = tujuan@perusahaan.com
cc_email = tujuan_cc@perusahaan.com

[CONTENT]
body = <b>Dear Mbak Afika,</b><br><br>...{nama_program}...
```

### Langkah 4 — Jalankan

```bash
python "Jalankan Sender.py"
```

### Langkah 5 — Pantau output terminal

```
--> Memulai eksekusi 1_EkstrakData.py
--> Memulai proses ekstraksi file
--> Menyalin data ke sheet baru
--> Melakukan auto-fit pada kolom
--> File isian_temp.xlsx berhasil dibuat dan disimpan
--> Memulai eksekusi 2_GmailSender.py
--> Membaca file konfigurasi config.conf
--> Membaca data dari isian_temp.xlsx
--> Membuka koneksi SMTP Gmail
--> Email untuk INV-001/2026 berhasil dikirim beserta lampiran INV0012026.pdf
--> Email untuk INV-002/2026 berhasil dikirim beserta lampiran INV0022026.pdf
--> File PDF INV0032026.pdf untuk invoice INV-003/2026 tidak ditemukan
--> Seluruh proses pengiriman selesai
--> Semua proses telah selesai dijalankan.
--> Tekan enter untuk keluar.
```

---

## 🔄 Alur Kerja Pipeline

```
[Jalankan Sender.py]
   │
   ├─── Validasi: folder Dapur/ + 4 file syarat ada
   ├─── Bersihkan Dapur/ dari *.xls*, *.pdf lama
   ├─── Salin ke Dapur/:
   │       Surat Saffiela - 130326.xlsm (dari folder utama)
   │       *.pdf / *.PDF (semua file PDF dari folder utama)
   │    Jika tidak ada file yang disalin → berhenti
   │
   ├─── [1] 1_EkstrakData.py
   │       Buka Surat Saffiela - 130326.xlsm (data_only=True, tanpa macro)
   │       Cari sheet "Isian"
   │       Salin seluruh nilai sel ke workbook baru
   │       Auto-fit lebar semua kolom
   │       Simpan sebagai isian_temp.xlsx
   │
   ├─── [2] 2_GmailSender.py
   │       Baca config.conf (SMTP, penerima, template body)
   │       Baca isian_temp.xlsx
   │       Cari posisi header "No Invoice Klaim" dan "Nama Program Klaim"
   │       Buka koneksi SMTP SSL ke smtp.gmail.com:465
   │       Per baris data (setelah header):
   │         Ambil No Invoice Klaim & Nama Program Klaim
   │         Lewati jika salah satu kosong
   │         Buat nama PDF: strip '-' dan '/' dari nomor invoice → + ".pdf"
   │         Jika PDF ada: kirim email dengan attachment
   │         Jika tidak: cetak peringatan & lanjut ke baris berikutnya
   │       Tutup koneksi SMTP
   │
   ├─── Bersihkan Dapur/ dari *.xls*, *.pdf
   └─── Bersihkan folder utama dari *.pdf
```

---

## ⚙️ Konfigurasi `config.conf`

```ini
[SMTP]
sender_email = emailanda@gmail.com
sender_password = xxxx xxxx xxxx xxxx

[RECIPIENT]
to_email = tujuan@perusahaan.com
cc_email = tujuan_cc@perusahaan.com

[CONTENT]
body = <b>Dear Mbak Afika,</b><br><br>Berikut terlampir data klaim "{nama_program}".<br>Mohon untuk dapat diproseskan.<br>Terima kasih.<br><br><b>Regards,</b><br><b>Saffiela</b>
```

| Seksi | Key | Keterangan |
|---|---|---|
| `[SMTP]` | `sender_email` | Alamat Gmail pengirim |
| `[SMTP]` | `sender_password` | App Password Gmail (format: 4 blok 4 huruf dipisah spasi) |
| `[RECIPIENT]` | `to_email` | Alamat email penerima utama (To:) |
| `[RECIPIENT]` | `cc_email` | Alamat email CC. Kosongkan jika tidak ada CC |
| `[CONTENT]` | `body` | Template body HTML. Gunakan `{nama_program}` sebagai placeholder |

**Placeholder yang tersedia di `body`:**

| Placeholder | Diganti dengan |
|---|---|
| `{nama_program}` | Nilai kolom `Nama Program Klaim` dari baris yang sedang diproses |

**Format body mendukung tag HTML:** `<b>`, `<br>`, `<i>`, `<a href>`, dan tag HTML email standar lainnya.

---

## 📋 Format File Excel (`Surat Saffiela`)

File `Surat Saffiela - 130326.xlsm` harus memiliki sheet bernama **`Isian`** dengan setidaknya dua kolom header:

| Header (nama persis) | Isi |
|---|---|
| `No Invoice Klaim` | Nomor invoice, misalnya `INV-001/2026` |
| `Nama Program Klaim` | Nama program klaim, misalnya `Shell Helix Program Q1` |

**Aturan:**
- Nama header harus **persis sama** (case-sensitive) dengan `No Invoice Klaim` dan `Nama Program Klaim`.
- Header boleh berada di baris mana saja — skrip mencari posisinya secara otomatis dengan scan seluruh sheet.
- Baris yang salah satu kolomnya kosong akan dilewati.
- Kolom lain boleh ada; hanya dua kolom di atas yang digunakan untuk pengiriman.

**Contoh struktur sheet `Isian`:**

| No Invoice Klaim | Nama Program Klaim | Keterangan |
|---|---|---|
| INV-001/2026 | Shell Helix Program Q1 | Klaim bulan Januari |
| INV-002/2026 | Shell Advance Program | Klaim bulan Februari |

> **Catatan:** File `.xlsm` dibaca dengan `data_only=True` — artinya **nilai sel**, bukan formula. Pastikan semua nilai yang dibutuhkan sudah berupa nilai statis, bukan formula yang belum dievaluasi (buka file di Excel dan simpan sekali sebelum menjalankan skrip jika perlu).

---

## 📄 Konvensi Penamaan File PDF

Nama file PDF diturunkan secara otomatis dari nilai `No Invoice Klaim` dengan aturan:

```
Nomor invoice → hapus semua karakter '-' dan '/' → tambahkan ".pdf"
```

**Contoh konversi:**

| No Invoice Klaim | Nama file PDF yang dicari |
|---|---|
| `INV-001/2026` | `INV0012026.pdf` |
| `INV-002/2026` | `INV0022026.pdf` |
| `SHELL/001-2026` | `SHELL0012026.pdf` |
| `INV001` | `INV001.pdf` |

> Semua file PDF harus ada di folder utama (sejajar dengan `Jalankan Sender.py`) **sebelum** menjalankan skrip. Jika file PDF tidak ditemukan, invoice tersebut dilewati dan pesan peringatan ditampilkan di terminal.

---

## 🔑 Setup Gmail App Password

Gmail tidak mengizinkan login dengan password biasa dari skrip eksternal. Gunakan **App Password**:

### Langkah-langkah

1. Masuk ke [myaccount.google.com](https://myaccount.google.com)
2. Pilih **Security** → **2-Step Verification** → pastikan sudah aktif
3. Di halaman Security yang sama, cari **App Passwords** (atau buka langsung: [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords))
4. Pilih **Select app → Mail** dan **Select device → Other (Custom name)**
5. Beri nama (misal: `Email Claim Sender`) → klik **Generate**
6. Salin **16-karakter App Password** yang muncul (format: `xxxx xxxx xxxx xxxx`)
7. Tempel ke `config.conf` di `sender_password`

> ⚠️ App Password hanya ditampilkan sekali. Simpan dengan aman setelah di-generate.

---

## 🔧 Penyesuaian untuk Pengguna Lain

Proyek ini awalnya dibuat untuk penggunaan spesifik. Dua nilai yang perlu diubah jika digunakan oleh orang lain:

### 1. Nama file Excel (hardcoded di `1_EkstrakData.py`)

Buka `Dapur/1_EkstrakData.py`, ubah baris 3:

```python
# Sebelum
source_file = "Surat Saffiela - 130326.xlsm"

# Sesudah (sesuaikan dengan nama file Anda)
source_file = "Surat NamaAnda - DDMMYY.xlsm"
```

Kemudian sesuaikan juga di `Jalankan Sender.py`, baris yang mencari file sumber:

```python
# Sebelum
file_sumber_pola = ["Surat Saffiela - 130326.xlsm", "*.pdf", "*.PDF"]

# Sesudah
file_sumber_pola = ["Surat NamaAnda - DDMMYY.xlsm", "*.pdf", "*.PDF"]
```

### 2. Subject email (hardcoded di `2_GmailSender.py`)

Buka `Dapur/2_GmailSender.py`, ubah baris subject:

```python
# Sebelum
subject = "INV PT. ABCXYZ " + no_invoice

# Sesudah (sesuaikan nama perusahaan)
subject = "INV PT. NAMA PERUSAHAAN " + no_invoice
```

---

## 🛠️ Troubleshooting

### ❌ `File sumber tidak ditemukan untuk diproses`
File `Surat Saffiela - 130326.xlsm` tidak ada di folder utama, atau tidak ada file PDF sama sekali. Pastikan keduanya ada sebelum menjalankan skrip.

### ❌ `Sheet Isian tidak ditemukan di dalam file sumber`
Sheet di file `.xlsm` tidak bernama `Isian`. Buka file di Excel, periksa nama tab sheet, dan ganti nama menjadi `Isian` atau ubah `sheet_name` di `1_EkstrakData.py`.

### ❌ `Kolom No Invoice Klaim atau Nama Program Klaim tidak ditemukan`
Nama header di sheet `Isian` tidak cocok. Pastikan persis `No Invoice Klaim` dan `Nama Program Klaim` (perhatikan huruf kapital dan spasi).

### ❌ `[Errno 111] Connection refused` atau timeout saat SMTP
Kemungkinan: (1) koneksi internet bermasalah; (2) port 465 diblokir oleh firewall; (3) `sender_email` belum mengaktifkan 2-Step Verification (diperlukan untuk App Password).

### ❌ `Username and Password not accepted`
App Password salah atau tidak valid. Pastikan: (1) 2-Step Verification aktif di akun Gmail; (2) App Password di-generate ulang jika pernah direvoke; (3) `sender_password` di `config.conf` diisi dengan App Password (16 karakter), bukan password Gmail biasa.

### ❌ `File PDF INV... tidak ditemukan`
Nama file PDF tidak sesuai konvensi atau file belum disalin ke folder utama. Cek [Konvensi Penamaan File PDF](#-konvensi-penamaan-file-pdf) dan pastikan file ada di folder utama sebelum menjalankan skrip.

### ❌ Email terkirim tapi isi `{nama_program}` masih literal (tidak terganti)
Pastikan placeholder di `config.conf` ditulis persis `{nama_program}` (dengan kurung kurawal, huruf kecil semua). Spasi atau huruf berbeda menyebabkan substitusi gagal.

### ❌ File PDF habis terhapus setelah proses selesai
Ini adalah perilaku yang disengaja — orkestrator menghapus semua `.pdf` dari folder utama di akhir proses untuk kebersihan. Simpan salinan PDF di tempat lain sebelum menjalankan skrip jika perlu.

---

## 📌 Catatan Penting

- **File PDF dihapus setelah proses** — Semua `.pdf` di folder utama dihapus otomatis di akhir pipeline. Pastikan sudah ada salinan di tempat lain jika file PDF perlu dipertahankan.
- **File `.xlsm` TIDAK dihapus dari folder utama** — Hanya salinannya di `Dapur/` yang dihapus. File sumber asli tetap aman.
- **Satu koneksi, banyak email** — Semua email dikirim dalam satu sesi SMTP. Jika koneksi terputus di tengah, email yang sudah terkirim sebelumnya tidak dapat dibatalkan.
- **Invoice tanpa PDF dilewati, bukan dihentikan** — Pipeline tidak berhenti jika satu PDF tidak ditemukan; lanjut ke invoice berikutnya dan menampilkan peringatan di terminal.
- **`data_only=True` tidak membaca formula** — File `.xlsm` dibaca tanpa menjalankan macro atau mengevaluasi formula. Buka dan simpan ulang file di Excel terlebih dahulu jika ada sel berformula yang belum terevaluasi.
- **`config.conf` berisi kredensial sensitif** — Tambahkan `Dapur/config.conf` ke `.gitignore`. Jangan commit ke repositori publik.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)
