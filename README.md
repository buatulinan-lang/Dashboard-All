# Dashboard Service Cabang (Offline, Python)

Aplikasi dashboard berbasis Python (Streamlit) yang jalan **sepenuhnya offline**
di komputer sendiri — tidak butuh internet, tidak butuh server, dan bisa
di-update sendiri setiap kali ada data baru cukup dengan upload ulang file
Excel dari dalam aplikasi.

Berisi 2 dashboard dalam 1 aplikasi (tab):

1. **Dashboard Utama** — status pengerjaan (Done/Pending/Cancel/Lainnya),
   tren tahunan, rata-rata transaksi/hari, rekap per cabang.
2. **Dashboard Pending** — breakdown teknisi dengan pending terbanyak dan
   jenis kerusakan yang paling sering menumpuk jadi pending, lengkap dengan
   persentase.

Semua angka mengikuti aturan yang sama seperti versi HTML sebelumnya: baris
yang seluruh kolomnya identik dihitung sebagai 1 transaksi.

## 1. Install Python (sekali saja)

Pastikan Python 3.9 ke atas sudah terpasang. Cek dengan:

```bash
python3 --version
```

Kalau belum ada, download dari https://www.python.org/downloads/

## 2. Install library yang dibutuhkan

Buka Terminal (Mac) / Command Prompt (Windows), masuk ke folder
`dashboard_app` ini, lalu jalankan:

```bash
pip install -r requirements.txt
```

Tunggu sampai selesai (butuh internet untuk sekali ini saja, saat install).

## 3. Jalankan dashboard

Masih di folder yang sama, jalankan:

```bash
streamlit run app.py
```

Browser akan otomatis terbuka ke `http://localhost:8501` — itulah
dashboard-nya. Semua proses jalan lokal di komputer, tidak terkirim ke
mana-mana.

## 4. Upload data

Di panel kiri, klik **"Upload file Excel"** dan pilih file data (format harus
sama seperti `Gabungan_Semua_Cabang.xlsx`: satu sheet per cabang, dengan
kolom seperti `NOMOR PENGIRIMAN PESANAN`, `TGL PENGIRIMAN`,
`STATUS PENGERJAAN`, `NAMA TEKNISI`, `KERUSAKAN UTAMA`, dst).

Setelah upload, dashboard otomatis menghitung ulang. Gunakan filter Tahun /
Bulan / Cabang di panel kiri untuk mempersempit tampilan.

> **Kalau dashboard di-hosting online (Streamlit Community Cloud):** app
> gratis akan "tidur" kalau lama tidak dipakai, dan begitu dibuka lagi, file
> yang sempat diupload akan hilang (harus upload ulang). Supaya tidak perlu
> upload berulang-ulang, pakai skema **data bawaan** di bagian 6 di bawah —
> upload di panel kiri jadi opsional, hanya dipakai kalau mau lihat data
> lain sementara.

## 5. Cara Update Data (Offline)

Dashboard ini **tidak menyimpan data sendiri** — setiap kali dibuka, dia
membaca ulang file Excel yang Anda upload. Jadi "update database" artinya
cukup: siapkan file Excel terbaru, lalu upload ulang. Tidak ada instalasi
ulang, tidak ada ubah kode.

### Skenario A — Anda punya file export baru dari sistem (paling disarankan)

Kalau sumber datanya adalah export/laporan dari sistem service (seperti
`Gabungan_Semua_Cabang.xlsx` yang sudah ada), biasanya setiap kali export
baru sudah otomatis berisi data lama **+** data baru dalam satu file.

Langkah-langkahnya:

1. Export ulang laporan dari sistem seperti biasa, mencakup rentang tanggal
   terbaru (bisa dari awal lagi atau digabung dengan data sebelumnya).
2. Pastikan hasil export masih format yang sama: **satu sheet per cabang**,
   dengan nama kolom persis seperti sebelumnya (`NOMOR PENGIRIMAN PESANAN`,
   `TGL PENGIRIMAN`, `STATUS PENGERJAAN`, `NAMA TEKNISI`,
   `KERUSAKAN UTAMA`, dst).
3. Buka dashboard (`streamlit run app.py` kalau belum jalan).
4. Di panel kiri, klik **"Browse files"** pada uploader dan pilih file
   export terbaru — ini akan **menimpa** file lama yang sedang dibuka.
5. Tunggu proses "Membaca & memproses file Excel..." selesai (bisa 30–90
   detik untuk file besar). Semua angka, filter Tahun/Bulan/Cabang, dan
   ranking otomatis terhitung ulang dari file baru.

Tidak perlu khawatir soal data dobel: baris yang isinya identik persis tetap
otomatis dihitung sebagai 1 transaksi, jadi meng-upload file yang berisi
gabungan data lama + baru itu aman.

### Skenario B — Anda menambah data secara manual di Excel

Kalau tidak ada sistem export dan data ditambah manual:

1. Buka file Excel sumber (mis. `Gabungan_Semua_Cabang.xlsx`) langsung di
   Excel/Google Sheets.
2. Masuk ke **sheet cabang yang sesuai** (nama sheet = nama cabang, ini yang
   dipakai dashboard sebagai label Cabang).
3. Tambahkan baris baru **di bawah baris terakhir**, isi setiap kolom sesuai
   urutan header yang sudah ada — jangan mengubah nama header, jangan
   menyisipkan kolom baru di tengah.
4. Pastikan kolom `TGL PENGIRIMAN` diisi sebagai **tanggal (Date)**, bukan
   teks biasa — kalau formatnya teks, filter Tahun/Bulan di dashboard tidak
   akan mengenali baris tsb.
5. Simpan file (tetap format `.xlsx`, jangan `.xls` atau `.csv`).
6. Upload ulang file tsb ke dashboard seperti Skenario A langkah 3–5.

### Hal yang perlu dihindari

- Jangan mengganti/mengetik ulang nama kolom (header) — kalau berbeda dari
  aslinya, dashboard akan menampilkan pesan error "kolom tidak ditemukan".
- Jangan mengganti nama sheet cabang secara tiba-tiba tanpa alasan — nama
  sheet itu yang muncul sebagai pilihan filter **Cabang**.
- Kalau menambah cabang baru, cukup buat sheet baru dengan nama cabang tsb
  dan header kolom yang sama — dashboard otomatis mendeteksinya tanpa
  perlu ubah kode.

## 6. Data Bawaan (Khusus Dashboard Online) — supaya tidak perlu upload terus

Kalau dashboard di-deploy online di Streamlit Community Cloud, app gratis
akan "tidur" saat lama tidak diakses. Begitu dibuka lagi, semua yang sempat
diupload lewat `st.file_uploader` hilang dari memori — makanya terasa
seperti "harus upload ulang terus".

Solusinya: simpan data langsung di dalam repo GitHub, di file
`data/latest_data.csv.gz`. Kalau file ini ada, `app.py` otomatis memakainya
sebagai data bawaan **tanpa perlu upload apa pun** — panel upload di kiri
jadi opsional, hanya untuk melihat data lain sementara (tidak permanen).

File `data/latest_data.csv.gz` ini sudah saya sertakan dan siap dipakai
(dikonversi dari `Gabungan_Semua_Cabang.xlsx`, ukurannya jauh lebih kecil —
sekitar 7MB — karena sudah dedup dan dikompres, supaya muat diupload lewat
browser GitHub yang batasnya 25MB per file).

### Cara pasang data bawaan ini di repo GitHub

1. Di komputer Anda, pastikan folder `dashboard_app/data/latest_data.csv.gz`
   ada (folder `data` berisi 1 file itu).
2. Buka repo Anda di GitHub → klik **"Add file"** → **"Upload files"**.
3. **Drag seluruh folder `data`** (bukan cuma file-nya) dari komputer Anda
   ke kotak upload di GitHub. Cara ini membuat GitHub otomatis membuat
   folder `data/` di repo dan menaruh file di dalamnya.
4. Isi commit message, misalnya `Tambah data bawaan`, klik **"Commit changes"**.
5. Tunggu Streamlit Cloud redeploy otomatis (atau klik **"Reboot app"** di
   menu titik tiga app Anda). Setelah itu, buka link app-nya — harusnya
   langsung tampil datanya tanpa diminta upload.

### Cara update data bawaan ini nanti (kalau ada data baru)

1. Di folder `dashboard_app` (di komputer Anda, lewat Terminal/Command
   Prompt), jalankan:
   ```bash
   python prepare_data.py path/ke/file_excel_terbaru.xlsx
   ```
   Ini akan menimpa `data/latest_data.csv.gz` dengan data terbaru (otomatis
   dedup & kompres juga).
2. Upload ulang file `data/latest_data.csv.gz` yang baru itu ke GitHub —
   buka file lama-nya di repo (`data/latest_data.csv.gz`), klik ikon pensil
   atau tombol **upload** untuk menimpanya, lalu commit.
3. Reboot app di Streamlit Cloud (atau tunggu redeploy otomatis) — dashboard
   online langsung pakai data terbaru, tanpa siapa pun perlu upload manual.

## 7. Tombol "Buat PPT dari Dashboard"

Di panel kiri paling bawah ada bagian **📤 Buat Presentasi**. Tombol ini
membuat file PowerPoint yang mengikuti template presentasi perusahaan
(logo Madinah Group & MFlash, background dan gaya judul yang sama).

Cara pakai:

1. Atur dulu filter Tahun / Bulan / Cabang sesuai periode yang mau
   dilaporkan — isi PPT mengikuti filter yang sedang aktif.
2. (Opsional) Isi **Nama penyusun**, misalnya
   `ARM MFLASH – BUDIARJA IBRAHIM`. Ini muncul di slide sampul.
3. (Opsional) Pilih slide mana saja yang mau disertakan lewat
   **Slide yang disertakan**.
4. Klik **🎬 Buat PPT dari Dashboard**, tunggu sebentar, lalu klik
   **⬇️ Unduh file PPT**.

Slide yang dihasilkan:

| Slide | Isi |
|---|---|
| Sampul | Logo, judul, periode, cabang, nama penyusun |
| Ringkasan Kinerja | 5 kartu KPI, grafik volume per bulan, catatan utama otomatis |
| Komposisi Status | Donut Done/Cancel/Pending/Lainnya + tabel rincian |
| Rekap Transaksi Harian | Hari aktif, rata-rata/hari, hari tertinggi & terendah, grafik harian, rata-rata per hari dalam pekan |
| Hari Transaksi Tertinggi | Tabel 10 tanggal tersibuk (tanggal, nama hari, jumlah, selisih vs rata-rata) + grafik perbandingan |
| Kinerja per Cabang | Tabel volume, % done, % cancel (dengan penanda warna) |
| Detail Pending | KPI + ranking teknisi & kerusakan |
| Detail Done | KPI + ranking teknisi & kerusakan |
| Detail Cancel | KPI + ranking teknisi & kerusakan |
| Kesimpulan | 4 poin tindak lanjut yang disusun otomatis dari angka |

Grafik harian menyesuaikan otomatis: kalau periodenya panjang (di atas 40
hari) ditampilkan sebagai garis, kalau pendek (misalnya filter satu bulan)
ditampilkan sebagai batang per tanggal agar lebih mudah dibaca.

Semua angka, nama teknisi, cabang, dan kesimpulan dihitung ulang otomatis
dari data — jadi tidak ada yang perlu diketik manual setiap bulan.

Catatan: slide "Kinerja per Cabang" hanya muncul kalau filter Cabang
diatur ke **Semua Cabang** (kalau sudah difilter satu cabang, tabel
perbandingan antar cabang jadi tidak relevan).

### Mengganti logo / background PPT

File gambar ada di folder `assets/`:

- `logo_madinah.png` — logo kiri di sampul
- `logo_mflash.png` — logo kanan di sampul
- `bg.jpg` — background bertekstur semua slide

Timpa file-file itu dengan gambar lain kalau template perusahaan berubah;
tidak perlu mengubah kode.

## Menghentikan aplikasi

Kembali ke Terminal/Command Prompt tempat `streamlit run app.py` dijalankan,
tekan `Ctrl + C`.

---

**Catatan:** file besar (ratusan ribu baris, puluhan MB) bisa butuh waktu
sekitar 30–90 detik untuk diproses saat pertama kali diupload. Setelah itu,
mengganti filter akan terasa instan karena data sudah di-cache di memori.
