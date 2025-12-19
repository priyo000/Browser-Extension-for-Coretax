# Coretax Easy

Ekstensi ini dirancang untuk mengotomatisasi pengambilan data pajak (e-Faktur) dari website Coretax secara massal, baik dalam format **Excel** maupun **PDF**.

## Fitur Utama

1. **Export Excel (.xls)**: Download seluruh data tabel **sesuai Filter** ke format Excel.
2. **Bulk Download PDF**: Download seluruh file PDF faktur secara otomatis. (Sesuaikan Filter, pastikan seluruh data terdapat file pdf yg dapat di download)

## 1. Instalasi Ekstensi

1. **Extract/Unzip** file yang Anda terima ke sebuah folder di komputer Anda.
   * Pastikan Anda tahu di mana lokasi foldernya.
2. Buka browser **Google Chrome**.
3. Ketik `chrome://extensions` di kolom alamat (address bar) lalu tekan **Enter**.
4. Di pojok kanan atas, hidupkan tombol **Developer mode** (Mode pengembang).
5. Akan muncul tombol baru di kiri atas, klik **Load unpacked**.
6. Pilih **folder hasil ekstrak tadi** (pilih folder luarnya saja).
7. Selesai! Ekstensi "Coretax Easy" akan muncul di daftar aktif.

## 2. Cara Menggunakan

### Persiapan

1. Pastikan ekstensi sudah terinstall dan aktif.
2. Setiap kali Anda menekan "Reload/Refresh" pada tabel data pajak di website, Anda wajib melakukan **Reload Halaman terlebih dahulu.**

### Langkah-langkah

1. **Buka Website Coretax**: Login dan navigasi ke halaman data faktur (Keluaran/Masukan).
2. **Trigger Data**: Lakukan filter atau klik tombol "Cari/Refresh" di tabel website sampai data muncul.
   * *Ekstensi akan otomatis menangkap (intercept) request tersebut.*
   * Anda akan melihat notifikasi di console atau popup jika data berhasil ditangkap.
3. **Buka Popup Ekstensi**:
   * Klik ikon ekstensi di toolbar Chrome.
   * Jika status menunjukkan: **"SUDAH SIAP! Klik tombol..."**, berarti Anda siap lanjut.
   * *Jika masih "Tunggu Sebentar", coba refresh halaman web dan muat ulang tabelnya.*

### Pilihan Download

* **Fetch Auto (Excel)**:

  * Klik tombol ini untuk mengambil seluruh data dari semua halaman.
  * Tunggu proses berjalan (indikator jumlah item akan bertambah).
  * File `.xls` akan otomatis terdownload setelah selesai.
  * *Note: Jika muncul warning format file saat membuka di Excel, klik Yes.*
* **Download Semua PDF**:

  * Klik tombol ini untuk mendownload file PDF fisik dari setiap faktur yang terdaftar.
  * Proses akan berjalan otomatis satu per satu untuk menghindari blokir server.
