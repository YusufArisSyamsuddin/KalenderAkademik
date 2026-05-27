# SiWaka - Kalender Pendidikan

Aplikasi kalender pendidikan offline untuk menyusun kalender akademik, rekapitulasi kegiatan, dan analisis waktu efektif.

## Fitur Utama

- Dashboard kalender bulanan, semester, dan 1 tahun pelajaran.
- Format Akademik dengan tab `Model B`, `Model C`, dan `Semua Format`.
- Rekapitulasi terpisah antara `Kegiatan Tetap KBM` dan `Hari Tidak Efektif KBM`.
- Analisis Waktu Efektif berisi pekan efektif/tidak efektif, hari efektif, dan jam efektif.
- Ekspor laporan ke PDF A3/A4 dengan pengaturan margin, orientasi, dan skala; Word `.docx`; serta Excel `.xlsx`.
- Pengaturan identitas, hari kerja 5/6 hari, database, dan kategori kegiatan.

## Cara Membuka

1. Buka `kalender_pendidikan_2025_2026 (1).html` di browser.
2. Jika dijalankan dari GitHub Pages atau server lokal, aplikasi otomatis membaca `kalender_database.xlsx`.
3. Jika dibuka langsung dari file komputer, gunakan menu `Pengaturan > Buka Database` untuk memilih `kalender_database.xlsx`.

## Penyimpanan Data

- Data utama tersimpan di `kalender_database.xlsx`.
- Sheet database yang dipakai: `Settings`, `Events`, dan `Categories`.
- Kegiatan disimpan berdasarkan tanggal absolut, sehingga data lintas tahun pelajaran tetap saling terhubung.
- Setelah mengubah kegiatan atau pengaturan, buka menu `Pengaturan` lalu pilih `Simpan Database`.
- Browser Chrome atau Edge dapat menyimpan langsung ke file database. Browser lain akan mengunduh salinan database terbaru.
- Untuk pindah komputer atau hosting, cukup bawa satu folder ini beserta `kalender_database.xlsx`.
- Pilihan tampilan terakhir seperti menu aktif, periode, tab dashboard, tab format, dan filter disimpan sebagai preferensi browser agar tetap sama setelah reload.
- Agar tombol `Simpan Database` dapat memperbarui file Excel yang sama setelah reload, buka database sekali melalui Chrome/Edge dan berikan izin akses file saat diminta.

## Berkas Penting

- `kalender_pendidikan_2025_2026 (1).html`: aplikasi utama.
- `app.js`: logika kalender, database, perhitungan, dan ekspor.
- `styles.css`: tampilan aplikasi dan format cetak.
- `kalender_database.xlsx`: database Excel.
- `vendor/`: library lokal agar aplikasi tetap berjalan tanpa internet.
