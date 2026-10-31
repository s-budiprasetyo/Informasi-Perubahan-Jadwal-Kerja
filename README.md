# Jadwal Kerja SK - Informasi Perubahan Jadwal Jam Kerja

## Menjalankan
```
pip install -r requirements.txt
streamlit run app.py
```
Data tersimpan di `data/jadwal.db` (SQLite). Satu pengguna, tanpa login.

## Departemen
Pilih di sidebar. Tiap departemen punya riwayat, saldo, dan file hasil sendiri.
- **SK (Shuttle Kiln)**: satu file berisi seksi KAMA (foreman Sulton/Fajar) dan QC SK (foreman Satrio) → sheet SATRIO, SULTON, PERUBAHAN REGU (R≥1 diberi akhiran `R1`).
- **QC (Quality Control)**, HRD "23. QC": R0 menghasilkan sheet `OFF` / `MASUK` (perubahan serentak ≥50% orang di satu tanggal, satu sheet per tanggal), `R0` (sisa tidak teratur), dan `PERUBAHAN REGU`. R1, R2, … menghasilkan sheet `R1`/`R2` (semua perubahan dari revisi sebelumnya) dan `PERUBAHAN REGU` bila ada.
  Susunan halaman (tim) dan batas serentak diatur di menu Pengaturan.
- Departemen baru: tambah satu entri di `core/depts.py`.

## Alur
1. Pilih seksi, bulan, lalu **Revisi** (otomatis dari riwayat: belum ada → R0, selain itu terakhir + 1).
2. **R0**: unggah ① Jadwal Kerja SK dan ② HRD "25. Shuttle Kiln 1". Hasil = selisih Jadwal Kerja terhadap pola HRD.
3. **R1, R2, …**: cukup unggah ① Jadwal Kerja SK terbaru. Hasil = selisih terhadap revisi sebelumnya (R0→R1, R1→R2), bukan terhadap HRD.
4. 🔔 berisi catatan (PWS tidak cocok pola, orang ada di HRD tapi tak ada di Jadwal, kode jam tak dikenal, dll.). Nilai bawaan sudah terpilih, keputusan tersimpan per bulan.
5. Tanggal efektif perubahan regu dapat diubah di pratinjau. Tombol **Simpan & buat Excel** menyimpan revisi, saldo hari, dan membuat file.
6. File keluaran hanya berisi form perubahan; R≥1 memberi akhiran nama sheet (`SULTON R1`). Tanggal header = tanggal dibuat.

## Saldo kurang/lebih hari
Kurang hari = target (kolom Total Masuk Kerja file HRD, cadangan AQ8 Jadwal) − hari kerja di Jadwal. Saldo akhir = saldo awal (bulan lalu) + kurang. Contoh: Okt kurang 4, Nov lebih 1 → sisa 3. Koreksi saldo awal manual ada di menu Saldo Hari.

## Backup Google Sheets (opsional)
1. Google Cloud → buat Service Account, unduh kunci JSON; aktifkan Google Sheets API.
2. Buat spreadsheet kosong, bagikan (Editor) ke email service account.
3. Salin `.streamlit/secrets.toml.example` ke `.streamlit/secrets.toml`, isi `backup_sheet_id` dan kunci.
4. Backup otomatis setiap Simpan. Bila database lokal kosong, aplikasi memulihkan otomatis dari Sheets.
Lembar per departemen: `saldo_hari_SK`/`_QC` dan `riwayat_SK`/`_QC` (terbaca manusia), `data_backup_SK`/`_QC` (untuk pemulihan).

## Tes
`python -m pytest tests` (tes memakai file contoh di folder attachments; sesuaikan path bila perlu).

## Belum ada
Unggahan SK26 (kolom ② pada R≥1) dan teks regu untuk PWS 4G01xx.

## Alur unduh & riwayat (penting)
1. Unggah file → atur keputusan (lonceng 🔔) → pilih **Hari nasional** (R0) → klik **Generate**. Generate **tidak menyimpan apa pun**.
2. Muncul dua tombol: *Unduh Jadwal Kerja* dan *Unduh Informasi Perubahan Jadwal*. Klik salah satu → jendela tinjau (isi sheet bisa dilihat).
3. Di jendela itu pilih **Simpan ke riwayat & unduh** atau **Unduh saja (tidak disimpan)**. File otomatis terunduh.
4. Bila revisi itu sudah ada di riwayat, jendela menampilkan peringatan (akan menimpa, dan revisi sesudahnya terhapus). Pilihan tetap bebas.
5. Hanya data yang Anda simpan yang menjadi dasar perbandingan revisi berikutnya dan saldo kurang hari bulan depan.
6. Riwayat → *Hapus revisi ini* (saldo dihitung ulang).

## Karyawan baru di tengah bulan & hari nasional
- Karyawan yang libur dari tgl 1 sampai tgl mulai: muncul di lonceng (R0) "mulai tgl X?". Bila *ya*, hari sebelum tgl mulai tidak dibuatkan perubahan dan tanggal perubahan regu = tgl mulai (bisa diubah).
- Hari nasional: pilih tanggalnya; libur di Jadwal pada tanggal itu tidak jadi perubahan, target hari kerja orang yang dasarnya masuk di tanggal itu berkurang 1.

## Saldo kurang / lebih hari & potong cuti
- Saldo dibaca dari **Jadwal Kerja yang diunggah** (nilai kolom di Excel): saldo awal (QC: LAMA / SK: KRNG. HARI), lebih hari (GANTI HARI), potong cuti (POT CUTI / POTONG CUTI), saldo akhir (QC: SISA = AN; SK: awal + AN - lebih - cuti).
- Saldo akhir = saldo awal + kurang bulan ini - lebih hari - potong cuti.
- Menu **Potong Cuti**: catat tanggal pengajuan, hari, keterangan. Riwayat memuat sisa kurang hari setelah tiap pengajuan; ikut ke Google Sheets (`potong_cuti_SK` / `potong_cuti_QC`).
- Bila saldo di Jadwal Kerja yang diunggah berbeda dengan riwayat aplikasi (saldo bulan lalu, potong cuti yang pernah dicatat, lebih hari revisi sebelumnya), muncul peringatan dan Anda **memilih**: data Jadwal Kerja unggahan (terbaru) atau data historis aplikasi. Generate baru aktif setelah memilih.
  - Memilih *Jadwal Kerja*: saat disimpan, selisih potong cuti dicatat sebagai "Penyesuaian dari Jadwal Kerja" dan saldo awal disamakan.
  - Memilih *historis*: file Jadwal Kerja hasil unduh diisi saldo awal / ganti hari / potong cuti dari aplikasi.
- Unduh ulang dari Riwayat selalu mengisi kolom-kolom itu dari saldo tersimpan.
