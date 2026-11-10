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

## Aturan sheet QC (R0)
- Sheet **OFF** = semua perubahan X → OFF, sheet **MASUK** = semua OFF → X; tiap orang satu baris dengan tanggalnya sendiri (tanggal boleh berbeda-beda), urutan mengikuti tim. Sheet **R0** hanya untuk tukar kode jam (X → Y) dan hanya dibuat bila ada. Sudah dicocokkan dengan form buatan manual (tests/fixtures/manual_qc_oktober.xlsx).
- PERUBAHAN REGU: urutan mengikuti tim; bila lebih dari 23 orang dibagi dua halaman seimbang, nomor mulai 1 tiap halaman.
- Jendela unduh menutup otomatis setelah tombol diklik; backup Google Sheets berjalan di latar belakang.

## Aturan karyawan baru & lonceng (final)
- Orang yang **ada di HRD** = karyawan lama: semua tanggal libur/masuk dibandingkan dengan jadwal HRD, termasuk tanggal 1 sampai tanggal ia mulai masuk (cocok dengan form manual SK Oktober).
- Orang yang **tidak ada di HRD** (karyawan baru): libur sebelum tanggal masuk tidak dibuatkan perubahan (otomatis, tanpa pertanyaan).
- Lonceng hanya memuat: orang di HRD yang tidak ada di Jadwal Kerja (pilih masukkan/abaikan), PWS lama karyawan baru, kode PWS yang tidak cocok, dan **info selisih hari kerja vs target HRD** (kurang/lebih n hari). Tanggal efektif regu dihitung otomatis (tidak ada isian).

## Pindah departemen (QC ⇄ SK) dan format saldo baru

- **NIK** dinormalkan tanpa nol di depan di semua pembaca dan penyimpanan (`0698` = `698`). Baris NIK ganda di Jadwal Kerja dibuang (yang berisi hari kerja dipertahankan).
- **Memori karyawan lintas departemen** (per NIK) dibaca dari riwayat yang sudah disimpan: HRD departemen lain bulan ini, revisi terbaru bulan lalu, atau revisi terakhir bulan ini yang masih memuat namanya.
- Orang yang ada di Jadwal Kerja tetapi tidak ada di HRD departemen ini, namun dikenali di departemen lain, ditandai **pindahan**: PWS lama dari departemen asal, jadwal sebelum tanggal pindah = jadwal lama, perubahan regu **hanya** di form departemen baru (tanggal efektif = hari pertama memakai kode PWS baru). Tidak masuk lonceng; tercatat di info "Karyawan pindahan".
- R0: pindahan masih ada di departemen lama; R1/R2: hilang dari departemen lama (tidak dihitung dan tidak membuat peringatan) dan muncul di departemen baru. Di departemen lama, orang HRD yang sudah tercatat pindah tampil sebagai "Sudah pindah" (bukan lonceng).
- **HRD departemen lain (opsional)** dapat diunggah pada R0 (③); tanpa unggahan dipakai HRD tersimpan / memori bulan lalu. Pola PWS departemen lain dipakai untuk membandingkan PWS lama.
- **Saldo awal** pindahan = saldo akhir bulan lalu di departemen asal. Satu bulan penuh dihitung departemen baru.
- **Kolom saldo format baru (AT–AW, SK = QC)**: AQ GANTI HARI = hari kerja − target (rumus, tidak ditimpa), AT LAMA, AU BARU, AV TOTAL, AW SISA. Aplikasi mengisi LAMA dari riwayat per NIK bagi pindahan / baris yang LAMA-nya kosong atau error (VLOOKUP ke file bulan lalu).

### Aturan pola PWS & memori (perbaikan 9 Nov 2026)
- Pola PWS dari departemen lain **hanya dari bulan yang sama** (unggahan ③ didahulukan, lalu HRD tersimpan bulan itu). Pola bulan lain tidak dipakai karena kalendernya beda; pola 05D0nn yang hilang diturunkan dari HRD departemen sendiri.
- Orang yang ada di Jadwal tetapi belum ada di HRD: dicari di riwayat (revisi terakhir departemen lain / departemen sendiri); PWS lama = PWS terakhir di riwayat. Hanya bila tidak ada jejak sama sekali = karyawan baru (QC: PWS lama 05D002).
- Bulan di file HRD dibaca dari sel mana pun di 3 baris atas (A–K) dan divalidasi; kolom "Total" tidak lagi terbaca sebagai hari.
- 05D035 di QC: tetap dibuatkan perubahan regu dan perubahan jadwal bila jadwal berbeda dari HRD.

### QC: perubahan regu dari jam kerja, dan satu halaman OFF/MASUK (9 Nov 2026)
- Bila kolom PWS di Jadwal tidak sesuai jam kerjanya (≥40% hari kerja di luar pola), regu diturunkan otomatis dari kode jam (DSnn → 05D0nn): segmen minimal 3 hari kerja beruntun, tanggal efektif = hari kerja pertama segmen. Satu orang bisa punya lebih dari satu baris di PERUBAHAN REGU (mis. 05D035→05D007 tgl 1, lalu 05D007→05D002 tgl 9) dan tidak lagi masuk sheet R0. PWS di riwayat = PWS segmen terakhir.
- R0: bila total baris OFF + MASUK ≤ 19, keduanya digabung satu sheet "OFF & MASUK" (satu halaman, urutan tim, per orang OFF lalu MASUK); bila banyak tetap sheet OFF dan MASUK terpisah per tim.

### Karyawan harian lepas (PWS HRD 05D035 / FRE1 selamanya) - QC
- `daily_pws` di `core/depts.py`. Bila ada riwayat departemen (revisi terakhir bulan-bulan sebelumnya), PWS lama = PWS di riwayat (bukan 05D035), jadi regu 05D035→... tidak berulang tiap bulan; hanya perubahan jadwal harian (mis. DS02→OFF, OFF→DS08) dengan kode lama dari riwayat. Regu hanya bila PWS berubah dari riwayat atau jam berganti di tengah bulan (kasus Nova/Wardoyo). Tanpa riwayat: perilaku semula (regu dari 05D035).
- Harian lepas: perubahan jam yang berlaku seterusnya sampai akhir bulan (≥3 hari kerja beruntun, kolom PWS Jadwal belum diperbarui) dibuat sebagai regu dari PWS riwayat; tukar 1-2 hari atau blok yang kembali ke jadwal lama = perubahan jadwal.

### Sel rumus di file Excel (10 Nov 2026)
- Jadwal Kerja dan HRD dibaca dari **hasil rumus yang tersimpan** (bukan teks rumus / ArrayFormula). Kalau ada sel rumus tanpa hasil tersimpan, aplikasi berhenti dengan pesan jelas (buka di Excel, Ctrl+S, unggah lagi), bukan menghasilkan perubahan palsu.

### Hari pengganti (sel biru) - 10 Nov 2026
- Sel Jadwal Kerja berwarna biru muda = hari pengganti penutup kurang hari. Tetap dihitung sebagai hari masuk (perubahan jadwal OFF→kode, dan mengurangi saldo), tetapi tidak dipakai menentukan tanggal efektif perubahan regu.
- SK: bila kolom PWS di Jadwal sama dengan HRD tetapi jam kerjanya berganti seterusnya (mis. shift → DS02), regu diturunkan dari jam kerja (Burhanul 2G0302→05D002 tgl 9). Label PWS yang keliru (Erwin) tetap perilaku lama.

### Kolom PWS di Jadwal Kerja hasil (10 Nov 2026)
- Jadwal Kerja hasil generate (dan unduh ulang dari Riwayat) memperbaiki kolom **PWS** menjadi PWS terakhir hasil aplikasi (mis. Erwin 05D002→2G0301, Burhanul 2G0302→05D002). Hanya sel PWS yang berbeda yang diubah; kolom F (tersembunyi, PWS lama manual) dan sel lain tidak disentuh.
- Lonceng 🔔 menampilkan info "kolom PWS di Jadwal tidak sama dengan PWS terakhir" (tidak menahan Generate).

### Riwayat: unggah Jadwal Kerja untuk revisi lama
Revisi yang disimpan sebelum berkas Jadwal Kerja ikut tersimpan tidak punya tombol unduh Jadwal. Di menu Riwayat, unggah Jadwal Kerja revisi itu sekali; berkas disimpan di database dan tombol unduh muncul (dengan revisi, tanggal, saldo, dan PWS diperbarui).
