# 📊 Visualisasi Spreadsheet  ·  🎓 AUTOCert  ·  ✅ Cek Absensi

Aplikasi Streamlit dengan tiga menu (sidebar kiri):

| Menu | Fungsi |
|---|---|
| 📊 **Visualisasi Spreadsheet** | Unggah spreadsheet → ringkasan, grafik interaktif, korelasi, tabel yang bisa difilter & diunduh |
| 🎓 **AUTOCert** | Template Word (.docx) + data peserta → sertifikat massal (PDF/DOCX) dalam satu ZIP |
| ✅ **Cek Absensi** | Bandingkan dua daftar lewat kolom kunci: sudah ada / belum ada / hanya di File 2 + laporan Excel |

## Menjalankan
```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## AUTOCert
1. Tulis tag di template Word: `«Nama»`, `{{Nama}}`, `<<Nama>>`, atau `[Nama]` (sama dengan judul kolom Excel;
   huruf besar-kecil dan spasi/garis bawah tidak dibedakan). Tag boleh di paragraf, tabel, text box, header, footer.
2. Unggah template + data peserta → periksa kecocokan → **Coba 1 sertifikat** → **Generate semua** → unduh ZIP.

**Mesin PDF** (dipilih otomatis; tanpa keduanya hanya DOCX):
- **Microsoft Word** (Windows, butuh `comtypes`) — paling akurat. Hanya untuk komputer pribadi.
- **LibreOffice** (Windows/Linux/Mac) — perlu terpasang (`soffice` di PATH). Di Linux/Docker: paket
  `libreoffice-writer` + font yang dipakai template, atau tata letak PDF bisa berbeda.

**Keamanan template:** hanya `.docx` biasa; ditolak jika memuat makro, field `INCLUDETEXT/INCLUDEPICTURE/LINK/DDE/IMPORT`,
atau relasi eksternal selain hyperlink. Maks. 10 MB, 1.000 peserta per proses. Berkas sementara memakai folder unik per
proses (aman untuk banyak pengguna) dan tidak ada `taskkill`.

## Cek Absensi
- **File 1** = daftar acuan (mis. semua pegawai), **File 2** = daftar pembanding (mis. yang hadir). Pilih kolom kunci
  (NIP, NIK, ID, email, nama); nama kolomnya boleh berbeda antar-file.
- Hasil: **sudah ada** (+ jumlah kemunculan), **belum ada**, **hanya di File 2**, serta **perlu dicek** (kunci kosong/kembar).
  Laporan lengkap bisa diunduh dalam satu file Excel.
- Aturan pencocokan bisa diatur: abaikan huruf besar/kecil, spasi, hanya angka, nol di depan. Kunci kosong tidak pernah dianggap cocok.
- **Penting:** NIK (16 digit) / NIP (18 digit) harus berformat *Teks* di Excel. Jika berformat Angka, Excel hanya menyimpan
  15 digit sehingga digit belakangnya berubah menjadi nol; aplikasi memberi peringatan jika mendeteksinya.
- Cepat untuk data besar (60.000 baris dibandingkan dalam < 1 detik).
- **Mode fleksibel (Unit + Nama → satu kolom teks):** File 1 punya kolom *Unit* dan *Nama* terpisah, File 2 hanya satu kolom
  seperti `BKT - Budi`. Dibaca per kata: urutan `Unit - Nama` atau `Nama - Unit` sama saja; nama boleh sebagian (`Budi`),
  inisial (`Budi S`), awalan, atau salah ketik 1 huruf; sapaan/gelar/label perangkat (`Pak`, `Dr.`, `iPhone`) diabaikan; unit
  dikenali dari kata atau singkatannya. Nama yang cocok dengan lebih dari satu orang **tidak ditebak** — ditandai
  *Hadir (belum pasti)*. Setiap kecocokan diberi keyakinan Tinggi/Sedang/Rendah, ada tab *Rincian* dan penguji satu teks.

## Struktur
| File | Isi |
|---|---|
| `app.py` | Pintu masuk + menu navigasi |
| `views/visualisasi.py`, `views/autocert.py`, `views/cek_absensi.py` | Antarmuka tiap menu |
| `views/common.py` | Helper bersama (escape markdown, ingat berkas antar-menu) |
| `sheetlib.py` | Validasi/baca/analisis spreadsheet (tanpa Streamlit) |
| `certlib.py` | Validasi template, mail merge, nama file aman, konversi PDF (tanpa Streamlit) |
| `checklib.py` | Normalisasi kunci, perbandingan dua daftar, laporan (tanpa Streamlit) |
| `matchlib.py` | Pencocokan fleksibel Unit+Nama ↔ satu kolom teks (tanpa Streamlit) |
| `tests/` | `python tests/test_sheetlib.py` · `test_certlib.py` · `test_checklib.py` · `test_matchlib.py` |

## Privasi
Berkas hanya diproses di memori server (cache kedaluwarsa ±30 menit) dan tidak disimpan permanen.
Data peserta adalah data pribadi — jika di-deploy publik, tambahkan login dan jalankan di container terbatas.
