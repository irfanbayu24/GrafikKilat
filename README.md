# 📊 Visualisasi Spreadsheet  ·  🎓 AUTOCert

Aplikasi Streamlit dengan dua menu (sidebar kiri):

| Menu | Fungsi |
|---|---|
| 📊 **Visualisasi Spreadsheet** | Unggah spreadsheet → ringkasan, grafik interaktif, korelasi, tabel yang bisa difilter & diunduh |
| 🎓 **AUTOCert** | Template Word (.docx) + data peserta → sertifikat massal (PDF/DOCX) dalam satu ZIP |

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

## Struktur
| File | Isi |
|---|---|
| `app.py` | Pintu masuk + menu navigasi |
| `views/visualisasi.py`, `views/autocert.py` | Antarmuka tiap menu |
| `views/common.py` | Helper bersama (escape markdown, ingat berkas antar-menu) |
| `sheetlib.py` | Validasi/baca/analisis spreadsheet (tanpa Streamlit) |
| `certlib.py` | Validasi template, mail merge, nama file aman, konversi PDF (tanpa Streamlit) |
| `tests/` | `python tests/test_sheetlib.py` · `python tests/test_certlib.py` |

## Privasi
Berkas hanya diproses di memori server (cache kedaluwarsa ±30 menit) dan tidak disimpan permanen.
Data peserta adalah data pribadi — jika di-deploy publik, tambahkan login dan jalankan di container terbatas.
