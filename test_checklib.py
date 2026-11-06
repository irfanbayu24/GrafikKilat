"""Tes logika Cek Absensi:  python tests/test_checklib.py"""
import io, os, sys, time
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import numpy as np, pandas as pd
import checklib as ck
import sheetlib as sl

fails = 0
def check(name, cond, extra=""):
    global fails
    print(("  ✔ " if cond else "  ✘ GAGAL ") + name, "" if cond else extra)
    fails += (not cond)

def cmp(k1, k2, opts=ck.Options(), **kw):
    return ck.compare(pd.DataFrame({"K": k1, "v": range(len(k1))}), "K",
                      pd.DataFrame({"K": k2, "w": range(len(k2))}), "K", opts, **kw)

print("== 1. Regresi bug skrip lama")
r = cmp(["budi.santoso@x.com", "budi.rahmat@x.com", "ani@x.com"], ["budi.santoso@x.com"])
check("email berisi titik tidak lagi terpotong (dulu 'budi' cocok ke semua)", r.found["K"].tolist() == ["budi.santoso@x.com"]
      and r.missing["K"].tolist() == ["budi.rahmat@x.com", "ani@x.com"], (r.found["K"].tolist(), r.missing["K"].tolist()))
r = cmp(["A1", None, "", "B2", "  "], ["A1", None, ""])
check("kunci kosong tidak pernah dianggap cocok (dulu 'sudah ada' bila File 2 punya sel kosong)",
      r.found["K"].tolist() == ["A1"] and r.missing["K"].tolist() == ["B2"] and len(r.blank_1) == 3, r.summary)
check("kunci kosong dilaporkan terpisah", r.summary["blank_1"] == 3 and r.summary["blank_2"] == 2)
r = cmp([12345.0, 67890.0, 111.0], ["12345", 67890, "999"])
check("12345.0 = '12345' = 12345 (angka vs teks)", r.found["K"].tolist() == [12345.0, 67890.0], r.found["K"].tolist())
df = pd.DataFrame([["NIP"], ["198001012005011001"], ["199002022010012002"]], dtype=object)
a, _ = sl.build_dataframe(df, 0); b, _ = sl.build_dataframe(pd.DataFrame([["NIP"], ["198001012005011001"], ["199999999999999999"]], dtype=object), 0)
r = ck.compare(a, "NIP", b, "NIP")
check("NIP 18 digit (teks) dicocokkan persis, tanpa kehilangan digit", r.found["NIP"].tolist() == ["198001012005011001"]
      and r.missing["NIP"].tolist() == ["199002022010012002"], (r.found["NIP"].tolist(), r.missing["NIP"].tolist()))

print("== 2. Aturan pencocokan")
check("huruf besar-kecil diabaikan (default)", len(cmp(["BUDI SANTOSO"], ["budi santoso"]).found) == 1)
check("huruf besar-kecil dibedakan bila dimatikan", len(cmp(["BUDI"], ["budi"], ck.Options(ignore_case=False)).found) == 0)
check("spasi & NBSP diabaikan", len(cmp(["1980 0101\u00a02005"], ["198001012005"]).found) == 1)
check("spasi ganda diratakan bila 'abaikan spasi' mati", len(cmp(["Budi   Santoso "], ["Budi Santoso"], ck.Options(ignore_spaces=False)).found) == 1)
check("hanya angka (NIP berpemisah)", len(cmp(["19800101 200501 1 001"], ["198001012005011001"], ck.Options(digits_only=True)).found) == 1)
check("nol di depan diabaikan bila dinyalakan", len(cmp(["00123"], ["123"], ck.Options(strip_zeros=True)).found) == 1
      and len(cmp(["00123"], ["123"]).found) == 0)
check("'0' tetap kunci sah saat strip nol", len(cmp(["0", "00"], ["0"], ck.Options(strip_zeros=True)).found) == 2)
check("karakter lebar-nol (zero-width) dibuang", len(cmp(["abc\u200b"], ["abc"]).found) == 1)

print("== 3. Hasil, jumlah kemunculan, kembar, hanya di File 2")
r = cmp(["A", "B", "C", "A"], ["A", "A", "A", "C", "Z"])
check("sudah ada + jumlah kemunculan di File 2", r.found["K"].tolist() == ["A", "C", "A"] and r.found[ck.COUNT_COL].tolist() == [3, 1, 3], r.found)
check("belum ada", r.missing["K"].tolist() == ["B"])
check("hanya di File 2", r.only_in_2["K"].tolist() == ["Z"])
check("kunci kembar File 1 & 2 dilaporkan", r.dup_1["Kunci"].tolist() == ["A"] and r.dup_1["Jumlah baris"].tolist() == [2]
      and r.dup_2["Kunci"].tolist() == ["A"] and r.dup_2["Jumlah baris"].tolist() == [3])
check("ringkasan konsisten", r.summary["found"] + r.summary["missing"] + r.summary["blank_1"] == r.summary["total_1"]
      and r.summary["found_unique"] == 2 and r.summary["pct_found"] == 75.0, r.summary)
r = ck.compare(pd.DataFrame({"K": ["A", "B"], ck.COUNT_COL: [1, 2]}), "K", pd.DataFrame({"K": ["A"]}), "K")
check("nama kolom 'Jumlah di File 2' tidak menimpa kolom yang sudah ada", list(r.found.columns) == ["K", ck.COUNT_COL, ck.COUNT_COL + " (2)"], list(r.found.columns))
d2 = pd.DataFrame({"K": ["A", "A", "C"], "Jam": ["08:00", "09:00", "08:30"], "Lokasi": ["X", "Y", "Z"]})
r = ck.compare(pd.DataFrame({"K": ["A", "B", "C"]}), "K", d2, "K", extra_cols=("Jam", "Lokasi"))
check("kolom tambahan dari File 2 (kemunculan pertama)", r.found["File 2: Jam"].tolist() == ["08:00", "08:30"] and r.found["File 2: Lokasi"].tolist() == ["X", "Z"], r.found)
r = ck.compare(pd.DataFrame({"Nama": ["ani", "budi"]}), "Nama", pd.DataFrame({"Hadir": ["Budi"]}), "Hadir")
check("nama kolom kunci boleh berbeda antar-file", r.found["Nama"].tolist() == ["budi"])
try:
    ck.compare(pd.DataFrame({"a": [1]}), "x", pd.DataFrame({"a": [1]}), "a"); check("kolom kunci salah ditolak", False)
except ck.CheckError:
    check("kolom kunci salah ditolak dengan pesan ramah", True)
e = ck.compare(pd.DataFrame({"K": []}), "K", pd.DataFrame({"K": []}), "K")
check("file kosong tidak error", e.summary["found"] == 0 and e.summary["pct_found"] == 0.0)

print("== 4. Tebakan kolom kunci & peringatan presisi")
check("utamakan kolom bernama sama (NIP)", ck.guess_key_columns(["No", "Nama", "NIP"], ["NIP", "Jam"]) == ("NIP", "NIP"))
check("beda nama tapi sama-sama bermakna", ck.guess_key_columns(["Nama Pegawai", "NIP"], ["Nama Hadir", "Email"]) == ("Nama Pegawai", "Nama Hadir"))
check("cadangan: kolom pertama", ck.guess_key_columns(["x", "y"], ["p", "q"]) == ("x", "p"))
check("angka ≥ 16 digit terdeteksi rusak", ck.precision_warning_count(pd.Series([3.2010101010100001e15, 123.0, "x"])) == 1
      and ck.precision_warning_count(pd.Series([123456789012345.0])) == 0)
check("bilangan bulat ≥ 16 digit (int) juga terdeteksi; teks 18 digit aman",
      ck.precision_warning_count(pd.Series([3201010101010001, 5])) == 1
      and ck.precision_warning_count(pd.Series(["198001012005011001"])) == 0
      and ck.precision_warning_count(pd.Series([True, False])) == 0 and ck.precision_warning_count(pd.Series([], dtype=object)) == 0)

print("== 5. Laporan Excel")
r = cmp(["A", "B", "", "=CMD()"], ["A", "Z"])
sheets = ck.build_report(r)
check("sheet laporan", list(sheets) == ["Ringkasan", "Sudah ada", "Belum ada", "Hanya di File 2", "Kunci kosong File 1"], list(sheets))
import openpyxl
wb = openpyxl.load_workbook(io.BytesIO(sl.to_excel_sheets(sheets)))
check("Excel terbuka & rumus berbahaya dinetralkan", wb.sheetnames == list(sheets) and all(c.data_type != "f" for row in wb["Belum ada"].iter_rows() for c in row))

print("== 6. Kinerja data besar (40.000+ baris)")
rng = np.random.default_rng(1)
ids = [f"{198000000000000000 + i * 7919}" for i in range(60_000)]
big1 = pd.DataFrame({"NIP": ids, "Nama": [f"Pegawai {i}" for i in range(60_000)]})
pick = rng.choice(60_000, 45_000, replace=False)
big2 = pd.DataFrame({"NIP": [ids[i] for i in pick] + ["999"] * 5, "Jam": "08:00"})
t = time.time(); r = ck.compare(big1, "NIP", big2, "NIP", extra_cols=("Jam",)); dt_ = time.time() - t
print(f"    60.000 vs 45.005 baris: {dt_:.2f} dtk")
check("hasil benar pada data besar", r.summary["found"] == 45_000 and r.summary["missing"] == 15_000 and r.summary["only_2"] == 5, r.summary)
check("selesai dalam waktu wajar (< 10 dtk)", dt_ < 10, dt_)

print("\nHASIL:", "SEMUA PENGUJIAN LULUS" if fails == 0 else f"{fails} GAGAL")
sys.exit(1 if fails else 0)
