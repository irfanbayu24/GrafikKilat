import io, sys, zipfile
import os; sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import numpy as np, pandas as pd
import sheetlib as sl

def expect_error(name, data, label):
    try:
        sl.validate_upload(name, data)
        print("GAGAL (seharusnya ditolak):", label)
    except sl.SpreadsheetError as e:
        print("OK ditolak  :", label, "->", str(e)[:70])

# ---- 1. Validasi: file palsu harus ditolak
expect_error("virus.exe", b"MZ....", "ekstensi .exe")
expect_error("foto.xlsx", b"\x89PNG\r\n\x1a\n" + b"0"*100, "PNG diganti .xlsx")
expect_error("kosong.xlsx", b"", "file kosong")
buf = io.BytesIO()
with zipfile.ZipFile(buf, "w") as z: z.writestr("readme.txt", "halo")
expect_error("zip.xlsx", buf.getvalue(), "ZIP biasa diganti .xlsx")
expect_error("pdf.csv", b"%PDF-1.4\n...", "PDF diganti .csv")
expect_error("biner.csv", b"abc\x00\x01\x02def", "biner diganti .csv")
expect_error("lama.xls", b"bukan xls", ".xls palsu")
expect_error("besar.csv", b"a,b\n" + b"1,2\n" * (sl.MAX_FILE_MB * 1024 * 1024 // 4 + 10), "terlalu besar")
bomb = io.BytesIO()
with zipfile.ZipFile(bomb, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("xl/workbook.xml", "<x/>")
    z.writestr("xl/big.bin", b"0" * (sl.MAX_UNZIPPED_MB * 1024 * 1024 + 1024))
print("   (zip-bomb size:", len(bomb.getvalue()), "byte terkompresi)")
expect_error("bomb.xlsx", bomb.getvalue(), "zip-bomb")

# ---- 2. File sampel (xlsx asli) harus lolos
data = sl.make_sample_xlsx()
kind = sl.validate_upload("data_contoh.xlsx", data)
sheets, notes = sl.load_workbook(kind, data)
print("\nSheet:", list(sheets), "| catatan:", notes)
raw = sheets["Penjualan"]
h = sl.guess_header_row(raw)
print("Header terdeteksi di baris ke-", h + 1, "(harapan 3)")
assert h == 2, h
df, removed = sl.build_dataframe(raw, h)
roles = sl.column_roles(df)
print(df.dtypes.to_string())
print(roles)
assert roles["Tanggal"] == "tanggal" and roles["Total Penjualan"] == "numerik"
assert roles["Wilayah"] == "kategori" and len(df) == 400

# ---- 3. CSV Indonesia (titik koma, format angka & tanggal lokal, total di bawah)
csv_text = ("Laporan Bulanan\n\nTanggal;Cabang;Omzet;Margin;Kode\n"
            "15/01/2025;A;Rp 1.250.000,50;85%;00123\n"
            "20 Februari 2025;B;Rp 980.000;70%;00456\n"
            "03-03-2025;A;2.100.000;90%;00789\n"
            "Total;;4.330.000,50;;\n").encode("utf-8")
k = sl.validate_upload("lokal.csv", csv_text)
sh, _ = sl.load_workbook(k, csv_text)
raw2 = sh["Data"]
h2 = sl.guess_header_row(raw2)
d2, rem = sl.build_dataframe(raw2, h2)
print("\nCSV header baris", h2 + 1, "| baris total dibuang:", rem)
print(d2.dtypes.to_string()); print(d2.to_string())
assert h2 == 2 and rem == 1 and len(d2) == 3
assert pd.api.types.is_datetime64_any_dtype(d2["Tanggal"])
assert d2["Tanggal"].dt.month.tolist() == [1, 2, 3], d2["Tanggal"].tolist()
assert abs(d2["Omzet"].iloc[0] - 1250000.5) < 1e-6 and d2["Omzet"].iloc[1] == 980000 and d2["Omzet"].iloc[2] == 2100000
assert d2["Margin"].tolist() == [85, 70, 90]
assert d2["Kode"].iloc[0] == "00123"   # kode berawalan 0 tetap teks

# ---- 4. Header ganda/kosong, duplikat, kolom/baris kosong
raw3 = pd.DataFrame([["Nama", "Nama", None, "Skor"], ["Ani", "x", 1, 10], [None]*4, ["Budi", "y", 2, 20], ["Ani", "x", 1, 10]], dtype=object)
d3, _ = sl.build_dataframe(raw3, 0, drop_duplicates=True)
print("\nKolom:", list(d3.columns), "| baris:", len(d3))
assert list(d3.columns) == ["Nama", "Nama (2)", "Kolom 3", "Skor"] and len(d3) == 2

# ---- 5. Agregasi, periode, korelasi, kualitas
agg = sl.aggregate(df, ["Wilayah"], "Total Penjualan", "sum").sort_values("Nilai", ascending=False)
assert agg["Nilai"].sum() == df["Total Penjualan"].sum()
cnt = sl.aggregate(df, ["Wilayah", "Kategori"], None, "size")
assert cnt["Nilai"].sum() == 400
mo = df.assign(Tanggal=sl.bucket_dates(df["Tanggal"], "Bulanan"))
assert sl.aggregate(mo, ["Tanggal"], "Jumlah Terjual", "sum")["Tanggal"].nunique() <= 12
print("Granularitas otomatis:", sl.auto_granularity(df["Tanggal"]))
nums = [c for c, r in roles.items() if r == "numerik"]
print(sl.top_correlations(df[nums].corr()).to_string(index=False))
print(sl.quality_report(df, roles).head(3).to_string(index=False))
print(sl.describe_numeric(df, nums[:2]).to_string(index=False))

# ---- 6. Ekspor aman dari injeksi rumus
evil = pd.DataFrame({"a": ["=HYPERLINK(\"http://x\")", "@cmd", "aman"], "b": [1, 2, 3]})
assert b"'=HYPERLINK" in sl.to_csv_bytes(evil)
import openpyxl
wb = openpyxl.load_workbook(io.BytesIO(sl.to_excel_bytes(evil)))
c = wb["Data"]["A2"]
print("\nCell xlsx:", c.data_type, repr(c.value)[:30]); assert c.data_type != "f"
rt, _ = sl.load_workbook("xlsx", sl.to_excel_bytes(df))
assert len(rt["Data"]) == 401   # header + 400 baris (ekspor bisa dibaca ulang)

# ---- 7. Edge case
e, _ = sl.build_dataframe(pd.DataFrame(), 0); assert e.empty
one, _ = sl.build_dataframe(pd.DataFrame([["Hanya header"]], dtype=object), 0); assert one.empty

# ---- 8. Gaya angka per kolom
def conv(vals):
    r = pd.DataFrame([["x"]] + [[v] for v in vals], dtype=object)
    return sl.build_dataframe(r, 0)[0]["x"].tolist()
assert conv(["1.234", "2.5", "3.75"]) == [1.234, 2.5, 3.75]            # desimal gaya Inggris
assert conv(["1.234", "2.500", "1.000.000"]) == [1234, 2500, 1000000]   # ribuan gaya Indonesia
assert conv(["1,234", "2,500", "1,000,000"]) == [1234, 2500, 1000000]  # ribuan gaya Inggris
assert conv(["12,5", "3,25", "4"]) == [12.5, 3.25, 4]                   # desimal koma
assert conv(["+62812", "+62813"]) == ["+62812", "+62813"]               # telepon tetap teks
assert conv([True, False, True]) == ["Ya", "Tidak", "Ya"]
print("\nSEMUA PENGUJIAN LULUS")
