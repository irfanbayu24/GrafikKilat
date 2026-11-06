"""Tes pencocokan fleksibel Unit+Nama:  python tests/test_matchlib.py"""
import os, sys, time, io
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import numpy as np, pandas as pd
import matchlib as ml

fails = 0
def check(name, cond, extra=""):
    global fails
    print(("  ✔ " if cond else "  ✘ GAGAL ") + name, "" if cond else extra)
    fails += (not cond)

F1 = pd.DataFrame({
    "Unit": ["BKT", "BKT", "Biro Keuangan", "Biro Keuangan", "Pusat Data dan Informasi", "Pusat Data dan Informasi", "Biro Umum", "Biro Umum", "BKT"],
    "Nama": ["Budi Santoso", "Citra Dewi Lestari", "Budi Hartono", "Dedi Kurniawan", "Eka Putri", "Fajar Nugroho, S.T., M.Si.", "Budi Santoso", "Gita", ""],
})
NOISE = ml.parse_noise(ml.DEFAULT_NOISE)

def mk(**kw):
    base = dict(noise=NOISE, unit_overlap=1.0, name_only_fallback=False, max_extra=0, typo=True)
    base.update(kw); return ml.Matcher(F1, "Unit", "Nama", **base)

def who(m, text):
    oc = m.match_text(text)
    return oc.status, [F1["Nama"][m.persons[p].row] + "|" + F1["Unit"][m.persons[p].row] for p in oc.persons], oc.confidence

print("== 1. Format 'Unit - Nama' dan variasinya")
m = mk()
check("nama sebagian: 'BKT - Budi' -> Budi Santoso (BKT), bukan Budi di unit lain", who(m, "BKT - Budi")[:2] == ("Cocok", ["Budi Santoso|BKT"]), who(m, "BKT - Budi"))
check("unit menentukan orang: 'Biro Keuangan - Budi' -> Budi Hartono", who(m, "Biro Keuangan - Budi")[1] == ["Budi Hartono|Biro Keuangan"])
check("inisial: 'Biro Umum - Budi S' -> Budi Santoso (Biro Umum)", who(m, "Biro Umum - Budi S")[1] == ["Budi Santoso|Biro Umum"])
check("urutan terbalik & huruf kecil: 'citra - bkt'", who(m, "citra - bkt")[1] == ["Citra Dewi Lestari|BKT"])
check("tanpa pemisah/format bebas: 'BKT_Citra_Dewi'", who(m, "BKT_Citra_Dewi")[1] == ["Citra Dewi Lestari|BKT"])
check("sapaan, gelar, perangkat diabaikan: 'Dr. Dedi (Biro Keuangan) iPhone'", who(m, "Dr. Dedi (Biro Keuangan) iPhone")[1] == ["Dedi Kurniawan|Biro Keuangan"])
check("gelar di File 1 diabaikan: 'Pusat Data dan Informasi - Fajar Nugroho'", who(m, "Pusat Data dan Informasi - Fajar Nugroho")[1] == ["Fajar Nugroho, S.T., M.Si.|Pusat Data dan Informasi"])
check("awalan nama: 'BKT - Cit' (≥3 huruf)", who(m, "BKT - Cit")[1] == ["Citra Dewi Lestari|BKT"])
check("salah ketik 1 huruf: 'BKT - Citra Dewi Lestary'", who(m, "BKT - Citra Dewi Lestary")[1] == ["Citra Dewi Lestari|BKT"])
check("huruf bertukar: 'BKT - Citra Dewi Letsari'", who(m, "BKT - Citra Dewi Letsari")[1] == ["Citra Dewi Lestari|BKT"])
check("aksen & tanda baca diabaikan: 'BKT – Bùdi,'", who(m, "BKT – Bùdi,")[1] == ["Budi Santoso|BKT"])
check("singkatan unit: 'PDI - Eka'", who(m, "PDI - Eka")[1] == ["Eka Putri|Pusat Data dan Informasi"])

print("== 2. Hal yang TIDAK boleh cocok (anti false-positive)")
check("tamu tak dikenal", who(m, "Tamu Luar")[0] == "Tidak cocok")
check("hanya unit tanpa nama", who(m, "BKT")[0] == "Tidak cocok")
check("hanya inisial tidak cukup: 'BKT - B'", who(m, "BKT - B")[0] == "Tidak cocok")
check("nama sama di unit yang salah: 'Biro Keuangan - Citra'", who(m, "Biro Keuangan - Citra")[0] == "Tidak cocok")
check("unit tak terbaca -> tidak cocok (mode wajib)", who(m, "Pusdatin - Eka Putri")[0] == "Tidak cocok")
check("kata asing ditolak pada mode ketat: 'BKT - Budi Raharjo'", who(m, "BKT - Budi Raharjo")[0] == "Tidak cocok")
check("teks kosong/NaN", who(m, "")[0] == "Kosong" and who(m, None)[0] == "Kosong" and who(m, float("nan"))[0] == "Kosong")

print("== 3. Ambigu tidak ditebak")
F1b = pd.concat([F1, pd.DataFrame({"Unit": ["BKT"], "Nama": ["Budi Hartanto"]})], ignore_index=True)
mb = ml.Matcher(F1b, "Unit", "Nama", noise=NOISE)
oc = mb.match_text("BKT - Budi")
check("dua Budi di BKT -> Ambigu (Rendah)", oc.status == "Ambigu" and len(oc.persons) == 2 and oc.confidence == "Rendah", oc)
oc = mb.match_text("BKT - Budi Santoso")
check("nama lengkap menyelesaikan ambigu -> Cocok Tinggi", oc.status == "Cocok" and oc.confidence == "Tinggi", oc)
F1c = pd.DataFrame({"Unit": ["BKT", "BKT"], "Nama": ["Budi", "Budi Santoso"]})
mc = ml.Matcher(F1c, "Unit", "Nama", noise=NOISE)
oc = mc.match_text("BKT - Budi")
check("'Budi' persis lebih diutamakan daripada 'Budi Santoso'", oc.status == "Cocok" and F1c["Nama"][mc.persons[oc.persons[0]].row] == "Budi", oc)

print("== 4. Opsi")
check("fallback nama saja: 'Pusdatin - Eka Putri' -> Cocok, keyakinan Rendah",
      who(mk(name_only_fallback=True), "Pusdatin - Eka Putri")[0::2] == ("Cocok", "Rendah"), who(mk(name_only_fallback=True), "Pusdatin - Eka Putri"))
check("fallback menuntut nama lengkap tercocokkan: 'Pusdatin - Eka' (sebagian) tetap ditolak",
      who(mk(name_only_fallback=True), "Pusdatin - Eka")[0] == "Tidak cocok" and who(mk(name_only_fallback=True), "Pusdatin Eka Putri")[0] == "Cocok")
check("fallback tidak dipakai jika unit terbaca (hasil tetap ketat)", who(mk(name_only_fallback=True), "BKT - Budi Raharjo")[0] == "Tidak cocok")
check("catatan menjelaskan unit tak terbaca", "unit tidak terbaca" in mk(name_only_fallback=True).match_text("Pusdatin - Eka Putri").note)
check("unit longgar (≥ separuh kata): 'Pusat Data - Fajar N'", who(mk(unit_overlap=0.5), "Pusat Data - Fajar N")[1] == ["Fajar Nugroho, S.T., M.Si.|Pusat Data dan Informasi"]
      and who(mk(), "Pusat Data - Fajar N")[0] == "Tidak cocok")
check("unit penuh diutamakan atas unit sebagian: 'Biro Umum - Budi Santoso'", who(mk(unit_overlap=0.5), "Biro Umum - Budi Santoso")[1] == ["Budi Santoso|Biro Umum"])
check("max_extra=1 menerima 1 kata asing", who(mk(max_extra=1), "BKT - Budi Raharjo")[1] == ["Budi Santoso|BKT"] and who(mk(max_extra=1), "BKT - Budi Raharjo")[2] == "Rendah")
check("typo dimatikan", who(mk(typo=False), "BKT - Citra Dewi Lestary")[0] == "Tidak cocok")
m0 = ml.Matcher(F1.drop(columns="Unit"), None, "Nama", noise=NOISE)
check("tanpa kolom unit (nama saja): 'Citra Dewi'", m0.match_text("Citra Dewi").status == "Cocok" and m0.match_text("Citra Dewi").confidence == "Sedang")
check("noise bisa diubah (mis. 'bkt' dibuang dari teks)", ml.tokenize("Pak Dr Budi iPhone", NOISE) == ["budi"] and ml.tokenize("Pak Budi", frozenset()) == ["pak", "budi"])
check("baris nama kosong di File 1 dicatat", mk().blank_rows == [8])

print("== 5. Seluruh data + laporan")
F2 = pd.DataFrame({"Display": ["BKT - Budi", "BKT - Budi Santoso", "Biro Keuangan - Budi", "citra - bkt", "Biro Umum - Budi S", "Tamu Luar",
                               "Pusdatin - Eka Putri", "Dr. Dedi (Biro Keuangan) iPhone", None, "BKT - Budi"], "Jam": ["08:00"] * 10})
mm = ml.Matcher(F1b, "Unit", "Nama", noise=NOISE)          # ada dua Budi di BKT
res = ml.run(mm, F1b, F2, "Display", "Unit", "Nama")
s = res.summary; print("   ", s)
st = dict(zip(res.people["Nama"] + "|" + res.people["Unit"], res.people["Status kehadiran"]))
check("'BKT - Budi Santoso' -> Hadir; Budi Hartanto -> Hadir (belum pasti) lewat 'BKT - Budi' ambigu",
      st["Budi Santoso|BKT"] == "Hadir" and st["Budi Hartanto|BKT"] == "Hadir (belum pasti)", st)
check("lainnya hadir/tidak sesuai", st["Budi Hartono|Biro Keuangan"] == "Hadir" and st["Citra Dewi Lestari|BKT"] == "Hadir"
      and st["Budi Santoso|Biro Umum"] == "Hadir" and st["Dedi Kurniawan|Biro Keuangan"] == "Hadir"
      and st["Eka Putri|Pusat Data dan Informasi"] == "Belum hadir" and st["Gita|Biro Umum"] == "Belum hadir"
      and st["|BKT"] == "Nama kosong", st)
check("ringkasan konsisten", s["hadir"] + s["belum_pasti"] + s["belum_hadir"] + s["blank_1"] == s["total_1"]
      and s["cocok"] + s["ambigu"] + s["tidak_cocok"] + s["kosong"] == s["total_2"], s)
check("jumlah entri cocok per orang (Budi Santoso BKT dicatat 1x lewat nama lengkap)",
      int(res.hadir.loc[res.hadir["Nama"].eq("Budi Santoso") & res.hadir["Unit"].eq("BKT"), "Jumlah entri cocok di File 2"].iloc[0]) == 1)
check("only_in_2 = baris asli File 2 (kolom Jam ikut)", list(res.only_in_2.columns) == ["Display", "Jam"] and set(res.only_in_2["Display"]) == {"Tamu Luar", "Pusdatin - Eka Putri"}, res.only_in_2)
check("rincian: ambigu menyebut kandidat", "Budi Santoso (BKT)" in res.detail.loc[res.detail["Status"] == "Ambigu", "Kandidat (jika ambigu)"].iloc[0])
import sheetlib as sl, openpyxl
wb = openpyxl.load_workbook(io.BytesIO(sl.to_excel_sheets(ml.build_report(res))))
check("laporan Excel", wb.sheetnames[:3] == ["Ringkasan", "Hadir", "Hadir (belum pasti)"] and "Nama kosong File 1" in wb.sheetnames, wb.sheetnames)
df_clash = pd.DataFrame({"Unit": ["A"], "Nama": ["Ani"], "Status kehadiran": ["x"]})
r2 = ml.run(ml.Matcher(df_clash, "Unit", "Nama", noise=NOISE), df_clash, pd.DataFrame({"t": ["A - Ani"]}), "t", "Unit", "Nama")
check("nama kolom hasil tidak menimpa kolom yang sudah ada", "Status kehadiran (2)" in r2.people.columns and r2.people["Status kehadiran (2)"].iloc[0] == "Hadir", list(r2.people.columns))

print("== 6. Kinerja (5.000 orang × 5.000 entri, 50 unit)")
rng = np.random.default_rng(7)
first = ["budi", "ani", "citra", "dedi", "eka", "fajar", "gita", "hadi", "indra", "joko", "muhammad", "siti", "sri", "agus", "dewi"]
last = ["santoso", "hartono", "lestari", "putra", "wibowo", "nugroho", "kurniawan", "pratama", "saputra", "rahayu", "utami", "wijaya"]
units = [f"Unit {chr(65 + i // 10)}{i % 10}" for i in range(50)]
rows = [(units[i % 50], f"{rng.choice(first).title()} {rng.choice(last).title()} {i}") for i in range(5000)]
big1 = pd.DataFrame(rows, columns=["Unit", "Nama"])
big2 = pd.DataFrame({"x": [f"{u} - {n.split()[0]} {n.split()[1]} {n.split()[2]}" for u, n in rows[:4500]] + [f"Tamu {i}" for i in range(500)]})
t = time.time(); mbig = ml.Matcher(big1, "Unit", "Nama", noise=NOISE); rbig = ml.run(mbig, big1, big2, "x", "Unit", "Nama"); el = time.time() - t
print(f"    {el:.1f} dtk · {rbig.summary}")
check("hasil benar pada data besar", rbig.summary["hadir"] == 4500 and rbig.summary["belum_hadir"] == 500 and rbig.summary["tidak_cocok"] == 500, rbig.summary)
check("selesai dalam waktu wajar (< 30 dtk)", el < 30, el)

print("\nHASIL:", "SEMUA PENGUJIAN LULUS" if fails == 0 else f"{fails} GAGAL")
sys.exit(1 if fails else 0)
