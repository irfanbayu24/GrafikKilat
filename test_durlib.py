"""Tes durasi kehadiran (join/leave):  python tests/test_durlib.py"""
import io, os, sys
from datetime import time, datetime
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))
import numpy as np, pandas as pd
import durlib as dl, checklib as ck, matchlib as ml, sheetlib as sl

fails = 0
def check(name, cond, extra=""):
    global fails
    print(("  ✔ " if cond else "  ✘ GAGAL ") + name, "" if cond else extra)
    fails += (not cond)

print("== 1. Membaca tanggal-jam")
z, notes = dl.parse_times(pd.Series(["01/15/2026 08:01:23 AM", "01/15/2026 01:05:00 PM", "", None]))
check("format Zoom AM/PM (Bulan/Hari/Tahun)", z.iloc[0] == pd.Timestamp("2026-01-15 08:01:23") and z.iloc[1] == pd.Timestamp("2026-01-15 13:05:00") and z.iloc[2:].isna().all(), z.tolist())
check("kosong tidak dianggap galat", notes == [], notes)
z, notes = dl.parse_times(pd.Series(["15/01/2026 08:01", "20/02/2026 13:30"]))
check("Hari/Bulan/Tahun terdeteksi otomatis (hari 15 > 12)", z.iloc[0] == pd.Timestamp("2026-01-15 08:01") and notes == [], (z.tolist(), notes))
z, notes = dl.parse_times(pd.Series(["01/15/2026 08:00", "02/20/2026 09:00"]))
check("Bulan/Hari/Tahun terdeteksi otomatis (tanggal 15 > 12)", z.iloc[0] == pd.Timestamp("2026-01-15 08:00") and notes == [])
z, notes = dl.parse_times(pd.Series(["03/04/2026 08:00"]))
check("ambigu: diasumsikan Bulan/Hari + diberi peringatan", z.iloc[0] == pd.Timestamp("2026-03-04 08:00") and any("ambigu" in n for n in notes), notes)
z, _ = dl.parse_times(pd.Series(["03/04/2026 08:00"]), "dmy")
check("urutan dipaksa Hari/Bulan", z.iloc[0] == pd.Timestamp("2026-04-03 08:00"))
z, nt = dl.parse_times(pd.Series(["01/15/2026 08:00"]), "dmy")
check("urutan dipaksa tetapi nilai mustahil (bulan 15) dibaca lunak sebagai 15 Januari", z.iloc[0] == pd.Timestamp("2026-01-15 08:00") and nt == [], (z.tolist(), nt))
z, _ = dl.parse_times(pd.Series(["Jan 15, 2026 08:01:23 AM", "2026-01-15 09:00:00", "15 Januari 2026 10:00"]))
check("nama bulan & ISO", z.iloc[0] == pd.Timestamp("2026-01-15 08:01:23") and z.iloc[1] == pd.Timestamp("2026-01-15 09:00") , z.tolist())
z, _ = dl.parse_times(pd.Series([datetime(2026, 1, 15, 8, 0), pd.Timestamp("2026-01-15 09:00"), time(8, 30, 15), "08:45", "01:10 PM"]))
check("objek datetime/time Excel & jam saja", z.iloc[0] == pd.Timestamp("2026-01-15 08:00") and z.iloc[2].time() == time(8, 30, 15)
      and z.iloc[3].time() == time(8, 45) and z.iloc[4].time() == time(13, 10), z.tolist())
z, notes = dl.parse_times(pd.Series(["01/15/2026 08:00 AM", "bukan waktu", "25:99"]))
check("nilai tak terbaca dilaporkan & jadi NaT", z.isna().tolist() == [False, True, True] and any("2 nilai" in n for n in notes), notes)

print("== 2. Mengenali tabel Zoom")
raw = pd.DataFrame([["Topic", "ID", "Host", "Duration (Minutes)", "Start Time", "End Time", "Participants"],
                    ["Rapat", "123", "a@b.c", 90, "01/15/2026 08:00:00 AM", "01/15/2026 09:30:00 AM", 3],
                    [None] * 7,
                    ["Name (Original Name)", "User Email", "Join Time", "Leave Time", "Duration (Minutes)", "Guest", None],
                    ["BKT - Budi", "b@x.id", "01/15/2026 08:00:00 AM", "01/15/2026 08:20:00 AM", 20, "No", None]], dtype=object)
check("header tabel peserta ditemukan di baris ke-4 (bukan blok ringkasan rapat di atasnya)", dl.find_session_header(raw) == 3, dl.find_session_header(raw))
check("tebakan header umum (sheetlib) juga memilih baris yang sama", sl.guess_header_row(raw) == 3, sl.guess_header_row(raw))
check("kolom join/leave/durasi ditebak", dl.guess_session_columns(["Name (Original Name)", "User Email", "Join Time", "Leave Time", "Duration (Minutes)", "Guest"]) == ("Join Time", "Leave Time", "Duration (Minutes)"))
check("versi Indonesia", dl.guess_session_columns(["Nama", "Waktu Bergabung", "Waktu Keluar", "Durasi (menit)"]) == ("Waktu Bergabung", "Waktu Keluar", "Durasi (menit)"))
check("tanpa kolom waktu -> None", dl.find_session_header(pd.DataFrame([["a", "b", "c"]], dtype=object)) is None)

print("== 3. Aturan durasi")
def T(h, m=0): return pd.Timestamp(f"2026-01-15 {h:02d}:{m:02d}:00")
def run(sessions, rule=dl.Rule(), who=None):
    """sessions: [(orang, join, leave)] -> tabel durasi per orang (0..n-1)"""
    people = sorted({s[0] for s in sessions}); idx = {p: i for i, p in enumerate(people)}
    tbl, info = dl.attendance_table(len(people), [[idx[s[0]]] for s in sessions],
                                    pd.Series([s[1] for s in sessions]), pd.Series([s[2] for s in sessions]), rule)
    tbl.index = people; return tbl, info
tb, _ = run([("budi", T(8), T(8, 20)), ("budi", T(8, 25), T(9, 10))])
check("reconnect (jeda 5 mnt ≤ 10): 20 + 45 + jeda 5 = 70, Memenuhi", tb.loc["budi", "Total durasi (menit)"] == 70.0 and tb.loc["budi", "Status durasi"] == dl.STATUS_OK
      and tb.loc["budi", "Reconnect (jeda kecil)"] == 1 and tb.loc["budi", "Jumlah sesi"] == 2, tb.loc["budi"].to_dict())
tb, _ = run([("budi", T(8), T(8, 20)), ("budi", T(8, 25), T(9, 10))], dl.Rule(count_gaps=False))
check("jeda reconnect tidak dihitung: 65", tb.loc["budi", "Total durasi (menit)"] == 65.0)
tb, _ = run([("dedi", T(8), T(8, 30)), ("dedi", T(8, 45), T(9, 20))])
check("jeda 15 mnt (> 10) bukan reconnect & tidak dihitung: 65, jeda panjang 1", tb.loc["dedi", "Total durasi (menit)"] == 65.0 and tb.loc["dedi", "Jeda panjang"] == 1 and tb.loc["dedi", "Reconnect (jeda kecil)"] == 0)
tb, _ = run([("dedi", T(8), T(8, 30)), ("dedi", T(8, 45), T(9, 20))], dl.Rule(method="span"))
check("metode rentang (join pertama–leave terakhir) = 80", tb.loc["dedi", "Total durasi (menit)"] == 80.0)
tb, _ = run([("dedi", T(8), T(8, 30)), ("dedi", T(8, 45), T(9, 5))], dl.Rule(tolerance_min=20))
check("toleransi dinaikkan ke 20: jeda 15 jadi reconnect (30+20+15=65)", tb.loc["dedi", "Total durasi (menit)"] == 65.0 and tb.loc["dedi", "Reconnect (jeda kecil)"] == 1)
tb, _ = run([("eka", T(8), T(9)), ("eka", T(8), T(9))])
check("dua perangkat bersamaan tidak dihitung dobel: 60 (bukan 120)", tb.loc["eka", "Total durasi (menit)"] == 60.0 and tb.loc["eka", "Jumlah sesi"] == 2)
tb, _ = run([("eka", T(8), T(8, 40)), ("eka", T(8, 20), T(9))])
check("sesi tumpang tindih sebagian digabung: 08:00–09:00 = 60", tb.loc["eka", "Total durasi (menit)"] == 60.0)
tb, _ = run([("eka", T(8), T(9))]); tb2, _ = run([("eka", T(8), T(9))], dl.Rule(inclusive=False))
check("tepat 60 menit: Memenuhi bila '≥', Kurang bila '>'", tb.loc["eka", "Status durasi"] == dl.STATUS_OK and tb2.loc["eka", "Status durasi"] == dl.STATUS_SHORT)
tb, _ = run([("citra", T(8), T(8, 59, ) if False else pd.Timestamp("2026-01-15 08:59:59"))])
check("59 mnt 59 dtk belum memenuhi", tb.loc["citra", "Status durasi"] == dl.STATUS_SHORT and tb.loc["citra", "Total durasi (menit)"] == 60.0 or tb.loc["citra", "Status durasi"] == dl.STATUS_SHORT, tb.loc["citra"].to_dict())
tb, _ = run([("fajar", T(7, 50), T(10, 30))], dl.Rule(win_start=time(8, 0), win_end=time(9, 0)))
check("dibatasi jam acara 08:00–09:00: ruang tunggu & setelah acara tidak dihitung = 60", tb.loc["fajar", "Total durasi (menit)"] == 60.0 and tb.loc["fajar", "Join pertama"] == T(8))
tb, info = run([("gita", T(6), T(7))], dl.Rule(win_start=time(8, 0), win_end=time(9, 0)))
check("sesi di luar jam acara diabaikan & dicatat", tb.loc["gita", "Status durasi"] == "" and info["outside_window"] == 1, info)
tb, info = run([("hadi", T(9), T(8)), ("hadi", pd.NaT, T(9)), ("hadi", T(8), T(8, 10))])
check("sesi rusak (leave < join / waktu kosong) dilewati & dicatat", tb.loc["hadi", "Total durasi (menit)"] == 10.0 and info["negative"] == 1 and info["unreadable"] == 1, info)
tb, _ = run([("indra", T(8), T(8, 10)), ("indra", T(8, 12), T(8, 20))], dl.Rule(threshold_min=15))
check("batas bisa diubah (15 menit)", tb.loc["indra", "Status durasi"] == dl.STATUS_OK)
tb, _ = attendance = dl.attendance_table(3, [[0], [], [1, 2]], pd.Series([T(8)] * 3), pd.Series([T(9)] * 3), dl.Rule())
check("satu baris bisa dikreditkan ke beberapa baris File 1 (kunci kembar)", tb["Status durasi"].tolist() == [dl.STATUS_OK, dl.STATUS_OK, dl.STATUS_OK] and tb["Jumlah sesi"].tolist() == [1, 1, 1])
tbd, _ = dl.attendance_from_duration(3, [[0], [0], [1]], pd.Series([30, 35, "x"]), dl.Rule())
check("cadangan kolom durasi: 30 + 35 = 65 Memenuhi; teks tak terbaca dilewati", tbd["Total durasi (menit)"].iloc[0] == 65.0 and tbd["Status durasi"].tolist() == [dl.STATUS_OK, "", ""], tbd)
check("status akhir", [dl.final_status(dl.STATUS_OK, False, False), dl.final_status(dl.STATUS_SHORT, False, False),
      dl.final_status("", True, False), dl.final_status("", False, False), dl.final_status("", False, True)] ==
      ["Hadir", dl.STATUS_SHORT, "Hadir (belum pasti)", "Belum hadir", "Nama/kunci kosong"])
check("penjelasan aturan terbaca manusia", "minimal 60 menit" in dl.Rule().describe() and "08:00–09:00" in dl.Rule(win_start=time(8), win_end=time(9)).describe())

print("== 4. Ujung ke ujung: laporan Zoom + pencocokan nama fleksibel")
F1 = pd.DataFrame({"Unit": ["Keuangan", "Keuangan", "Umum", "Umum", "Data"], "Nama": ["Budi Santoso", "Citra Dewi", "Dedi Kurniawan", "Eka Putri", "Fajar Nugroho"]})
zoom = pd.DataFrame({"Name (Original Name)": ["Keuangan - Budi", "Keuangan - Budi", "Keuangan - Budi S", "citra - keuangan", "Umum - Dedi", "Umum - Dedi",
                                              "Umum - Eka Putri", "Umum - Eka Putri", "Tamu Luar"],
                     "Join Time": ["01/15/2026 07:55:00 AM", "01/15/2026 08:25:00 AM", "01/15/2026 08:10:00 AM", "01/15/2026 08:00:00 AM", "01/15/2026 08:00:00 AM",
                                   "01/15/2026 08:45:00 AM", "01/15/2026 08:00:00 AM", "01/15/2026 08:00:00 AM", "01/15/2026 08:00:00 AM"],
                     "Leave Time": ["01/15/2026 08:20:00 AM", "01/15/2026 09:10:00 AM", "01/15/2026 08:50:00 AM", "01/15/2026 08:40:00 AM", "01/15/2026 08:30:00 AM",
                                    "01/15/2026 09:20:00 AM", "01/15/2026 09:00:00 AM", "01/15/2026 09:00:00 AM", "01/15/2026 09:30:00 AM"]})
noise = ml.parse_noise(ml.DEFAULT_NOISE)
m = ml.Matcher(F1, "Unit", "Nama", noise=noise); res = ml.run(m, F1, zoom, "Name (Original Name)", "Unit", "Nama")
join, _ = dl.parse_times(zoom["Join Time"]); leave, _ = dl.parse_times(zoom["Leave Time"])
rule = dl.Rule(win_start=time(8, 0), win_end=time(9, 0))
tbl, info = dl.attendance_table(len(F1), res.row_persons, join, leave, rule)
tbl.insert(0, "Nama", F1["Nama"]); print(tbl[["Nama", "Total durasi (menit)", "Jumlah sesi", "Reconnect (jeda kecil)", "Status durasi"]].to_string(index=False))
st = dict(zip(tbl["Nama"], tbl["Status durasi"])); tot = dict(zip(tbl["Nama"], tbl["Total durasi (menit)"]))
check("Budi: 3 baris (2 reconnect + 1 sesi lain) digabung jadi satu orang; jam acara memotong 07:55 -> 08:00", tot["Budi Santoso"] == 60.0 and st["Budi Santoso"] == dl.STATUS_OK, tbl.iloc[0].to_dict())
check("Citra hanya 40 menit -> Kurang dari batas", st["Citra Dewi"] == dl.STATUS_SHORT and tot["Citra Dewi"] == 40.0)
check("Dedi putus 15 menit (> toleransi) -> 30 + 15 (jam acara s.d. 09:00) = 45, kurang", tot["Dedi Kurniawan"] == 45.0 and st["Dedi Kurniawan"] == dl.STATUS_SHORT, tot)
check("Eka dua perangkat 60 menit -> memenuhi; Fajar tidak ada sesi", st["Eka Putri"] == dl.STATUS_OK and st["Fajar Nugroho"] == "")
check("tamu tak dikenal tidak dikreditkan ke siapa pun", res.row_persons[-1] == [])
fin = [dl.final_status(s, False, False) for s in tbl["Status durasi"]]
check("status akhir per orang", fin == ["Hadir", dl.STATUS_SHORT, dl.STATUS_SHORT, "Hadir", "Belum hadir"], fin)

print("== 5. Mode persis (kunci email) + penugasan baris")
F1e = pd.DataFrame({"Email": ["a@x.id", "b@x.id", "c@x.id"], "Nama": ["A", "B", "C"]})
Z = pd.DataFrame({"User Email": ["A@x.id", "a@x.id ", "b@x.id", "z@x.id", None], "Join Time": ["01/15/2026 08:00 AM"] * 5, "Leave Time": ["01/15/2026 08:40 AM", "01/15/2026 09:00 AM", "01/15/2026 08:30 AM", "01/15/2026 09:00 AM", "01/15/2026 09:00 AM"]})
asg = ck.row_assignments(F1e, "Email", Z, "User Email")
check("penugasan baris: huruf besar/spasi diabaikan, tak dikenal & kosong = []", asg == [[0], [0], [1], [], []], asg)
j, _ = dl.parse_times(Z["Join Time"]); l, _ = dl.parse_times(Z["Leave Time"])
t, _ = dl.attendance_table(3, asg, j, l, dl.Rule())
check("A: dua sesi tumpang tindih (08:00–09:00) = 60 Memenuhi; B 30 menit Kurang; C tidak ada", t["Status durasi"].tolist() == [dl.STATUS_OK, dl.STATUS_SHORT, ""], t)

print("== 6. Kinerja (20.000 sesi)")
import time as _t
rng = np.random.default_rng(3)
n = 20000; base = pd.Timestamp("2026-01-15 08:00:00")
jj = pd.Series(base + pd.to_timedelta(rng.integers(0, 3000, n), unit="s")); ll = jj + pd.to_timedelta(rng.integers(60, 3000, n), unit="s")
t0 = _t.time(); tb, _ = dl.attendance_table(5000, [[int(rng.integers(0, 5000))] for _ in range(n)], jj, ll, dl.Rule()); el = _t.time() - t0
print(f"    {el:.1f} dtk")
check("selesai dalam waktu wajar (< 20 dtk)", el < 20, el)
strs = pd.Series([f"01/15/2026 {8 + (i % 4):02d}:{i % 60:02d}:00 AM" for i in range(5000)])
t0 = _t.time(); dl.parse_times(strs); el = _t.time() - t0; print(f"    parse 5.000 teks: {el:.1f} dtk")
check("parse 5.000 teks AM/PM < 20 dtk", el < 20, el)

print("\nHASIL:", "SEMUA PENGUJIAN LULUS" if fails == 0 else f"{fails} GAGAL")
sys.exit(1 if fails else 0)
