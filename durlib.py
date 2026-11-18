"""
durlib.py — durasi kehadiran dari log join/leave (mis. laporan peserta Zoom). Tanpa Streamlit, mudah diuji.

Zoom mencatat SATU BARIS PER SESI: orang yang putus jaringan lalu masuk lagi muncul beberapa kali. Alurnya:
  1. Tiap baris dicocokkan ke orang di File 1 (cara persis atau fleksibel — lihat checklib/matchlib).
  2. Sesi tiap orang digabung: sesi yang TUMPANG TINDIH (mis. dua perangkat) tidak dihitung dobel.
  3. Jeda antar-sesi ≤ toleransi dianggap RECONNECT (tersambung); jeda lebih lama dihitung sebagai keluar.
  4. Hadir bila total durasi memenuhi batas (mis. 60 menit).
Opsional: batasi ke jam acara (waktu di ruang tunggu / setelah acara tidak dihitung).
"""
from __future__ import annotations

import re
import warnings
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import Optional

import numpy as np
import pandas as pd

BASE_DATE = date(2000, 1, 1)         # tanggal acuan bila data hanya berisi jam (08:01:23)
_JOIN = re.compile(r"join|bergabung|masuk", re.IGNORECASE)
_LEAVE = re.compile(r"leave|left|keluar|meninggalkan|pergi", re.IGNORECASE)
_DUR = re.compile(r"duration|durasi|time in session|lama", re.IGNORECASE)

STATUS_OK = "Memenuhi durasi"
STATUS_SHORT = "Kurang dari batas"
COLS = ["Total durasi (menit)", "Jumlah sesi", "Join pertama", "Leave terakhir",
        "Reconnect (jeda kecil)", "Jeda panjang", "Status durasi"]


# ---------------------------------------------------------------------------
# Mengenali tabel & kolom
# ---------------------------------------------------------------------------
def find_session_header(raw: pd.DataFrame, max_scan: int = 40) -> Optional[int]:
    """Indeks baris header tabel peserta (memuat kolom join DAN leave). Laporan Zoom CSV biasanya diawali
    ringkasan rapat (Topic, ID, Host, ...) sebelum tabel peserta, sehingga tebakan umum akan keliru."""
    for i in range(min(len(raw), max_scan)):
        cells = [v for v in raw.iloc[i].tolist() if isinstance(v, str)]
        if len(cells) >= 3 and any(_JOIN.search(c) for c in cells) and any(_LEAVE.search(c) for c in cells):
            return i
    return None


def guess_session_columns(cols: list) -> tuple:
    """(kolom join, kolom leave, kolom durasi) — None bila tidak ditemukan."""
    join = next((c for c in cols if _JOIN.search(str(c))), None)
    leave = next((c for c in cols if c != join and _LEAVE.search(str(c))), None)
    dur = next((c for c in cols if c not in (join, leave) and _DUR.search(str(c))), None)
    return join, leave, dur


# ---------------------------------------------------------------------------
# Membaca tanggal-jam
# ---------------------------------------------------------------------------
_NUMDATE = re.compile(r"^\s*(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2,4})")
_TIMEONLY = re.compile(r"^\s*(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])?\s*$")


def _detect_dayfirst(values: list) -> tuple:
    """(dayfirst atau None bila tak ada tanggal numerik, ambigu?)."""
    firsts, seconds = [], []
    for v in values[:3000]:
        m = _NUMDATE.match(v)
        if m:
            firsts.append(int(m.group(1)))
            seconds.append(int(m.group(2)))
    if not firsts:
        return None, False
    if max(firsts) > 12:
        return True, False
    if max(seconds) > 12:
        return False, False
    return False, True               # ambigu (mis. 03/04/2026): anggap Bulan/Hari/Tahun, format bawaan Zoom


def _to_datetime(series: pd.Series, dayfirst: bool) -> pd.Series:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            return pd.to_datetime(series, errors="coerce", dayfirst=dayfirst, format="mixed")
        except (TypeError, ValueError):
            return pd.to_datetime(series, errors="coerce", dayfirst=dayfirst)


def _time_only(text: str) -> Optional[datetime]:
    m = _TIMEONLY.match(text)
    if not m:
        return None
    h, mi, sec, ap = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0), m.group(4)
    if ap:
        if not 1 <= h <= 12:
            return None
        h = h % 12 + (12 if ap.lower() == "pm" else 0)
    if h > 23 or mi > 59 or sec > 59:
        return None
    return datetime.combine(BASE_DATE, time(h, mi, sec))


def parse_times(series: pd.Series, order: str = "auto") -> tuple:
    """Ubah kolom tanggal-jam apa pun (objek datetime/time Excel, atau teks) menjadi datetime.

    order: 'auto' | 'dmy' (Hari/Bulan/Tahun) | 'mdy' (Bulan/Hari/Tahun). Mengembalikan (Series datetime, [catatan]).
    """
    notes: list = []
    vals = series.tolist()
    result: list = [pd.NaT] * len(vals)
    str_idx: list = []
    str_vals: list = []
    non_empty = 0
    for i, v in enumerate(vals):
        if v is None or (not isinstance(v, str) and pd.isna(v)):
            continue
        if isinstance(v, str) and not v.strip():
            continue
        non_empty += 1
        if isinstance(v, datetime):
            result[i] = pd.Timestamp(v.replace(tzinfo=None))
        elif isinstance(v, date):
            result[i] = pd.Timestamp(v)
        elif isinstance(v, time):
            result[i] = pd.Timestamp(datetime.combine(BASE_DATE, v))
        elif isinstance(v, str):
            t = _time_only(v)
            if t is not None:
                result[i] = pd.Timestamp(t)
            else:
                str_idx.append(i)
                str_vals.append(v.strip())
    if str_vals:
        det, ambiguous = _detect_dayfirst(str_vals)
        dayfirst = {"dmy": True, "mdy": False}.get(order, det if det is not None else False)
        if order == "auto" and ambiguous:
            notes.append("Format tanggal ambigu (mis. 03/04/2026 bisa 3 April atau 4 Maret); diasumsikan "
                         "Bulan/Hari/Tahun. Ubah pengaturan format bila salah.")
        parsed = _to_datetime(pd.Series(str_vals), dayfirst)
        for i, ts in zip(str_idx, parsed.tolist()):
            if pd.notna(ts):
                result[i] = pd.Timestamp(ts).tz_localize(None) if getattr(ts, "tzinfo", None) else pd.Timestamp(ts)
    out = pd.Series(pd.to_datetime(result), index=series.index)
    bad = int(out.isna().sum()) - (len(vals) - non_empty)
    if bad > 0:
        notes.append(f"{bad:,} nilai waktu tidak dapat dibaca dan diabaikan.")
    return out, notes


# ---------------------------------------------------------------------------
# Aturan & perhitungan
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class Rule:
    threshold_min: float = 60.0
    inclusive: bool = True             # True: ≥ batas; False: > batas
    tolerance_min: float = 10.0        # jeda ≤ ini dianggap reconnect (tersambung)
    count_gaps: bool = True            # jeda reconnect ikut dihitung sebagai waktu hadir
    method: str = "merged"             # 'merged' = waktu terhubung · 'span' = join pertama s.d. leave terakhir
    win_start: Optional[time] = None   # batasi ke jam acara (opsional)
    win_end: Optional[time] = None

    def passes(self, minutes: float) -> bool:
        need = self.threshold_min * 60
        sec = round(minutes * 60, 6)
        return sec >= need if self.inclusive else sec > need

    def describe(self) -> str:
        op = "minimal" if self.inclusive else "lebih dari"
        how = ("dari join pertama sampai leave terakhir" if self.method == "span" else
               "total waktu terhubung" + (", jeda reconnect ikut dihitung" if self.count_gaps else ""))
        win = f"; hanya jam {self.win_start:%H:%M}–{self.win_end:%H:%M}" if self.win_start and self.win_end else ""
        return (f"Hadir bila {how} {op} {self.threshold_min:g} menit. Jeda ≤ {self.tolerance_min:g} menit dianggap "
                f"reconnect; sesi yang tumpang tindih tidak dihitung dobel{win}.")


def _clip(s: pd.Timestamp, e: pd.Timestamp, rule: Rule):
    if rule.win_start is None or rule.win_end is None:
        return s, e
    day = s.normalize()
    ws = day + pd.Timedelta(hours=rule.win_start.hour, minutes=rule.win_start.minute, seconds=rule.win_start.second)
    we = day + pd.Timedelta(hours=rule.win_end.hour, minutes=rule.win_end.minute, seconds=rule.win_end.second)
    if we <= ws:
        we += pd.Timedelta(days=1)
    s2, e2 = max(s, ws), min(e, we)
    return (s2, e2) if e2 > s2 else None


def _person_stats(ivs: list, rule: Rule) -> tuple:
    ivs = sorted(ivs)
    union: list = []
    for s, e in ivs:                                   # gabungkan sesi yang tumpang tindih (jangan dihitung dobel)
        if union and s <= union[-1][1]:
            union[-1][1] = max(union[-1][1], e)
        else:
            union.append([s, e])
    tol = pd.Timedelta(minutes=rule.tolerance_min)
    covered = sum((e - s for s, e in union), pd.Timedelta(0))
    bridged, n_small, n_long = pd.Timedelta(0), 0, 0
    for a, b in zip(union, union[1:]):
        gap = b[0] - a[1]
        if gap <= tol:
            n_small += 1
            bridged += gap
        else:
            n_long += 1
    if rule.method == "span":
        total = union[-1][1] - union[0][0]
    else:
        total = covered + (bridged if rule.count_gaps else pd.Timedelta(0))
    return total.total_seconds() / 60, len(ivs), union[0][0], union[-1][1], n_small, n_long


def _empty_table(n1: int) -> pd.DataFrame:
    return pd.DataFrame({"Total durasi (menit)": [np.nan] * n1, "Jumlah sesi": [0] * n1,
                         "Join pertama": [pd.NaT] * n1, "Leave terakhir": [pd.NaT] * n1,
                         "Reconnect (jeda kecil)": [0] * n1, "Jeda panjang": [0] * n1, "Status durasi": [""] * n1})


def attendance_table(n1: int, row_persons: list, join: pd.Series, leave: pd.Series, rule: Rule) -> tuple:
    """Durasi per baris File 1 dari log join/leave. Mengembalikan (tabel sejajar posisi baris File 1, info)."""
    by_p: dict = defaultdict(list)
    info = {"rows_matched": 0, "unreadable": 0, "negative": 0, "outside_window": 0}
    for i, plist in enumerate(row_persons):
        if not plist:
            continue
        info["rows_matched"] += 1
        s, e = join.iloc[i], leave.iloc[i]
        if pd.isna(s) or pd.isna(e):
            info["unreadable"] += 1
            continue
        if e < s:
            info["negative"] += 1
            continue
        iv = _clip(s, e, rule)
        if iv is None:
            info["outside_window"] += 1
            continue
        for p in plist:
            by_p[p].append(iv)
    table = _empty_table(n1)
    for p, ivs in by_p.items():
        total, n_sessions, first, last, n_small, n_long = _person_stats(ivs, rule)
        table.loc[p, ["Total durasi (menit)", "Jumlah sesi", "Join pertama", "Leave terakhir",
                      "Reconnect (jeda kecil)", "Jeda panjang"]] = [round(total, 1), n_sessions, first, last, n_small, n_long]
        table.loc[p, "Status durasi"] = STATUS_OK if rule.passes(total) else STATUS_SHORT
    table["Jumlah sesi"] = table["Jumlah sesi"].astype(int)
    return table, info


def attendance_from_duration(n1: int, row_persons: list, minutes: pd.Series, rule: Rule) -> tuple:
    """Cadangan bila file hanya punya kolom durasi (menit) tanpa join/leave: durasi tiap baris dijumlahkan.
    Sesi yang tumpang tindih tidak bisa dideteksi tanpa jam join/leave."""
    vals = pd.to_numeric(minutes, errors="coerce")
    sums: dict = defaultdict(float)
    counts: dict = defaultdict(int)
    info = {"rows_matched": 0, "unreadable": 0, "negative": 0, "outside_window": 0}
    for i, plist in enumerate(row_persons):
        if not plist:
            continue
        info["rows_matched"] += 1
        v = vals.iloc[i]
        if pd.isna(v):
            info["unreadable"] += 1
        elif v < 0:
            info["negative"] += 1
        else:
            for p in plist:
                sums[p] += float(v)
                counts[p] += 1
    table = _empty_table(n1)
    for p, total in sums.items():
        table.loc[p, ["Total durasi (menit)", "Jumlah sesi"]] = [round(total, 1), counts[p]]
        table.loc[p, "Status durasi"] = STATUS_OK if rule.passes(total) else STATUS_SHORT
    table["Jumlah sesi"] = table["Jumlah sesi"].astype(int)
    return table, info


def final_status(duration_status: str, ambiguous: bool, blank: bool) -> str:
    """Status akhir satu orang di File 1."""
    if blank:
        return "Nama/kunci kosong"
    if duration_status == STATUS_OK:
        return "Hadir"
    if duration_status == STATUS_SHORT:
        return STATUS_SHORT
    return "Hadir (belum pasti)" if ambiguous else "Belum hadir"
