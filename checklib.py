"""
checklib.py — logika inti menu "Cek Absensi" (tanpa Streamlit, sehingga mudah diuji).

Membandingkan dua daftar (File 1 = daftar acuan, File 2 = daftar pembanding/yang hadir) lewat satu kolom kunci
(NIP, NIK, ID, email, nama, ...) dan menghasilkan:
  * sudah ada   : baris File 1 yang kuncinya ada di File 2 (+ jumlah kemunculannya)
  * belum ada   : baris File 1 yang kuncinya tidak ada di File 2
  * hanya di 2  : baris File 2 yang kuncinya tidak ada di File 1 (mis. hadir tapi tidak terdaftar)
  * perlu dicek : kunci kosong dan kunci kembar

Perbaikan dari skrip awal (Menu 1 di app lama):
  * Kunci tidak lagi dipotong pada titik pertama (`budi.santoso@x.com` dulu menjadi `budi`).
  * Kunci kosong tidak pernah dianggap cocok (dulu bisa "sudah ada" jika File 2 punya satu sel kosong).
  * Aturan pencocokan eksplisit: abaikan huruf besar-kecil, spasi, hanya angka, nol di depan.
  * Peringatan untuk ID ≥16 digit yang tersimpan sebagai *angka* di Excel (digit belakang sudah rusak).
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime

import numpy as np
import pandas as pd

import sheetlib as sl

COUNT_COL = "Jumlah di File 2"


class CheckError(sl.SpreadsheetError):
    """Pesan aman & ramah untuk ditampilkan ke pengguna."""


@dataclass(frozen=True)
class Options:
    ignore_case: bool = True       # "BUDI" = "budi"
    ignore_spaces: bool = True     # "1980 0101 2005" = "198001012005" (semua spasi dibuang)
    digits_only: bool = False      # ambil angka saja (cocok untuk NIP/NIK yang ditulis dengan pemisah)
    strip_zeros: bool = False      # "00123" = "123" (hanya untuk kunci yang seluruhnya angka)


# ---------------------------------------------------------------------------
# Normalisasi kunci
# ---------------------------------------------------------------------------
def _key_text(v) -> str:
    """Nilai sel -> teks kunci, tanpa 'nan', '.0', atau jam 00:00:00."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, (bool, np.bool_)):
        return str(bool(v))
    if isinstance(v, (int, np.integer)):
        return str(int(v))
    if isinstance(v, (float, np.floating)):
        f = float(v)
        return str(int(f)) if f.is_integer() else np.format_float_positional(f, trim="-")
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return str(v)


key_text = _key_text   # alias publik (dipakai matchlib)


def normalize_keys(series: pd.Series, opts: Options = Options()) -> pd.Series:
    """Samakan bentuk kunci agar pencocokan akurat. Kunci yang hasilnya kosong berarti 'tidak punya kunci'."""
    t = series.map(_key_text).astype(object)
    t = t.str.replace("\u00a0", " ", regex=False).str.replace(r"[\u200b-\u200d\ufeff]", "", regex=True)
    if opts.digits_only:
        t = t.str.replace(r"\D", "", regex=True)
    elif opts.ignore_spaces:
        t = t.str.replace(r"\s+", "", regex=True)
    else:
        t = t.str.replace(r"\s+", " ", regex=True).str.strip()
    if opts.ignore_case:
        t = t.str.casefold()
    if opts.strip_zeros:
        stripped = t.str.lstrip("0").where(lambda x: x != "", "0")
        t = t.where(~t.str.fullmatch(r"\d+"), stripped)
    return t


def precision_warning_count(series: pd.Series) -> int:
    """Jumlah nilai ANGKA (bukan teks) ≥ 1e15 pada kolom kunci.

    Excel hanya menyimpan 15 digit signifikan: NIK (16 digit) atau NIP (18 digit) yang diketik/ditempel sebagai angka
    sudah kehilangan digit belakangnya (jadi nol). Nilai bisa terbaca sebagai int maupun float, keduanya dihitung.
    Kunci yang tersimpan sebagai teks aman dan tidak dihitung.
    """
    def suspicious(v) -> bool:
        if isinstance(v, (bool, np.bool_)):
            return False
        if isinstance(v, (int, np.integer)):
            return abs(int(v)) >= 10 ** 15
        if isinstance(v, (float, np.floating)):
            return bool(np.isfinite(v)) and abs(float(v)) >= 1e15
        return False
    return int(series.map(suspicious).sum()) if len(series) else 0


# ---------------------------------------------------------------------------
# Tebakan kolom kunci
# ---------------------------------------------------------------------------
_HINTS = (r"^nip$", r"^nik$", r"^id$", r"nip|nik", r"\bid\b|kode|no\.? ?(pegawai|induk|identitas|peserta)",
          r"e-?mail", r"nama|name")


def _norm(s) -> str:
    return re.sub(r"[\s_]+", " ", str(s).strip()).lower()


def guess_key_columns(cols1: list, cols2: list) -> tuple:
    """Tebak kolom kunci di kedua file: utamakan kolom bernama sama, lalu petunjuk umum (NIP/NIK/ID/email/nama)."""
    n1, n2 = {_norm(c): c for c in cols1}, {_norm(c): c for c in cols2}
    common = [n for n in n1 if n in n2]
    for pat in _HINTS:
        for n in common:
            if re.search(pat, n):
                return n1[n], n2[n]
    for pat in _HINTS:
        a = next((c for c in cols1 if re.search(pat, _norm(c))), None)
        b = next((c for c in cols2 if re.search(pat, _norm(c))), None)
        if a is not None and b is not None:
            return a, b
    if common:
        return n1[common[0]], n2[common[0]]
    return cols1[0], cols2[0]


# ---------------------------------------------------------------------------
# Perbandingan
# ---------------------------------------------------------------------------
@dataclass
class CompareResult:
    found: pd.DataFrame
    missing: pd.DataFrame
    only_in_2: pd.DataFrame
    blank_1: pd.DataFrame
    dup_1: pd.DataFrame
    dup_2: pd.DataFrame
    summary: dict = field(default_factory=dict)


def _free_name(existing, name: str) -> str:
    cand, n = name, 2
    while cand in set(existing):
        cand = f"{name} ({n})"
        n += 1
    return cand


def _dups(df: pd.DataFrame, key: str, norm_s: pd.Series, ok: pd.Series) -> pd.DataFrame:
    vc = norm_s[ok].value_counts()
    vc = vc[vc > 1]
    if vc.empty:
        return pd.DataFrame(columns=["Kunci", "Jumlah baris"])
    first = df.loc[ok].assign(__k=norm_s[ok]).drop_duplicates("__k").set_index("__k")[key]
    return pd.DataFrame({"Kunci": first.reindex(vc.index).map(_key_text).to_numpy(),
                         "Jumlah baris": vc.to_numpy()})


def compare(df1: pd.DataFrame, key1: str, df2: pd.DataFrame, key2: str,
            opts: Options = Options(), extra_cols: tuple = ()) -> CompareResult:
    if key1 not in df1.columns or key2 not in df2.columns:
        raise CheckError("Kolom kunci tidak ditemukan pada salah satu file.")
    n1, n2 = normalize_keys(df1[key1], opts), normalize_keys(df2[key2], opts)
    ok1, ok2 = n1 != "", n2 != ""

    counts2 = n2[ok2].value_counts()
    in2 = ok1 & n1.isin(counts2.index)

    found = df1.loc[in2].copy()
    found[_free_name(df1.columns, COUNT_COL)] = n1[in2].map(counts2).astype("int64").to_numpy()
    extras = [c for c in extra_cols if c in df2.columns]
    if extras:
        first2 = df2.loc[ok2, extras].assign(__k=n2[ok2]).drop_duplicates("__k").set_index("__k")
        for c in extras:
            found[_free_name(found.columns, f"File 2: {c}")] = n1[in2].map(first2[c]).to_numpy()

    missing = df1.loc[ok1 & ~in2]
    blank_1 = df1.loc[~ok1]
    only_in_2 = df2.loc[ok2 & ~n2.isin(set(n1[ok1]))]
    dup_1, dup_2 = _dups(df1, key1, n1, ok1), _dups(df2, key2, n2, ok2)

    valid_1 = int(ok1.sum())
    summary = {
        "total_1": len(df1), "total_2": len(df2), "valid_1": valid_1,
        "found": len(found), "missing": len(missing), "blank_1": len(blank_1), "blank_2": int((~ok2).sum()),
        "only_2": len(only_in_2),
        "found_unique": int(n1[in2].nunique()), "missing_unique": int(n1[ok1 & ~in2].nunique()),
        "pct_found": round(len(found) / valid_1 * 100, 1) if valid_1 else 0.0,
        "dup_keys_1": len(dup_1), "dup_keys_2": len(dup_2),
    }
    r = lambda d: d.reset_index(drop=True)   # noqa: E731
    return CompareResult(r(found), r(missing), r(only_in_2), r(blank_1), dup_1, dup_2, summary)


def build_report(res: CompareResult) -> dict:
    """Susun sheet laporan Excel: Ringkasan + hasil. Sheet 'perlu dicek' hanya dibuat jika ada isinya."""
    s = res.summary
    rows = [
        ("Total baris File 1", s["total_1"]),
        ("Baris File 1 tanpa kunci (tidak dicocokkan)", s["blank_1"]),
        ("Sudah ada di File 2", s["found"]),
        ("Belum ada di File 2", s["missing"]),
        ("Persentase sudah ada (dari baris yang punya kunci, %)", s["pct_found"]),
        ("Total baris File 2", s["total_2"]),
        ("Baris File 2 tanpa kunci", s["blank_2"]),
        ("Hanya ada di File 2 (tidak ada di File 1)", s["only_2"]),
        ("Kunci kembar di File 1", s["dup_keys_1"]),
        ("Kunci kembar di File 2", s["dup_keys_2"]),
    ]
    sheets = {
        "Ringkasan": pd.DataFrame(rows, columns=["Keterangan", "Jumlah"]),
        "Sudah ada": res.found,
        "Belum ada": res.missing,
        "Hanya di File 2": res.only_in_2,
    }
    if len(res.blank_1):
        sheets["Kunci kosong File 1"] = res.blank_1
    if len(res.dup_1):
        sheets["Kunci kembar File 1"] = res.dup_1
    if len(res.dup_2):
        sheets["Kunci kembar File 2"] = res.dup_2
    return sheets
