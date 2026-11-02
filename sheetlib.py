"""
sheetlib.py — logika inti aplikasi (tanpa Streamlit, sehingga mudah diuji).

Isi:
  1. Validasi file unggahan (hanya spreadsheet asli yang lolos)
  2. Pembacaan workbook (xlsx, xlsm, xls, ods, csv)
  3. Pembersihan tabel & deteksi tipe kolom
  4. Fungsi analisis (agregasi, korelasi, kualitas data)
  5. Ekspor & data contoh
"""
from __future__ import annotations

import csv
import io
import itertools
import re
import warnings
import zipfile
from datetime import date, datetime

import numpy as np
import pandas as pd

# ---------------------------------------------------------------------------
# Konfigurasi
# ---------------------------------------------------------------------------
ALLOWED_EXTENSIONS = ("xlsx", "xlsm", "xls", "ods", "csv")
MAX_FILE_MB = 25            # batas ukuran file unggahan
MAX_UNZIPPED_MB = 200       # batas ukuran setelah diekstrak (anti zip-bomb)
MAX_ROWS = 200_000          # batas baris per sheet
MAX_COLS = 300              # batas kolom per sheet

XLS_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"   # tanda file OLE (.xls lama)
BINARY_SIGNATURES = (b"PK\x03\x04", b"%PDF", b"MZ", XLS_MAGIC,
                     b"\x89PNG", b"GIF8", b"\xff\xd8\xff", b"Rar!", b"\x1f\x8b")

ROLE_LABEL = {
    "numerik": "Angka",
    "kategori": "Kategori",
    "tanggal": "Tanggal",
    "teks": "Teks bebas",
    "id": "ID / nomor urut",
}

AGG_CHOICES = {
    "Jumlah (sum)": "sum",
    "Rata-rata": "mean",
    "Median": "median",
    "Minimum": "min",
    "Maksimum": "max",
    "Jumlah nilai unik": "nunique",
}

GRANULARITY = {"Harian": "D", "Mingguan": "W", "Bulanan": "M", "Kuartal": "Q", "Tahunan": "Y"}


class SpreadsheetError(Exception):
    """Error yang pesannya aman & ramah untuk ditampilkan ke pengguna."""


# ---------------------------------------------------------------------------
# 1. Validasi unggahan
# ---------------------------------------------------------------------------
def get_extension(filename: str) -> str:
    return filename.rsplit(".", 1)[-1].lower() if "." in filename else ""


def _check_zip(data: bytes, required: str, label: str) -> zipfile.ZipFile:
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        raise SpreadsheetError(
            f"Isi file bukan dokumen {label} yang valid. File mungkin rusak, "
            "dilindungi password (hapus password lebih dulu), atau hanya diganti ekstensinya."
        )
    if required not in zf.namelist():
        raise SpreadsheetError(f"File ini bukan spreadsheet {label} yang asli (struktur dokumen tidak sesuai).")
    if sum(i.file_size for i in zf.infolist()) > MAX_UNZIPPED_MB * 1024 * 1024:
        raise SpreadsheetError("Isi file terlalu besar setelah diekstrak. Kurangi ukuran data lalu coba lagi.")
    return zf


def validate_upload(filename: str, data: bytes) -> str:
    """Pastikan file benar-benar spreadsheet (bukan sekadar ekstensinya).

    Mengembalikan jenis file ('xlsx', 'xlsm', 'xls', 'ods', 'csv') atau melempar SpreadsheetError.
    """
    ext = get_extension(filename)
    if ext not in ALLOWED_EXTENSIONS:
        raise SpreadsheetError(
            f"Format .{ext or '?'} tidak didukung. Unggah file spreadsheet: "
            + ", ".join(f".{e}" for e in ALLOWED_EXTENSIONS) + "."
        )
    if not data:
        raise SpreadsheetError("File kosong.")
    if len(data) > MAX_FILE_MB * 1024 * 1024:
        raise SpreadsheetError(f"Ukuran file melebihi batas {MAX_FILE_MB} MB.")

    if ext in ("xlsx", "xlsm"):
        _check_zip(data, "xl/workbook.xml", "Excel")
    elif ext == "ods":
        zf = _check_zip(data, "content.xml", "OpenDocument")
        mime = zf.read("mimetype").decode("ascii", "ignore") if "mimetype" in zf.namelist() else ""
        if "spreadsheet" not in mime:
            raise SpreadsheetError("File OpenDocument ini bukan spreadsheet.")
    elif ext == "xls":
        if not data.startswith(XLS_MAGIC):
            raise SpreadsheetError("Isi file bukan spreadsheet .xls yang valid.")
    else:  # csv
        if b"\x00" in data[:8192] and not data.startswith((b"\xff\xfe", b"\xfe\xff")):
            raise SpreadsheetError("Isi file bukan teks CSV (terdeteksi data biner).")
        if data.startswith(BINARY_SIGNATURES):
            raise SpreadsheetError("File ini bukan CSV; isinya format lain yang hanya diganti ekstensinya.")
    return ext


# ---------------------------------------------------------------------------
# 2. Membaca workbook (mentah, tanpa header)
# ---------------------------------------------------------------------------
def _decode_text(data: bytes) -> str:
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16")
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("latin-1")


def _detect_delimiter(text: str) -> str:
    """Pilih pemisah yang membuat jumlah kolom paling konsisten di banyak baris.

    (csv.Sniffer sering keliru pada angka berformat Indonesia seperti 1.250.000,50.)
    """
    lines = [ln for ln in text[:50000].splitlines() if ln.strip()][:40]
    best, best_score = ",", (0, 0)
    for d in (",", ";", "\t", "|"):
        try:
            counts = [len(r) for r in csv.reader(lines, delimiter=d)]
        except csv.Error:
            continue
        mode = max(set(counts), key=counts.count)
        score = (counts.count(mode) if mode > 1 else 0, mode)
        if score > best_score:
            best, best_score = d, score
    return best


def _read_csv_raw(data: bytes) -> pd.DataFrame:
    text = _decode_text(data)
    sep = _detect_delimiter(text)
    try:
        rows = list(itertools.islice(csv.reader(io.StringIO(text), delimiter=sep), MAX_ROWS + 1))
    except csv.Error as e:
        raise SpreadsheetError(f"CSV tidak dapat dibaca: {e}")
    if not rows:
        raise SpreadsheetError("File CSV kosong.")
    width = min(max(len(r) for r in rows), MAX_COLS + 1)
    padded = [r[:width] + [None] * (width - len(r[:width])) for r in rows]
    df = pd.DataFrame(padded, dtype=object)
    return df.replace(r"^\s*$", np.nan, regex=True)


def load_workbook(kind: str, data: bytes) -> tuple[dict[str, pd.DataFrame], list[str]]:
    """Baca semua sheet sebagai tabel mentah (header=None). Mengembalikan (sheets, catatan)."""
    notes: list[str] = []
    if kind == "csv":
        sheets = {"Data": _read_csv_raw(data)}
    else:
        engine = {"xlsx": "openpyxl", "xlsm": "openpyxl", "xls": "xlrd", "ods": "odf"}[kind]
        try:
            sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None,
                                   dtype=object, engine=engine, nrows=MAX_ROWS + 1)
        except ImportError:
            raise SpreadsheetError(f"Pustaka pembaca file .{kind} belum terpasang di server (lihat requirements.txt).")
        except Exception:
            raise SpreadsheetError(
                "File tidak dapat dibaca. Kemungkinan rusak, dilindungi password, atau formatnya tidak valid."
            )
    for name in list(sheets):
        df = sheets[name]
        if len(df) > MAX_ROWS:
            df = df.iloc[:MAX_ROWS]
            notes.append(f"Sheet '{name}' dipotong ke {MAX_ROWS:,} baris pertama.")
        if df.shape[1] > MAX_COLS:
            df = df.iloc[:, :MAX_COLS]
            notes.append(f"Sheet '{name}' dipotong ke {MAX_COLS} kolom pertama.")
        sheets[name] = df
    return sheets, notes


# ---------------------------------------------------------------------------
# 3. Pembersihan tabel & deteksi tipe
# ---------------------------------------------------------------------------
def guess_header_row(raw: pd.DataFrame, max_scan: int = 30) -> int:
    """Tebak baris header (indeks mulai 0): baris pertama yang hampir penuh, berisi teks unik."""
    if raw.empty:
        return 0
    scan = raw.head(max_scan)
    counts = scan.notna().sum(axis=1)
    max_count = int(counts.max())
    if max_count == 0:
        return 0
    need = max(1, int(np.ceil(0.6 * max_count)))
    for i in range(len(scan)):
        row = scan.iloc[i].dropna()
        if len(row) < need:
            continue
        n_str = sum(isinstance(v, str) for v in row)
        labels = row.astype(str).str.strip().str.lower()
        if n_str / len(row) >= 0.8 and labels.nunique() == len(row):
            return i
    return 0


def _make_headers(values: list) -> list[str]:
    names, seen = [], {}
    for i, h in enumerate(values):
        if h is None or (not isinstance(h, str) and pd.isna(h)):
            name = ""
        elif isinstance(h, float) and h.is_integer():
            name = str(int(h))
        else:
            name = re.sub(r"\s+", " ", str(h)).strip()
        if not name:
            name = f"Kolom {i + 1}"
        if name in seen:
            seen[name] += 1
            name = f"{name} ({seen[name]})"
        else:
            seen[name] = 1
        names.append(name)
    return names


make_headers = _make_headers   # alias publik (dipakai AUTOCert)


def _strip(v):
    if isinstance(v, str):
        t = v.strip()
        return t if t else np.nan
    return v


SUMMARY_LABELS = {"total", "jumlah", "grand total", "subtotal", "sub total", "rata-rata", "rata rata", "sum", "average"}


def _drop_summary_rows(body: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Buang baris 'Total/Jumlah' di bawah tabel agar tidak terhitung dua kali."""
    mask = pd.Series(False, index=body.index)
    for c in body.columns[:3]:
        labels = body[c].map(lambda v: v.strip().lower() if isinstance(v, str) else None)
        mask |= labels.isin(SUMMARY_LABELS)
    return body[~mask], int(mask.sum())


_ID_MONTHS = {
    "januari": "January", "februari": "February", "maret": "March", "april": "April", "mei": "May",
    "juni": "June", "juli": "July", "agustus": "August", "september": "September",
    "oktober": "October", "november": "November", "desember": "December",
    "agu": "Aug", "ags": "Aug", "okt": "Oct", "des": "Dec",
}
_MONTH_RE = re.compile(r"\b(" + "|".join(_ID_MONTHS) + r")\b", re.IGNORECASE)
_DATE_NUM = re.compile(r"^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}([ T]\d{1,2}:\d{2}(:\d{2}(\.\d+)?)?)?$")
_DATE_TEXT = re.compile(r"^\d{1,2}\s+[A-Za-z]{3,}\.?,?\s+\d{4}$")
_DT_TYPES = (datetime, date, pd.Timestamp)


def _translate_months(v):
    return _MONTH_RE.sub(lambda m: _ID_MONTHS[m.group(0).lower()], v) if isinstance(v, str) else v


def _parse_dates(series: pd.Series) -> pd.Series:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            return pd.to_datetime(series, errors="coerce", dayfirst=True, format="mixed")
        except (TypeError, ValueError):
            return pd.to_datetime(series, errors="coerce", dayfirst=True)


def _try_datetime(s: pd.Series, nn: pd.Series):
    if nn.map(lambda v: isinstance(v, _DT_TYPES)).mean() >= 0.9:
        return pd.to_datetime(s.map(lambda v: v if isinstance(v, _DT_TYPES) else pd.NaT), errors="coerce")
    strs = nn[nn.map(lambda v: isinstance(v, str))]
    if len(strs) / len(nn) < 0.9:
        return None
    sample = strs.head(200).map(lambda t: _translate_months(t.strip()))
    if sample.map(lambda t: bool(_DATE_NUM.match(t) or _DATE_TEXT.match(t))).mean() < 0.9:
        return None
    parsed = _parse_dates(s.map(_translate_months))
    return parsed if parsed.notna().sum() / len(nn) >= 0.9 else None


def _clean_number_str(v, dot_thousands: bool = False, comma_thousands: bool = False):
    """Rapikan teks angka: 'Rp 1.234.567,50' -> '1234567.50', '85%' -> '85'.

    Angka seperti '1.234' ambigu (1,234 atau seribu dua ratus tiga puluh empat). Jika kolom yang sama memuat
    gaya Indonesia (mis. '1.250.000,50' atau awalan 'Rp'), titik dianggap pemisah ribuan (dot_thousands).
    """
    if not isinstance(v, str):
        return v
    t = v.strip().replace("\u00a0", " ")
    t = re.sub(r"^(rp\.?|idr|usd|\$|€)\s*", "", t, flags=re.IGNORECASE).rstrip("%").strip()
    if re.fullmatch(r"-?\d{1,3}(\.\d{3}){2,}(,\d+)?", t) or re.fullmatch(r"-?\d{1,3}(\.\d{3})+,\d+", t) \
            or (dot_thousands and re.fullmatch(r"-?\d{1,3}(\.\d{3})+", t)):
        t = t.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(,\d{3}){2,}(\.\d+)?", t) or re.fullmatch(r"-?\d{1,3}(,\d{3})+\.\d+", t) \
            or (comma_thousands and re.fullmatch(r"-?\d{1,3}(,\d{3})+", t)):
        t = t.replace(",", "")
    elif re.fullmatch(r"-?\d+,\d+", t):
        t = t.replace(",", ".")
    return t


def _number_style(strs: pd.Series) -> tuple[bool, bool]:
    """Deteksi gaya penulisan angka di satu kolom: (titik=ribuan, koma=ribuan)."""
    indo = strs.str.contains(r"^(?:rp\.?|idr)\s*-?\d{1,3}(?:\.\d{3})+$|\d{1,3}(?:\.\d{3}){2,}|\d{1,3}(?:\.\d{3})+,\d+",
                             case=False, regex=True).any()
    eng = strs.str.contains(r"\d{1,3}(?:,\d{3}){2,}|\d{1,3}(?:,\d{3})+\.\d+", regex=True).any()
    return bool(indo), bool(eng) and not indo


def _maybe_int(out: pd.Series) -> pd.Series:
    if len(out) and out.notna().all() and np.isfinite(out).all() \
            and (out == out.round()).all() and out.abs().max() < 9e15:
        return out.astype("int64")
    return out.astype("float64")


def _try_numeric(s: pd.Series, nn: pd.Series):
    strs = nn[nn.map(lambda v: isinstance(v, str))]
    dot_k, comma_k = False, False
    if len(strs):
        if strs.str.fullmatch(r"(0\d+|\+\d+)").any():
            return None  # kode berawalan 0 / nomor telepon: biarkan sebagai teks
        dot_k, comma_k = _number_style(strs)

    def clean(v):
        return _clean_number_str(v, dot_k, comma_k)

    conv = pd.to_numeric(nn.map(clean), errors="coerce")
    if conv.notna().sum() / len(nn) < 0.9:
        return None
    return _maybe_int(pd.to_numeric(s.map(clean), errors="coerce"))


def _infer_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_bool_dtype(s):
        return s.map(lambda v: "Ya" if v else "Tidak")
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    if pd.api.types.is_numeric_dtype(s):
        return _maybe_int(pd.to_numeric(s, errors="coerce"))
    nn = s.dropna()
    if nn.empty:
        return s
    if nn.map(lambda v: isinstance(v, (bool, np.bool_))).all():
        return s.map(lambda v: v if pd.isna(v) else ("Ya" if v else "Tidak"))
    for attempt in (_try_datetime, _try_numeric):
        result = attempt(s, nn)
        if result is not None:
            return result
    return s.map(lambda v: v if pd.isna(v) else str(v))


def build_dataframe(raw: pd.DataFrame, header_row: int = 0, drop_duplicates: bool = False,
                    drop_summary: bool = True) -> tuple[pd.DataFrame, int]:
    """Ubah tabel mentah menjadi DataFrame rapi. Mengembalikan (df, jumlah_baris_total_dibuang)."""
    if raw.empty:
        return pd.DataFrame(), 0
    header_row = max(0, min(int(header_row), len(raw) - 1))
    body = raw.iloc[header_row + 1:].copy()
    body.columns = _make_headers(raw.iloc[header_row].tolist())
    body = body.dropna(how="all", axis=0).dropna(how="all", axis=1)
    for c in body.columns:
        body[c] = body[c].map(_strip)
    removed = 0
    if drop_summary and not body.empty:
        body, removed = _drop_summary_rows(body)
    body = body.reset_index(drop=True)
    for c in body.columns:
        body[c] = _infer_series(body[c])
    if drop_duplicates:
        body = body.drop_duplicates().reset_index(drop=True)
    return body, removed


_ID_NAME = re.compile(r"\b(no|nomor|id|kode|nip|nik|urut)\b", re.IGNORECASE)
_PERIOD_NAME = re.compile(r"\b(tahun|year|bulan|month|kuartal|quarter|semester|minggu|week)\b", re.IGNORECASE)


def column_roles(df: pd.DataFrame) -> dict[str, str]:
    """Klasifikasikan kolom: numerik / kategori / tanggal / teks / id."""
    roles: dict[str, str] = {}
    for c in df.columns:
        s = df[c]
        nn = s.dropna()
        if pd.api.types.is_datetime64_any_dtype(s):
            roles[c] = "tanggal"
        elif pd.api.types.is_numeric_dtype(s) and not pd.api.types.is_bool_dtype(s):
            is_int = pd.api.types.is_integer_dtype(s)
            unique = nn.nunique()
            if is_int and len(nn) > 1 and unique == len(nn) and (_ID_NAME.search(c) or nn.is_monotonic_increasing):
                roles[c] = "id"
            elif is_int and _PERIOD_NAME.search(c) and unique <= 12:
                roles[c] = "kategori"
            else:
                roles[c] = "numerik"
        else:
            unique = nn.nunique()
            all_unique_big = unique == len(nn) and len(nn) > 20
            roles[c] = "kategori" if (unique <= 50 and not all_unique_big) or unique <= 0.02 * max(len(nn), 1) else "teks"
    return roles


# ---------------------------------------------------------------------------
# 4. Analisis
# ---------------------------------------------------------------------------
def quality_report(df: pd.DataFrame, roles: dict[str, str]) -> pd.DataFrame:
    n = len(df)
    rows = []
    for c in df.columns:
        nn = df[c].dropna()
        rows.append({
            "Kolom": c,
            "Jenis": ROLE_LABEL[roles[c]],
            "Terisi": int(len(nn)),
            "Kosong": int(n - len(nn)),
            "% Kosong": round((n - len(nn)) / n * 100, 1) if n else 0.0,
            "Nilai unik": int(nn.nunique()),
            "Contoh": str(nn.iloc[0])[:40] if len(nn) else "",
        })
    return pd.DataFrame(rows)


def describe_numeric(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    if not cols:
        return pd.DataFrame()
    d = df[cols].describe().T
    d.insert(0, "Kolom", d.index)
    d = d.rename(columns={"count": "Jumlah", "mean": "Rata-rata", "std": "Std. deviasi", "min": "Min",
                          "25%": "Q1", "50%": "Median", "75%": "Q3", "max": "Maks"})
    return d.round(2).reset_index(drop=True)


def default_agg(col: str) -> str:
    """Kolom bertipe rasio/harga/skor lebih masuk akal dirata-rata daripada dijumlahkan."""
    return "Rata-rata" if re.search(r"%|persen|rata|rasio|ratio|rate|skor|score|nilai|harga|price|suhu", col, re.I) \
        else "Jumlah (sum)"


def bucket_dates(s: pd.Series, label: str) -> pd.Series:
    freq = GRANULARITY.get(label)
    if not freq:
        return s
    return s.dt.to_period(freq).dt.to_timestamp()


def auto_granularity(s: pd.Series) -> str:
    nn = s.dropna()
    if nn.empty:
        return "Asli"
    span = (nn.max() - nn.min()).days
    if span <= 45:
        return "Harian"
    if span <= 900:
        return "Bulanan"
    return "Kuartal" if span <= 3650 else "Tahunan"


def aggregate(df: pd.DataFrame, by: list[str], value: str | None = None, agg: str = "size") -> pd.DataFrame:
    """Kelompokkan & hitung. Hasil: kolom `by` + kolom 'Nilai'."""
    by = list(dict.fromkeys(b for b in by if b))
    work = df[by].copy()
    for c in by:
        if not (pd.api.types.is_datetime64_any_dtype(work[c]) or pd.api.types.is_numeric_dtype(work[c])):
            work[c] = work[c].astype(object).where(work[c].notna(), "(Kosong)")
    if agg == "size" or value is None:
        return work.groupby(by, sort=False).size().reset_index(name="Nilai")
    work["__v"] = df[value].to_numpy()
    out = work.groupby(by, sort=False)["__v"].agg(agg).reset_index()
    return out.rename(columns={"__v": "Nilai"})


def top_correlations(corr: pd.DataFrame, n: int = 5) -> pd.DataFrame:
    cols = list(corr.columns)
    pairs = [(a, b, corr.loc[a, b]) for i, a in enumerate(cols) for b in cols[i + 1:] if pd.notna(corr.loc[a, b])]
    pairs.sort(key=lambda t: abs(t[2]), reverse=True)
    out = pd.DataFrame(pairs[:n], columns=["Variabel A", "Variabel B", "Korelasi"])
    out["Korelasi"] = out["Korelasi"].round(3)
    return out


# ---------------------------------------------------------------------------
# 5. Ekspor & data contoh
# ---------------------------------------------------------------------------
def to_csv_bytes(df: pd.DataFrame) -> bytes:
    safe = df.copy()
    for c in safe.select_dtypes(exclude=["number", "datetime"]).columns:
        safe[c] = safe[c].map(lambda v: "'" + v if isinstance(v, str) and v[:1] in ("=", "@") else v)
    return safe.to_csv(index=False).encode("utf-8-sig")


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Data")
        ws = writer.sheets["Data"]
        for row in ws.iter_rows(min_row=2):  # jangan biarkan teks berawalan '=' dieksekusi sebagai rumus
            for cell in row:
                if cell.data_type == "f":
                    cell.data_type = "s"
    return buf.getvalue()


def make_sample_xlsx() -> bytes:
    """Buat file Excel contoh (dengan judul di atas tabel & 2 sheet) untuk demo."""
    rng = np.random.default_rng(42)
    n = 400
    produk = {
        "Elektronik": (["Headphone", "Power Bank", "Keyboard", "Webcam"], 350_000),
        "Fashion": (["Kaos", "Jaket", "Sepatu", "Tas"], 180_000),
        "Makanan": (["Kopi", "Snack Box", "Madu", "Teh"], 60_000),
        "Rumah Tangga": (["Lampu", "Rak Buku", "Blender", "Kipas"], 220_000),
    }
    kategori = rng.choice(list(produk), n)
    nama = [rng.choice(produk[k][0]) for k in kategori]
    harga = np.array([produk[k][1] for k in kategori]) * rng.uniform(0.8, 1.3, n)
    qty = rng.integers(1, 15, n)
    diskon = rng.choice([0, 5, 10, 15, 20], n)
    df = pd.DataFrame({
        "Tanggal": pd.Timestamp("2025-01-01") + pd.to_timedelta(rng.integers(0, 365, n), unit="D"),
        "Wilayah": rng.choice(["Jakarta", "Bandung", "Surabaya", "Medan", "Makassar"], n, p=[.3, .2, .2, .15, .15]),
        "Kategori": kategori,
        "Produk": nama,
        "Jumlah Terjual": qty,
        "Harga Satuan": (harga // 1000 * 1000).astype(int),
        "Diskon (%)": diskon,
    })
    df["Total Penjualan"] = (df["Jumlah Terjual"] * df["Harga Satuan"] * (1 - df["Diskon (%)"] / 100)).round().astype(int)
    df["Metode Bayar"] = rng.choice(["Transfer", "E-Wallet", "Kartu Kredit", "COD"], n)
    df = df.sort_values("Tanggal").reset_index(drop=True)
    target = pd.DataFrame({"Wilayah": ["Jakarta", "Bandung", "Surabaya", "Medan", "Makassar"],
                           "Target Tahunan": [900_000_000, 600_000_000, 600_000_000, 450_000_000, 450_000_000]})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Penjualan", startrow=2)
        w.sheets["Penjualan"]["A1"] = "Laporan Penjualan Contoh 2025"
        target.to_excel(w, index=False, sheet_name="Target")
    return buf.getvalue()
