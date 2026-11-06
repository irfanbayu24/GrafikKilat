"""
certlib.py — logika inti AUTOCert (tanpa Streamlit, sehingga mudah diuji).

Template Word (.docx) bertanda «Nama» / {{Nama}} / <<Nama>> / [Nama]  +  data peserta (spreadsheet)
  ->  satu sertifikat per peserta (DOCX atau PDF)  ->  ZIP.

Perbedaan penting dibanding skrip awal:
  * Tidak memakai `taskkill WINWORD.EXE` (itu menutup semua Word milik pengguna lain).
  * Semua berkas sementara berada di folder sementara unik per proses (aman untuk banyak pengguna).
  * Tag juga diganti di text box, tabel, header & footer, serta tag yang terpecah antar-run (format tetap).
  * Dua mesin PDF: Microsoft Word (Windows) atau LibreOffice (Windows/Linux/Mac). Tanpa keduanya: DOCX.
  * Template diperiksa: harus .docx asli, tanpa makro, tanpa field/relasi yang menarik berkas luar.
"""
from __future__ import annotations

import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unicodedata
import zipfile
from bisect import bisect_right
from dataclasses import dataclass, field
from datetime import date, datetime
from itertools import accumulate
from pathlib import Path
from typing import Callable, Optional

import numpy as np
import pandas as pd
from lxml import etree

import sheetlib as sl

try:
    from docx import Document
except ImportError:  # pragma: no cover
    Document = None

MAX_TEMPLATE_MB = 10
MAX_UNZIPPED_MB = 100
MAX_PARTICIPANTS = 1000
LO_CHUNK = 40          # jumlah file per sekali panggil LibreOffice
LO_TIMEOUT = 180       # detik per panggilan


class CertError(sl.SpreadsheetError):
    """Pesan aman & ramah untuk ditampilkan ke pengguna."""


def require_docx() -> None:
    if Document is None:
        raise CertError("Pustaka python-docx belum terpasang. Jalankan: pip install python-docx")


# ---------------------------------------------------------------------------
# 1. Validasi template
# ---------------------------------------------------------------------------
_INSTR_TEXT = re.compile(r"<w:instrText[^>]*>([^<]*)</w:instrText>")
_INSTR_ATTR = re.compile(r'w:instr="([^"]*)"')
_DANGEROUS_FIELD = re.compile(r"\b(INCLUDETEXT|INCLUDEPICTURE|DDEAUTO|DDE|LINK|IMPORT)\b", re.IGNORECASE)
_REL = re.compile(r"<Relationship\b[^>]*>")
_BENIGN_EXTERNAL = {"hyperlink", "attachedtemplate", "mailmergesource"}
_EXTERNAL_DESC = {
    "image": "gambar yang hanya tertaut ke file di luar dokumen (bukan tertanam)",
    "oleobject": "objek OLE tertaut ke file lain",
    "frame": "frame tertaut ke file lain",
    "subdocument": "sub-dokumen tertaut ke file lain",
}


def validate_template(filename: str, data: bytes) -> None:
    """Pastikan berkas adalah .docx biasa yang aman diproses. Melempar CertError jika tidak."""
    if sl.get_extension(filename) != "docx":
        raise CertError("Template harus berformat .docx (bukan .doc, .docm, atau .dotx). "
                        "Simpan ulang dari Word sebagai 'Dokumen Word (*.docx)'.")
    if not data:
        raise CertError("Berkas template kosong.")
    if len(data) > MAX_TEMPLATE_MB * 1024 * 1024:
        raise CertError(f"Ukuran template melebihi {MAX_TEMPLATE_MB} MB.")
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        raise CertError("Isi berkas bukan dokumen Word (.docx) yang valid. Berkas mungkin rusak, "
                        "diberi password, atau hanya diganti ekstensinya.")
    names = zf.namelist()
    if "word/document.xml" not in names or "[Content_Types].xml" not in names:
        raise CertError("Berkas ini bukan dokumen Word yang asli (struktur dokumen tidak sesuai).")
    if sum(i.file_size for i in zf.infolist()) > MAX_UNZIPPED_MB * 1024 * 1024:
        raise CertError("Isi template terlalu besar setelah diekstrak.")
    content_types = zf.read("[Content_Types].xml").decode("utf-8", "ignore")
    if "wordprocessingml.document.main+xml" not in content_types:
        raise CertError("Berkas ini bukan dokumen Word biasa (mungkin template .dotx atau berformat lain).")
    if any(n.lower().endswith("vbaproject.bin") for n in names):
        raise CertError("Template mengandung makro (VBA). Hapus makro lalu simpan sebagai .docx biasa.")

    # Field & relasi yang bisa menarik berkas lokal/jaringan ke dalam hasil (kebocoran data) ditolak.
    for n in names:
        if re.fullmatch(r"word/[^/]+\.xml", n):
            xml = zf.read(n).decode("utf-8", "ignore")
            instr = _INSTR_TEXT.findall(xml) + _INSTR_ATTR.findall(xml)
            if _DANGEROUS_FIELD.search(" ".join(instr)) or _DANGEROUS_FIELD.search("".join(instr)):
                raise CertError("Template memakai field yang mengambil berkas lain (INCLUDETEXT, INCLUDEPICTURE, "
                                "LINK, DDE, atau IMPORT). Hapus field tersebut demi keamanan.")
        elif re.fullmatch(r"word/_rels/[^/]+\.rels", n):
            for rel in _REL.findall(zf.read(n).decode("utf-8", "ignore")):
                if 'TargetMode="External"' not in rel:
                    continue
                typ = re.search(r'Type="([^"]*)"', rel)
                kind = typ.group(1).rsplit("/", 1)[-1] if typ else "?"
                if kind.lower() in _BENIGN_EXTERNAL:
                    continue            # hyperlink biasa; sumber data mail merge & template terlampir dibersihkan otomatis
                target = re.search(r'Target="([^"]*)"', rel)
                desc = _EXTERNAL_DESC.get(kind.lower(), f"objek tertaut berjenis '{kind}'")
                raise CertError(
                    f"Template memuat {desc} (bagian: {n.split('/')[-1].replace('.rels', '')}, "
                    f"tujuan: {(target.group(1) if target else '?')[:90]}). Sematkan isinya ke dalam dokumen: "
                    "di Word buka File → Info → Edit Links to Files → pilih tautan → Break Link, atau sisipkan "
                    "ulang gambar lewat Insert → Pictures (bukan 'Link to File'). Lalu simpan sebagai .docx.")


_MAILMERGE = re.compile(r"<w:mailMerge(?:\s[^>]*)?/>|<w:mailMerge(?:\s[^>]*)?>.*?</w:mailMerge>", re.DOTALL)
_ATTACHED = re.compile(r"<w:attachedTemplate\b[^>]*/>")
_DROP_REL = re.compile(r'<Relationship\b[^>]*Type="[^"]*/(?:attachedTemplate|mailMergeSource)"[^>]*/>')


def sanitize_template(data: bytes) -> bytes:
    """Buang sumber data Mail Merge & template terlampir (tautan ke file di PC pembuat template).

    Keduanya tidak dibutuhkan AUTOCert (data diambil dari Excel yang diunggah) dan bisa membuat Word menampilkan
    dialog koneksi data saat dokumen dibuka otomatis. Tag/field «Nama» tetap utuh. Idempoten.
    """
    try:
        zin = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        raise CertError("Isi berkas bukan dokumen Word (.docx) yang valid.")
    changed = False
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            blob = zin.read(info.filename)
            if info.filename == "word/settings.xml":
                text = blob.decode("utf-8")
                new = _ATTACHED.sub("", _MAILMERGE.sub("", text))
            elif info.filename == "word/_rels/settings.xml.rels":
                text = blob.decode("utf-8")
                new = _DROP_REL.sub("", text)
            else:
                zout.writestr(info.filename, blob)
                continue
            if new != text:
                changed = True
                blob = new.encode("utf-8")
            zout.writestr(info.filename, blob)
    return out.getvalue() if changed else data


# ---------------------------------------------------------------------------
# 2. Data peserta -> tabel teks
# ---------------------------------------------------------------------------
BULAN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus",
         "September", "Oktober", "November", "Desember"]
DATE_FORMATS = {
    "20-01-2025": "%d-%m-%Y",
    "20/01/2025": "%d/%m/%Y",
    "20 Januari 2025": "ID_LONG",
    "2025-01-20": "%Y-%m-%d",
}


def format_date(v, fmt: str) -> str:
    if fmt == "ID_LONG":
        return f"{v.day} {BULAN[v.month - 1]} {v.year}"
    return v.strftime(fmt)


def cell_to_text(v, date_format: str = "%d-%m-%Y") -> str:
    """Ubah nilai sel menjadi teks yang pantas dicetak di sertifikat (tanpa 'nan', '.0', atau jam 00:00:00)."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(v, (datetime, date)):          # pd.Timestamp adalah subclass datetime
        return format_date(v, date_format)
    if isinstance(v, (bool, np.bool_)):
        return "Ya" if v else "Tidak"
    if isinstance(v, (int, np.integer)):
        return str(int(v))
    if isinstance(v, (float, np.floating)):
        return str(int(v)) if float(v).is_integer() else np.format_float_positional(float(v), trim="-")
    return str(v).strip()


def participants_frame(raw: pd.DataFrame, header_row: Optional[int] = None,
                       date_format: str = "%d-%m-%Y") -> pd.DataFrame:
    """Tabel mentah (header=None) -> DataFrame berisi teks siap-cetak. Judul di atas tabel dilewati."""
    if raw.empty:
        return pd.DataFrame()
    if header_row is None:
        header_row = sl.guess_header_row(raw)
    header_row = max(0, min(int(header_row), len(raw) - 1))
    body = raw.iloc[header_row + 1:].copy()
    body.columns = sl.make_headers(raw.iloc[header_row].tolist())
    body = body.dropna(how="all", axis=0).dropna(how="all", axis=1).reset_index(drop=True)
    for c in body.columns:
        body[c] = body[c].map(lambda v: cell_to_text(v, date_format)).astype(object)
    return body


def guess_name_columns(columns: list[str]) -> list[str]:
    """Tebak kolom nama peserta untuk nama file PDF. Default: HANYA kolom nama (tanpa status/nomor).

    Pengguna tetap bisa menambah kolom lain lewat pilihan "Susun nama file" di layar.
    """
    for c in columns:                                   # utamakan judul kolom yang persis "Nama"
        if norm(c) in ("nama", "name", "nama lengkap", "nama peserta"):
            return [c]
    for c in columns:
        if re.search(r"nama|name", str(c), re.IGNORECASE):
            return [c]
    return list(columns[:1])


# ---------------------------------------------------------------------------
# 3. Nama file aman
# ---------------------------------------------------------------------------
_FORBIDDEN = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}


def safe_filename(parts: list[str], fallback: str, max_len: int = 120) -> str:
    name = "_".join(str(p).strip() for p in parts if p is not None and str(p).strip())
    name = unicodedata.normalize("NFKC", name)
    name = _FORBIDDEN.sub("_", name)
    name = re.sub(r"\s+", " ", name).strip(" .")
    name = name[:max_len].rstrip(" .")
    if not name or name.split(".")[0].upper() in _RESERVED:
        name = fallback
    return name


def unique_name(name: str, used: set) -> str:
    candidate, n = name, 2
    while candidate.lower() in used:
        candidate = f"{name} ({n})"
        n += 1
    used.add(candidate.lower())
    return candidate


# ---------------------------------------------------------------------------
# 4. Mail merge di dalam dokumen Word
# ---------------------------------------------------------------------------
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
_NS = {"w": W_NS}
_ALL_P = etree.XPath("//w:p", namespaces=_NS)       # termasuk paragraf di tabel & text box
_TEXT_NODES = etree.XPath(
    "./w:r/w:t | ./w:hyperlink/w:r/w:t | ./w:fldSimple/w:r/w:t | ./w:ins/w:r/w:t | ./w:smartTag/w:r/w:t"
    " | ./w:sdt/w:sdtContent/w:r/w:t", namespaces=_NS)
_PART_NAME = re.compile(r"/word/(document|header\d*|footer\d*)\.xml")
_BAD_XML = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")
_STRONG_TAG = re.compile(r"«\s*([^»\n]{1,80}?)\s*»|\{\{\s*([^{}\n]{1,80}?)\s*\}\}|<<\s*([^<>\n]{1,80}?)\s*>>")
_DELIMS = re.compile(r"^(?:«|\{\{|<<|\[)\s*|\s*(?:»|\}\}|>>|\])$")


def norm(s) -> str:
    """Normalisasi nama kolom/tag: abaikan huruf besar-kecil, spasi, dan garis bawah (MERGEFIELD memakai '_')."""
    return re.sub(r"[\s_]+", " ", str(s).strip()).lower()


def _roots(doc):
    for part in doc.part.package.iter_parts():
        if _PART_NAME.fullmatch(str(part.partname)) and hasattr(part, "_element"):
            yield part._element


def _clean_value(v) -> str:
    return re.sub(r"\s*[\r\n]+\s*", " ", _BAD_XML.sub("", str(v)))


def _merge_paragraph(p, regex: re.Pattern, values: dict) -> int:
    """Ganti tag di satu paragraf; format run tempat tag dimulai dipertahankan."""
    nodes = _TEXT_NODES(p)
    if not nodes:
        return 0
    texts = [n.text or "" for n in nodes]
    full = "".join(texts)
    matches = list(regex.finditer(full))
    if not matches:
        return 0
    starts = list(accumulate((len(t) for t in texts), initial=0))[:-1]
    touched: set[int] = set()
    count = 0
    for m in reversed(matches):                      # dari belakang agar posisi di depan tetap valid
        key = norm(_DELIMS.sub("", m.group(0)))
        if key not in values:
            continue
        s, e = m.start(), m.end()
        i0 = bisect_right(starts, s) - 1
        i1 = bisect_right(starts, e - 1) - 1
        o0, o1 = s - starts[i0], e - starts[i1]
        if i0 == i1:
            texts[i0] = texts[i0][:o0] + values[key] + texts[i0][o1:]
        else:
            texts[i0] = texts[i0][:o0] + values[key]
            for k in range(i0 + 1, i1):
                texts[k] = ""
            texts[i1] = texts[i1][o1:]
        touched.update(range(i0, i1 + 1))
        count += 1
    for i in touched:
        nodes[i].text = texts[i]
        nodes[i].set(XML_SPACE, "preserve")
    return count


class Merger:
    """Menyiapkan pola tag dari nama kolom, lalu menggabungkan template dengan satu baris data."""

    def __init__(self, columns):
        self.columns = [str(c) for c in columns]
        alts = []
        for c in sorted(self.columns, key=len, reverse=True):
            body = r"[\s_]+".join(re.escape(part) for part in c.split()) or re.escape(c)
            alts += [rf"«\s*{body}\s*»", rf"\{{\{{\s*{body}\s*\}}\}}", rf"<<\s*{body}\s*>>", rf"\[\s*{body}\s*\]"]
        self.regex = re.compile("|".join(alts), re.IGNORECASE) if alts else None

    def merge(self, template: bytes, row: dict) -> bytes:
        require_docx()
        values: dict[str, str] = {}
        for k, v in row.items():
            values.setdefault(norm(k), _clean_value(v))
        doc = Document(io.BytesIO(template))
        if self.regex is not None:
            for root in _roots(doc):
                for p in _ALL_P(root):
                    _merge_paragraph(p, self.regex, values)
        out = io.BytesIO()
        doc.save(out)
        return out.getvalue()


def _paragraph_texts(doc):
    for root in _roots(doc):
        for p in _ALL_P(root):
            text = "".join(n.text or "" for n in _TEXT_NODES(p))
            if text.strip():
                yield text


def extract_text(docx_bytes: bytes) -> list[str]:
    """Teks semua paragraf (termasuk tabel, text box, header/footer) — untuk pratinjau."""
    require_docx()
    return list(_paragraph_texts(Document(io.BytesIO(docx_bytes))))


def analyze_template(template: bytes, columns: list[str]) -> dict:
    """Bandingkan tag di template dengan kolom Excel.

    missing: tag «..»/{{..}}/<<..>> di template yang tidak punya kolom
    used:    kolom yang dipakai template (semua gaya tag, termasuk [Kolom])
    unused:  kolom yang tidak muncul di template
    """
    require_docx()
    doc = Document(io.BytesIO(template))
    merger = Merger(columns)
    col_norms = {norm(c) for c in columns}
    strong: dict[str, str] = {}
    used: set[str] = set()
    for text in _paragraph_texts(doc):
        for m in _STRONG_TAG.finditer(text):
            name = next(g for g in m.groups() if g)
            strong.setdefault(norm(name), name)
        if merger.regex is not None:
            for m in merger.regex.finditer(text):
                used.add(norm(_DELIMS.sub("", m.group(0))))
    return {
        "missing": [disp for n, disp in strong.items() if n not in col_norms],
        "used": [c for c in columns if norm(c) in used],
        "unused": [c for c in columns if norm(c) not in used],
    }


# ---------------------------------------------------------------------------
# 5. Mesin konversi PDF
# ---------------------------------------------------------------------------
def find_soffice() -> Optional[str]:
    for exe in ("soffice", "libreoffice"):
        found = shutil.which(exe)
        if found:
            return found
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
              "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if os.path.exists(p):
            return p
    return None


def word_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import comtypes.client  # noqa: F401
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application"))
        return True
    except Exception:
        return False


BACKEND_LABELS = {"word": "Microsoft Word (paling akurat)", "libreoffice": "LibreOffice"}


def available_backends() -> list[str]:
    found = []
    if word_available():
        found.append("word")
    if find_soffice():
        found.append("libreoffice")
    return found


def _convert_libreoffice(paths: list[Path], out_dir: Path) -> None:
    exe = find_soffice()
    if not exe:
        raise CertError("LibreOffice tidak ditemukan.")
    with tempfile.TemporaryDirectory(prefix="lo_profile_") as profile:   # profil terpisah: aman paralel
        cmd = [exe, f"-env:UserInstallation={Path(profile).as_uri()}", "--headless", "--norestore",
               "--nolockcheck", "--convert-to", "pdf", "--outdir", str(out_dir), *map(str, paths)]
        try:
            subprocess.run(cmd, capture_output=True, timeout=LO_TIMEOUT, check=False)
        except subprocess.TimeoutExpired:
            raise CertError("Konversi PDF melewati batas waktu. Kurangi jumlah peserta per proses.")


def _convert_word(pairs: list[tuple[Path, Path]], tick: Callable[[], None]) -> None:
    """Konversi lewat Microsoft Word (COM, hanya Windows). Hanya menutup instance Word milik proses ini."""
    import comtypes
    import comtypes.client
    comtypes.CoInitialize()
    app = None
    try:
        app = comtypes.client.CreateObject("Word.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        try:
            app.AutomationSecurity = 3          # msoAutomationSecurityForceDisable: matikan makro
        except Exception:
            pass
        for src, dst in pairs:
            for _ in range(3):
                try:
                    doc = app.Documents.Open(str(src.resolve()))
                    doc.SaveAs2(str(dst.resolve()), FileFormat=17)    # 17 = wdFormatPDF
                    doc.Close(SaveChanges=False)
                    break
                except Exception:
                    time.sleep(0.5)
            tick()
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        comtypes.CoUninitialize()


# ---------------------------------------------------------------------------
# 6. Pembuatan sertifikat massal
# ---------------------------------------------------------------------------
@dataclass
class GenResult:
    data: bytes                                   # isi ZIP
    ok: int
    fmt: str
    backend: str
    seconds: float
    failed: list[tuple[int, str, str]] = field(default_factory=list)   # (peserta ke-, nama file, alasan)


Progress = Callable[[int, int, str], None]


def generate(template: bytes, df: pd.DataFrame, name_cols: list[str], fmt: str = "pdf",
             backend: Optional[str] = None, progress: Optional[Progress] = None) -> GenResult:
    require_docx()
    template = sanitize_template(template)
    fmt = fmt.lower()
    if fmt not in ("pdf", "docx"):
        raise CertError("Format keluaran harus PDF atau DOCX.")
    if df.empty:
        raise CertError("Data peserta kosong.")
    if len(df) > MAX_PARTICIPANTS:
        raise CertError(f"Maksimal {MAX_PARTICIPANTS:,} peserta per proses. Pecah datanya menjadi beberapa berkas.")
    if fmt == "pdf":
        backends = available_backends()
        backend = backend if backend in backends else (backends[0] if backends else None)
        if backend is None:
            raise CertError("Tidak ada mesin PDF. Pasang Microsoft Word (Windows) atau LibreOffice, "
                            "atau pilih keluaran DOCX.")
    else:
        backend = "-"

    started = time.time()
    total = len(df) * (2 if fmt == "pdf" else 1)
    done = 0

    def tick(text: str) -> None:
        nonlocal done
        done += 1
        if progress:
            progress(min(done, total), total, text)

    merger = Merger(df.columns)
    used: set[str] = set()
    failed: list[tuple[int, str, str]] = []
    items: list[tuple[int, str, Path]] = []
    with tempfile.TemporaryDirectory(prefix="autocert_") as tmp_root:
        tmp = Path(tmp_root)
        src, out = tmp / "src", tmp / "out"
        src.mkdir()
        out.mkdir()
        for n, (_, row) in enumerate(df.iterrows()):
            values = {c: row[c] for c in df.columns}
            name = unique_name(safe_filename([values[c] for c in name_cols], f"sertifikat_{n + 1}"), used)
            path = src / f"{n:05d}.docx"                 # nama netral agar aman bagi mesin konversi
            try:
                path.write_bytes(merger.merge(template, values))
                items.append((n, name, path))
            except Exception as e:
                failed.append((n + 1, name, f"Gagal menyusun dokumen ({type(e).__name__})"))
            tick("Menyusun dokumen")

        if fmt == "pdf" and items:
            if backend == "word":
                _convert_word([(p, out / f"{n:05d}.pdf") for n, _, p in items], lambda: tick("Mengonversi ke PDF"))
            else:
                for i in range(0, len(items), LO_CHUNK):
                    chunk = items[i:i + LO_CHUNK]
                    _convert_libreoffice([p for _, _, p in chunk], out)
                    for _ in chunk:
                        tick("Mengonversi ke PDF")

        buf = io.BytesIO()
        ok = 0
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for n, name, path in items:
                if fmt == "docx":
                    zf.write(path, f"{name}.docx")
                    ok += 1
                else:
                    pdf = out / f"{n:05d}.pdf"
                    if pdf.exists() and pdf.stat().st_size > 0:
                        zf.write(pdf, f"{name}.pdf")
                        ok += 1
                    else:
                        failed.append((n + 1, name, "Konversi PDF gagal"))
    if ok == 0:
        raise CertError("Tidak ada sertifikat yang berhasil dibuat. " +
                        (failed[0][2] if failed else "Periksa template dan data."))
    failed.sort()
    return GenResult(buf.getvalue(), ok, fmt, backend, time.time() - started, failed)


def first_file(zip_bytes: bytes) -> tuple[str, bytes]:
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        name = zf.namelist()[0]
        return name, zf.read(name)
