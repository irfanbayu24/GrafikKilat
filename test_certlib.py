"""Tes logika AUTOCert:  python tests/test_certlib.py"""
import io, os, sys, zipfile
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))

import pandas as pd
from docx import Document
from docx.shared import Pt
from lxml import etree
import certlib as cl

fails = 0
def check(name, cond, extra=""):
    global fails
    print(("  ✔ " if cond else "  ✘ GAGAL ") + name, "" if cond else extra)
    fails += (not cond)

def expect_error(label, fn):
    try:
        fn(); check(f"tolak: {label}", False, "(tidak melempar error)")
    except cl.CertError as e:
        check(f"tolak: {label}", True); print("       ->", str(e)[:80])

# ------------------------------------------------------------------ template uji
def make_template() -> bytes:
    d = Document()
    p = d.add_paragraph()
    p.add_run("Diberikan kepada ")
    r = p.add_run("«Na"); r.bold = True; r.font.size = Pt(30)      # tag terpecah antar-run
    r2 = p.add_run("ma»"); r2.italic = True
    d.add_paragraph("Nomor: {{Nomor_Sertifikat}} | Status: [status] | Tgl: <<Tanggal>>")
    d.add_paragraph("«Nama» sebagai «Status» ( «Tidak Ada» )")
    t = d.add_table(rows=1, cols=1); t.rows[0].cells[0].paragraphs[0].add_run("Sel tabel: «Nama»")
    sec = d.sections[0]; sec.header.is_linked_to_previous = False
    sec.header.paragraphs[0].text = "Header untuk «Nama»"
    # text box (VML) berisi tag
    xml = ('<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
           'xmlns:v="urn:schemas-microsoft-com:vml"><w:pict><v:shape><v:textbox><w:txbxContent>'
           '<w:p><w:r><w:t>Kotak: «Nama»</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></w:r>')
    d.add_paragraph()._p.append(etree.fromstring(xml))
    out = io.BytesIO(); d.save(out); return out.getvalue()

TPL = make_template()
cols = ["Nama", "Status", "Nomor Sertifikat", "Tanggal"]
row = {"Nama": "Budi & <Santoso>", "Status": "Peserta", "Nomor Sertifikat": "001/CERT/2025", "Tanggal": "20 Januari 2025"}

print("== 1. Validasi template")
cl.validate_template("t.docx", TPL); check("template sah diterima", True)
def tweak(extra=None, rels=None, ctype=None, drop=None):
    zin = zipfile.ZipFile(io.BytesIO(TPL)); out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for i in zin.infolist():
            data = zin.read(i.filename)
            if i.filename == "word/document.xml" and extra: data = data.replace(b"</w:body>", extra + b"</w:body>")
            if i.filename == "word/_rels/document.xml.rels" and rels: data = data.replace(b"</Relationships>", rels + b"</Relationships>")
            if i.filename == "[Content_Types].xml" and ctype: data = data.replace(*ctype)
            z.writestr(i.filename, data)
        if drop: z.writestr(drop, b"x")
    return out.getvalue()
expect_error("ekstensi .doc", lambda: cl.validate_template("t.doc", TPL))
expect_error("ekstensi .docm", lambda: cl.validate_template("t.docm", TPL))
expect_error("bukan zip", lambda: cl.validate_template("t.docx", b"%PDF-1.4 bukan docx"))
expect_error("kosong", lambda: cl.validate_template("t.docx", b""))
buf = io.BytesIO()
with zipfile.ZipFile(buf, "w") as z: z.writestr("a.txt", "x")
expect_error("zip biasa", lambda: cl.validate_template("t.docx", buf.getvalue()))
expect_error("mengandung makro", lambda: cl.validate_template("t.docx", tweak(drop="word/vbaProject.bin")))
fld = (b'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:instrText>'
       b' INCLUDETEXT "C:\\rahasia.txt" </w:instrText></w:r></w:p>')
expect_error("field INCLUDETEXT", lambda: cl.validate_template("t.docx", tweak(extra=fld)))
expect_error("relasi eksternal (gambar remote)", lambda: cl.validate_template("t.docx", tweak(
    rels=b'<Relationship Id="rX" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="http://evil/x.png" TargetMode="External"/>')))
cl.validate_template("t.docx", tweak(rels=b'<Relationship Id="rY" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://contoh.id" TargetMode="External"/>'))
check("hyperlink biasa tetap boleh", True)
expect_error("terlalu besar", lambda: cl.validate_template("t.docx", b"PK" + b"0" * (cl.MAX_TEMPLATE_MB * 1024 * 1024 + 1)))

print("== 2. Mail merge")
m = cl.Merger(cols)
out = m.merge(TPL, row)
texts = cl.extract_text(out); blob = "\n".join(texts); print("   " + blob.replace("\n", "\n   "))
check("tag terpecah antar-run terganti", "Diberikan kepada Budi & <Santoso>" in blob)
check("gaya {{ }} + garis bawah + [ ] + << >>", "Nomor: 001/CERT/2025 | Status: Peserta | Tgl: 20 Januari 2025" in blob)
check("beberapa tag dalam satu paragraf", "Budi & <Santoso> sebagai Peserta" in blob)
check("tag tak dikenal dibiarkan", "«Tidak Ada»" in blob)
check("tabel", "Sel tabel: Budi & <Santoso>" in blob)
check("header", "Header untuk Budi & <Santoso>" in blob)
check("text box", "Kotak: Budi & <Santoso>" in blob)
check("tidak ada sisa tag yang dikenal", not any(t in blob for t in ("«Nama»", "{{Nomor", "[status]", "<<Tanggal>>")))
d2 = Document(io.BytesIO(out))   # hasil harus tetap dokumen Word yang valid
r = d2.paragraphs[0].runs
check("format run pertama tag dipertahankan (tebal)", any(x.bold and "Budi" in x.text for x in r), [(x.text, x.bold) for x in r])
inj = m.merge(TPL, {**row, "Nama": "«Nomor Sertifikat»", "Status": "x\ny\x00z"})
t2 = "\n".join(cl.extract_text(inj))
check("nilai berisi tag tidak diproses ulang", "Diberikan kepada «Nomor Sertifikat»" in t2)
check("karakter kontrol XML & newline dibersihkan", "x y" in t2.replace("\x00", "") and "\x00" not in t2)
ci = cl.Merger(["nama"]).merge(TPL, {"nama": "Ani"})
check("huruf besar-kecil tak dibedakan", "Ani" in "\n".join(cl.extract_text(ci)))

print("== 3. Analisis template")
a = cl.analyze_template(TPL, cols)
check("tag tanpa kolom terdeteksi", a["missing"] == ["Tidak Ada"], a)
check("semua kolom terpakai", a["unused"] == [] and set(a["used"]) == set(cols), a)
a2 = cl.analyze_template(TPL, cols + ["Alamat"]); check("kolom tak terpakai terdeteksi", a2["unused"] == ["Alamat"])

print("== 4. Nama file aman")
sf = cl.safe_filename
check("path traversal", "/" not in sf(["../../etc/passwd"], "f") and "\\" not in sf(["..\\..\\x"], "f"), sf(["../../etc/passwd"], "f"))
check("nama cadangan jika hanya titik", sf(["..", "  "], "sertifikat_1") == "sertifikat_1")
check("nama reserved Windows", sf(["CON"], "fb") == "fb" and sf(["nul"], "fb") == "fb")
check("karakter terlarang", sf(['a:b*c?"d<e>f|g'], "fb") == "a_b_c__d_e_f_g", sf(['a:b*c?"d<e>f|g'], "fb"))
check("dibatasi panjangnya", len(sf(["x" * 500], "fb")) <= 120)
check("gabungan nama_status_nomor", sf(["Budi", "Peserta", "001/CERT"], "fb") == "Budi_Peserta_001_CERT")
u = set(); names = [cl.unique_name("Budi", u) for _ in range(3)] + [cl.unique_name("BUDI", u)]
check("duplikat diberi nomor", names == ["Budi", "Budi (2)", "Budi (3)", "BUDI (4)"], names)

print("== 5. Data peserta (xlsx nyata)")
import sheetlib as sl
df_x = pd.DataFrame({"Nama": ["Ani", "Budi"], "Tanggal": pd.to_datetime(["2025-01-20", "2025-02-03"]),
                     "Skor": [90.0, 85.5], "Nomor": ["00123", "00456"], "Jumlah": [3, 4], "Kosong": [None, "ada"]})
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as w:
    df_x.to_excel(w, index=False, startrow=2, sheet_name="Peserta"); w.sheets["Peserta"]["A1"] = "Daftar Peserta"
sheets, _ = sl.load_workbook(sl.validate_upload("p.xlsx", buf.getvalue()), buf.getvalue())
pf = cl.participants_frame(sheets["Peserta"], None, "ID_LONG"); print(pf.to_string(index=False))
check("judul dilewati, header benar", list(pf.columns) == list(df_x.columns))
check("tanggal format Indonesia", pf["Tanggal"].tolist() == ["20 Januari 2025", "3 Februari 2025"], pf["Tanggal"].tolist())
check("90.0 -> 90 ; 85.5 tetap", pf["Skor"].tolist() == ["90", "85.5"], pf["Skor"].tolist())
check("angka nol di depan (teks) terjaga", pf["Nomor"].tolist() == ["00123", "00456"])
check("sel kosong -> '' (bukan 'nan')", pf["Kosong"].tolist() == ["", "ada"], pf["Kosong"].tolist())
check("default nama file = kolom nama saja", cl.guess_name_columns(["Nomor Sertifikat", "Status", "Nama", "Email"]) == ["Nama"])
check("utamakan 'Nama' persis daripada 'Nama Instansi'", cl.guess_name_columns(["Nama Instansi", "Nama"]) == ["Nama"])
check("tanpa kolom nama -> kolom pertama", cl.guess_name_columns(["Peserta", "Status"]) == ["Peserta"])

print("== 6. Generate massal")
people = pd.DataFrame([{"Nama": n, "Status": s, "Nomor Sertifikat": f"00{i}/CERT/2025", "Tanggal": "20 Januari 2025"}
                       for i, (n, s) in enumerate([("Ani", "Peserta"), ("Budi", "Pembicara"), ("Budi", "Pembicara"), ("../../x", "Panitia"), ("", "")], 1)])
res = cl.generate(TPL, people, ["Nama", "Status", "Nomor Sertifikat"], "docx")
z = zipfile.ZipFile(io.BytesIO(res.data)); names = z.namelist(); print("   ", names)
check("5 file DOCX, nama aman & unik", res.ok == 5 and len(set(n.lower() for n in names)) == 5 and not any("/" in n for n in names), names)
check("isi tiap DOCX tergabung", "Ani" in "\n".join(cl.extract_text(z.read(names[0]))))
expect_error("melebihi batas peserta", lambda: cl.generate(TPL, pd.concat([people] * 300), ["Nama"], "docx"))
expect_error("data kosong", lambda: cl.generate(TPL, people.iloc[0:0], ["Nama"], "docx"))

print("== 7. PDF (LibreOffice bila tersedia)")
bk = cl.available_backends(); print("    mesin PDF:", bk)
if "libreoffice" in bk:
    res = cl.generate(TPL, people.head(3), ["Nama", "Nomor Sertifikat"], "pdf", "libreoffice")
    z = zipfile.ZipFile(io.BytesIO(res.data)); names = z.namelist(); print("   ", names, f"{res.seconds:.1f}s")
    check("3 PDF valid", res.ok == 3 and all(z.read(n)[:5] == b"%PDF-" for n in names) and not res.failed, res.failed)
    open("/tmp/sample_cert.pdf", "wb").write(z.read(names[0]))
else:
    print("    (dilewati: LibreOffice tidak ada)")
_orig = cl.available_backends
cl.available_backends = lambda: []
expect_error("PDF tanpa mesin yang tersedia", lambda: cl.generate(TPL, people.head(1), ["Nama"], "pdf"))
cl.available_backends = _orig

print("\nHASIL:", "SEMUA PENGUJIAN LULUS" if fails == 0 else f"{fails} GAGAL")
sys.exit(1 if fails else 0)
