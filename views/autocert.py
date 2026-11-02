"""
views/autocert.py — menu "AUTOCert": sertifikat massal dari template Word + data peserta (mail merge).

Alur: unggah template .docx  ->  unggah data peserta  ->  periksa kecocokan tag & kolom
      ->  coba 1 sertifikat  ->  generate semua  ->  unduh ZIP (PDF atau DOCX).
"""
import hashlib

import pandas as pd
import streamlit as st

import certlib
import sheetlib as sl
from views.common import DF_KW, esc_md, remembered_upload

MIME = {
    "pdf": "application/pdf",
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}


def show_df(data: pd.DataFrame) -> None:
    st.dataframe(data, hide_index=True, **DF_KW)


# ---------------------------------------------------------------------------
# Cache (hanya di memori server; kedaluwarsa otomatis)
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False, max_entries=3, ttl=1800)
def cached_validate_template(filename: str, data: bytes) -> bool:
    certlib.validate_template(filename, data)
    return True


@st.cache_data(show_spinner="Membaca data peserta…", max_entries=3, ttl=1800)
def cached_load(filename: str, data: bytes):
    return sl.load_workbook(sl.validate_upload(filename, data), data)


@st.cache_data(show_spinner=False, max_entries=6, ttl=1800)
def cached_frame(file_key, header_row, date_format, _raw):
    return certlib.participants_frame(_raw, header_row, date_format)


@st.cache_data(show_spinner=False, max_entries=3, ttl=1800)
def cached_analysis(template: bytes, columns: tuple):
    return certlib.analyze_template(template, list(columns))


# ---------------------------------------------------------------------------
def render_howto() -> None:
    with st.expander("📘 Cara pakai", expanded=True):
        st.markdown(
            "1. **Siapkan template Word (.docx)** dan tulis tag di tempat data peserta akan muncul, "
            "misalnya `«Nama»`, `{{Nama}}`, `<<Nama>>`, atau `[Nama]`. Tag boleh berada di paragraf, tabel, "
            "kotak teks (text box), header, maupun footer.\n"
            "2. **Siapkan data peserta** (Excel/CSV). Judul kolom harus sama dengan nama tag "
            "(huruf besar-kecil dan spasi/garis bawah tidak dibedakan: `Nomor Sertifikat` = `«Nomor_Sertifikat»`).\n"
            "3. Unggah keduanya di atas, periksa kecocokannya, **coba 1 sertifikat**, lalu **generate semua**.\n\n"
            "**Tips:** kolom nomor berawalan nol (mis. `00123`) sebaiknya berformat *Teks* di Excel. "
            "Tanggal otomatis ditulis sesuai format yang Anda pilih."
        )
    st.caption("🔒 Berkas hanya diproses di memori server selama sesi ini dan tidak disimpan permanen.")


def render() -> None:
    st.title("🎓 AUTOCert")
    st.caption("Otomatisasi sertifikat: template Word + data peserta (mail merge) → ZIP berisi PDF atau DOCX.")
    try:
        certlib.require_docx()
    except certlib.CertError as e:
        st.error(f"❌ {esc_md(e)}")
        return

    c1, c2 = st.columns(2)
    template = remembered_upload(
        "ac_tpl", "1️⃣ Template sertifikat (.docx)", ["docx"], container=c1,
        help="Tulis tag seperti «Nama», {{Nama}}, <<Nama>>, atau [Nama] sesuai judul kolom di Excel.",
    )
    participants = remembered_upload(
        "ac_data", "2️⃣ Data peserta (spreadsheet)", list(sl.ALLOWED_EXTENSIONS), container=c2,
        help="Format: " + ", ".join("." + e for e in sl.ALLOWED_EXTENSIONS),
    )
    if not template or not participants:
        render_howto()
        return
    tpl_name, tpl_bytes = template
    data_name, data_bytes = participants

    try:
        cached_validate_template(tpl_name, tpl_bytes)
    except sl.SpreadsheetError as e:
        st.error(f"❌ Template: {esc_md(e)}")
        return
    except Exception:
        st.error("❌ Template tidak dapat dibaca. Pastikan berkas tidak rusak.")
        return
    try:
        sheets, notes = cached_load(data_name, data_bytes)
    except sl.SpreadsheetError as e:
        st.error(f"❌ Data peserta: {esc_md(e)}")
        return
    except Exception:
        st.error("❌ Data peserta tidak dapat dibaca. Pastikan berkas tidak rusak.")
        return
    for note in notes:
        st.warning(esc_md(note))

    usable = [n for n, d in sheets.items() if not d.dropna(how="all").empty]
    if not usable:
        st.warning("Berkas data peserta tidak memiliki sheet yang berisi data.")
        return
    file_id = hashlib.md5(data_bytes).hexdigest()[:8]
    sheet = st.selectbox("Sheet data peserta", usable, key=f"ac_sheet_{file_id}") if len(usable) > 1 else usable[0]
    fk = f"{file_id}_{sheet}"
    raw = sheets[sheet]
    guess = sl.guess_header_row(raw)

    with st.expander("⚙️ Pengaturan data peserta"):
        header_row = guess
        if st.checkbox("Atur baris header manual", key=f"ac_man_{fk}"):
            header_row = int(st.number_input("Baris header (1 = baris pertama)", 1, len(raw), guess + 1,
                                             key=f"ac_hdr_{fk}")) - 1
        else:
            st.caption(f"Header terdeteksi otomatis di baris {guess + 1}.")
        date_labels = list(certlib.DATE_FORMATS)
        date_label = st.selectbox("Format tanggal di sertifikat", date_labels,
                                  index=date_labels.index("20 Januari 2025"), key=f"ac_date_{fk}")
    df = cached_frame(fk, header_row, certlib.DATE_FORMATS[date_label], raw)
    if df.empty:
        st.warning("Tidak ada data peserta setelah header. Coba atur baris header secara manual.")
        return
    if len(df) > certlib.MAX_PARTICIPANTS:
        st.error(f"Terdapat {len(df):,} peserta; maksimal {certlib.MAX_PARTICIPANTS:,} per proses. "
                 "Pecah datanya menjadi beberapa berkas.")
        return

    m = st.columns(3)
    m[0].metric("Jumlah peserta", f"{len(df):,}")
    m[1].metric("Kolom data", len(df.columns))
    m[2].metric("Sheet", esc_md(sheet))
    with st.expander("👀 Pratinjau data peserta", expanded=True):
        show_df(df.head(10))

    # --- Pemeriksaan kecocokan template <-> kolom ---------------------------------------------
    st.markdown("#### 🔗 Pemeriksaan template")
    try:
        info = cached_analysis(tpl_bytes, tuple(df.columns))
    except Exception:
        st.error("❌ Template tidak dapat dianalisis. Pastikan berkas adalah dokumen Word yang utuh.")
        return
    if not info["used"]:
        st.error("Tidak ada tag di template yang cocok dengan judul kolom data. Pastikan penulisan tag sama dengan "
                 "judul kolom, mis. kolom `Nama` → tag `«Nama»`.")
        return
    st.success("✅ Kolom yang dipakai template: " + ", ".join(esc_md(c) for c in info["used"]))
    if info["missing"]:
        st.warning("Tag di template yang **tidak punya kolom** di data (dibiarkan apa adanya): "
                   + ", ".join(f"«{esc_md(x)}»" for x in info["missing"]))
    if info["unused"]:
        st.caption("Kolom data yang tidak dipakai template: " + ", ".join(esc_md(c) for c in info["unused"]))

    # --- Pengaturan keluaran ---------------------------------------------------------------------
    st.markdown("#### 🧾 Pengaturan keluaran")
    cols = list(df.columns)
    backends = certlib.available_backends()
    name_cols = st.multiselect("Susun nama file (urut sesuai pilihan)", cols,
                               default=certlib.guess_name_columns(cols), key=f"ac_name_{fk}")
    fmt_label = st.radio("Format keluaran", ["PDF", "DOCX"] if backends else ["DOCX"],
                         horizontal=True, key=f"ac_fmt_{fk}")
    fmt = fmt_label.lower()
    backend = None
    if not backends:
        st.info("PDF membutuhkan Microsoft Word (Windows) atau LibreOffice di komputer/server ini. "
                "Saat ini hanya keluaran DOCX yang tersedia.")
    elif fmt == "pdf":
        if len(backends) > 1:
            backend = st.radio("Mesin konversi PDF", backends, format_func=certlib.BACKEND_LABELS.get,
                               horizontal=True, key=f"ac_backend_{fk}")
        else:
            backend = backends[0]
            st.caption(f"Mesin PDF: {certlib.BACKEND_LABELS[backend]}")
        if backend == "libreoffice":
            st.caption("ℹ️ Jika font template tidak terpasang di komputer/server ini, tata letak PDF bisa sedikit "
                       "berbeda dari tampilan di Word. Gunakan **Coba 1 sertifikat** untuk memeriksa.")
    example = certlib.safe_filename([df.iloc[0][c] for c in name_cols], "sertifikat_1")
    st.caption(f"Contoh nama file: {esc_md(example)}.{fmt}")
    if name_cols:
        stems = [certlib.safe_filename(list(v), "") for v in df[name_cols].itertuples(index=False, name=None)]
        stems = [x.lower() for x in stems if x]
        dupes = len(stems) - len(set(stems))
        if dupes:
            st.info(f"Ada {dupes:,} nama yang kembar. File-nya tetap dibuat dan diberi akhiran (2), (3), dst. "
                    "Tambahkan kolom lain (mis. nomor) di **Susun nama file** bila ingin membedakannya.")

    b1, b2, _ = st.columns([1, 1, 2])
    run_one = b1.button("🔍 Coba 1 sertifikat", key=f"ac_one_{fk}",
                        help="Buat sertifikat peserta pertama saja untuk memeriksa tata letak.")
    run_all = b2.button(f"🚀 Generate semua ({len(df):,})", type="primary", key=f"ac_all_{fk}")
    sample_key, result_key = f"ac_sample_{fk}", f"ac_result_{fk}"

    if run_one:
        try:
            with st.spinner("Membuat sampel…"):
                res = certlib.generate(tpl_bytes, df.head(1), name_cols, fmt, backend)
                fname, data = certlib.first_file(res.data)
                merged = certlib.Merger(df.columns).merge(tpl_bytes, {c: df.iloc[0][c] for c in cols})
                st.session_state[sample_key] = (fname, data, certlib.extract_text(merged))
        except sl.SpreadsheetError as e:
            st.error(f"❌ {esc_md(e)}")
        except Exception:
            st.error("❌ Gagal membuat sampel. Periksa template dan data peserta.")

    if run_all:
        bar = st.progress(0.0, text="Memulai…")

        def on_progress(done: int, total: int, text: str) -> None:
            bar.progress(done / total, text=f"{text} ({done}/{total})")

        try:
            st.session_state[result_key] = certlib.generate(tpl_bytes, df, name_cols, fmt, backend, on_progress)
        except sl.SpreadsheetError as e:
            st.error(f"❌ {esc_md(e)}")
        except Exception:
            st.error("❌ Terjadi kesalahan saat membuat sertifikat. Coba lagi atau kurangi jumlah peserta.")
        bar.empty()

    # --- Hasil (disimpan di session_state agar tombol unduh tetap ada setelah rerun) -------------
    sample = st.session_state.get(sample_key)
    if sample:
        fname, data, lines = sample
        st.markdown("#### 🔍 Sampel peserta pertama")
        st.download_button(f"📥 Unduh sampel ({fname.rsplit('.', 1)[-1].upper()})", data, file_name=fname,
                           mime=MIME.get(fname.rsplit(".", 1)[-1], "application/octet-stream"),
                           key=f"ac_dl_sample_{fk}")
        with st.expander("Teks sertifikat setelah digabung", expanded=True):
            st.text("\n".join(lines))

    result = st.session_state.get(result_key)
    if result:
        st.markdown("#### ✅ Hasil")
        r = st.columns(3)
        r[0].metric("Berhasil", f"{result.ok:,}")
        r[1].metric("Gagal", f"{len(result.failed):,}")
        r[2].metric("Waktu", f"{result.seconds:.0f} detik")
        if result.failed:
            st.warning("Sebagian sertifikat gagal dibuat:")
            show_df(pd.DataFrame(result.failed, columns=["Peserta ke-", "Nama file", "Alasan"]))
        stem = certlib.safe_filename([tpl_name.rsplit(".", 1)[0]], "sertifikat")
        d1, d2, _ = st.columns([1, 1, 2])
        d1.download_button("📥 Unduh semua (ZIP)", result.data, file_name=f"{stem}_{result.fmt}.zip",
                           mime="application/zip", key=f"ac_dl_all_{fk}")
        if d2.button("Hapus hasil", key=f"ac_clear_{fk}"):
            st.session_state.pop(result_key, None)
            st.rerun()
        st.caption("Hasil disimpan sementara di memori sesi ini. Klik **Hapus hasil** atau tutup tab untuk membuangnya.")
