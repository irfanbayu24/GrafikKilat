"""
views/cek_absensi.py — menu "Cek Absensi": bandingkan dua daftar lewat kolom kunci (NIP/NIK/ID/email/nama).

File 1 = daftar acuan (mis. semua pegawai/peserta)   File 2 = daftar pembanding (mis. yang hadir)
Hasil: sudah ada · belum ada · hanya di File 2 · perlu dicek (kunci kosong/kembar) + laporan Excel.
"""
import hashlib

import pandas as pd
import streamlit as st

import checklib
import matchlib
import sheetlib as sl
from views.common import DF_KW, esc_md, remembered_upload

PREVIEW_ROWS = 1000


def show_df(data: pd.DataFrame) -> None:
    st.dataframe(data, hide_index=True, **DF_KW)


# ---------------------------------------------------------------------------
# Cache (hanya di memori server; kedaluwarsa otomatis)
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Membaca berkas…", max_entries=4, ttl=1800)
def cached_load(filename: str, data: bytes):
    return sl.load_workbook(sl.validate_upload(filename, data), data)


@st.cache_data(show_spinner=False, max_entries=8, ttl=1800)
def cached_frame(file_key, header_row, drop_summary, _raw):
    return sl.build_dataframe(_raw, header_row, False, drop_summary)


@st.cache_data(show_spinner="Membandingkan data…", max_entries=4, ttl=1800)
def cached_compare(sig, key1, key2, ignore_case, ignore_spaces, digits_only, strip_zeros, extras, _df1, _df2):
    opts = checklib.Options(ignore_case, ignore_spaces, digits_only, strip_zeros)
    return checklib.compare(_df1, key1, _df2, key2, opts, tuple(extras))


@st.cache_data(show_spinner="Menyiapkan pencocok…", max_entries=4, ttl=1800)
def cached_matcher(sig, unit_col, name_col, noise_text, overlap, fallback, max_extra, typo, _df1):
    return matchlib.Matcher(_df1, unit_col, name_col, noise=matchlib.parse_noise(noise_text), unit_overlap=overlap,
                            name_only_fallback=fallback, max_extra=max_extra, typo=typo)


@st.cache_data(show_spinner="Mencocokkan nama…", max_entries=4, ttl=1800)
def cached_flex_run(sig, text_col, unit_col, name_col, _matcher, _df1, _df2):
    return matchlib.run(_matcher, _df1, _df2, text_col, unit_col, name_col)


@st.cache_data(show_spinner=False, max_entries=8, ttl=900)
def cached_csv(sig, name, _df) -> bytes:
    return sl.to_csv_bytes(_df)


@st.cache_data(show_spinner="Menyusun laporan Excel…", max_entries=2, ttl=900)
def cached_report(sig, _res) -> bytes:
    return sl.to_excel_sheets(checklib.build_report(_res))


# ---------------------------------------------------------------------------
def file_panel(slot: str, title: str, box):
    """Unggah + pilih sheet + muat satu file. Mengembalikan dict atau None."""
    mem = remembered_upload(slot, title, list(sl.ALLOWED_EXTENSIONS), container=box,
                            help="Format: " + ", ".join("." + e for e in sl.ALLOWED_EXTENSIONS))
    if not mem:
        return None
    name, data = mem
    try:
        sheets, notes = cached_load(name, data)
    except sl.SpreadsheetError as e:
        box.error(f"❌ {esc_md(e)}")
        return None
    except Exception:
        box.error("❌ Berkas tidak dapat dibaca. Pastikan berkas tidak rusak.")
        return None
    for note in notes:
        box.warning(esc_md(note))
    usable = [n for n, d in sheets.items() if not d.dropna(how="all").empty]
    if not usable:
        box.warning("Berkas ini tidak memiliki sheet yang berisi data.")
        return None
    file_id = hashlib.md5(data).hexdigest()[:8]
    with box:
        sheet = st.selectbox("Sheet", usable, key=f"{slot}_sheet_{file_id}") if len(usable) > 1 else usable[0]
        fk = f"{file_id}_{sheet}"
        raw = sheets[sheet]
        guess = sl.guess_header_row(raw)
        header_row = guess
        if st.checkbox("Atur baris header manual", key=f"{slot}_man_{fk}"):
            header_row = int(st.number_input("Baris header (1 = baris pertama)", 1, len(raw), guess + 1,
                                             key=f"{slot}_hdr_{fk}")) - 1
        else:
            st.caption(f"Header terdeteksi di baris {guess + 1}.")
        drop_summary = st.checkbox("Buang baris Total/Jumlah", True, key=f"{slot}_sum_{fk}")
    df, removed = cached_frame(fk, header_row, drop_summary, raw)
    if df.empty:
        box.warning("Tidak ada data setelah header. Coba atur baris header secara manual.")
        return None
    box.caption(f"📄 {esc_md(name)} · {len(df):,} baris · {len(df.columns)} kolom"
                + (f" · {removed} baris total dibuang" if removed else ""))
    return {"name": name, "fk": fk, "df": df, "header": header_row, "drop": drop_summary}


MODE_EXACT = "Satu kolom kunci (persis) — NIP/NIK/ID/email"
MODE_FLEX = "Unit + Nama → satu kolom teks di File 2 (fleksibel, nama boleh tidak lengkap)"
NO_UNIT = "(Tidak ada — cocokkan nama saja)"
_CONF_RANK = {"Rendah": 0, "Sedang": 1, "Tinggi": 2}


def _guess_col(cols: list, pattern: str, avoid=None):
    import re
    for c in cols:
        if c != avoid and re.search(pattern, str(c), re.IGNORECASE):
            return c
    return None


def flex_tab(df: pd.DataFrame, sig: str, name: str, empty_msg: str, sort_low_conf_first: bool = False) -> None:
    if df.empty:
        st.success(empty_msg)
        return
    if sort_low_conf_first and "Keyakinan" in df.columns[-1:]:
        df = df.sort_values(df.columns[-1], key=lambda s: s.map(_CONF_RANK).fillna(0), kind="stable")
    if len(df) > PREVIEW_ROWS:
        st.caption(f"Menampilkan {PREVIEW_ROWS:,} dari {len(df):,} baris. Unduh untuk melihat semuanya.")
    show_df(df.head(PREVIEW_ROWS))
    st.download_button("📥 Unduh CSV", cached_csv(sig, name, df), file_name=f"{name}.csv", mime="text/csv",
                       key=f"ck_flex_csv_{name}_{sig[:8]}")


def render_flexible(p1: dict, p2: dict) -> None:
    df1, df2 = p1["df"], p2["df"]
    cols1, cols2 = list(df1.columns), list(df2.columns)

    st.markdown("#### 🔑 Kolom yang dipakai")
    g_unit = _guess_col(cols1, r"unit|satker|satuan kerja|bagian|instansi|divisi|bidang|biro|seksi|departemen|dinas")
    g_name = _guess_col(cols1, r"nama|name|peserta", avoid=g_unit) or next((c for c in cols1 if c != g_unit), cols1[0])
    g_text = _guess_col(cols2, r"nama|name|peserta|participant|display") or cols2[0]
    a, b, c = st.columns(3)
    unit_opts = [NO_UNIT] + cols1
    unit_pick = a.selectbox("File 1 — kolom UNIT", unit_opts, index=unit_opts.index(g_unit) if g_unit else 0,
                            key=f"ck_f_unit_{p1['fk']}")
    name_col = b.selectbox("File 1 — kolom NAMA", cols1, index=cols1.index(g_name), key=f"ck_f_name_{p1['fk']}")
    text_col = c.selectbox("File 2 — kolom teks (mis. 'Unit - Nama')", cols2, index=cols2.index(g_text),
                           key=f"ck_f_text_{p2['fk']}")
    unit_col = None if unit_pick == NO_UNIT else unit_pick
    if unit_col == name_col:
        st.warning("Kolom Unit dan Nama tidak boleh sama.")
        return
    samples = df2[text_col].map(checklib.key_text).loc[lambda x: x != ""].head(3).tolist()
    if samples:
        st.caption("Contoh teks File 2 yang dibaca: " + " · ".join(esc_md(v) for v in samples))

    with st.expander("⚙️ Aturan pencocokan"):
        o1, o2 = st.columns(2)
        unit_rule = o1.selectbox("Pencocokan unit", ["Ketat — semua kata unit harus ada di teks",
                                                     "Longgar — cukup separuh kata unit"], key="ck_f_unitrule",
                                 disabled=unit_col is None)
        extra_rule = o2.selectbox("Kata di File 2 yang tidak dikenali", ["Tidak boleh ada (ketat)",
                                                                         "Boleh 1 kata (longgar)"], key="ck_f_extra")
        fallback = o1.checkbox("Jika unit tidak terbaca di File 2, coba cocokkan nama saja (keyakinan rendah)", False,
                               key="ck_f_fallback", disabled=unit_col is None,
                               help="Berguna jika sebagian peserta menulis singkatan unit yang tidak dikenal. "
                                    "Nama lengkap di File 1 harus ikut tercocokkan seluruhnya.")
        typo = o2.checkbox("Toleransi salah ketik 1 huruf (kata ≥ 5 huruf)", True, key="ck_f_typo")
        noise_text = st.text_area("Kata yang diabaikan (sapaan, gelar, label perangkat) — pisahkan dengan spasi",
                                  matchlib.DEFAULT_NOISE, height=110, key="ck_f_noise")
    overlap = 0.5 if unit_rule.startswith("Longgar") else 1.0
    max_extra = 1 if extra_rule.startswith("Boleh") else 0

    sig_m = hashlib.md5(repr((p1["fk"], p1["header"], p1["drop"], unit_col, name_col, noise_text, overlap, fallback,
                              max_extra, typo)).encode()).hexdigest()
    try:
        matcher = cached_matcher(sig_m, unit_col, name_col, noise_text, overlap, fallback, max_extra, typo, df1)
        sig = hashlib.md5(repr((sig_m, p2["fk"], p2["header"], p2["drop"], text_col)).encode()).hexdigest()
        res = cached_flex_run(sig, text_col, unit_col, name_col, matcher, df1, df2)
    except Exception:
        st.error("❌ Terjadi kesalahan saat mencocokkan nama. Periksa pilihan kolom lalu coba lagi.")
        return
    s = res.summary

    st.markdown("#### 📊 Ringkasan")
    m = st.columns(4)
    m[0].metric("Baris File 1", f"{s['total_1']:,}")
    m[1].metric("✅ Hadir", f"{s['hadir']:,}")
    m[2].metric("❓ Hadir (belum pasti)", f"{s['belum_pasti']:,}")
    m[3].metric("❌ Belum hadir", f"{s['belum_hadir']:,}")
    st.progress(min(s["pct_hadir"] / 100, 1.0), text=f"{s['pct_hadir']}% hadir (pasti) dari baris File 1 yang punya nama")
    st.caption(f"Entri File 2: {s['total_2']:,} · cocok {s['cocok']:,} · ambigu {s['ambigu']:,} · tidak cocok "
               f"{s['tidak_cocok']:,} · kosong {s['kosong']:,}"
               + (f" · ⚠️ {s['blank_1']:,} baris File 1 tanpa nama tidak dicocokkan" if s["blank_1"] else ""))
    if s["belum_pasti"] or s["ambigu"]:
        st.info("❓ **Belum pasti** = nama di File 2 cocok dengan lebih dari satu orang di File 1 (mis. dua 'Budi' di unit yang sama), "
                "jadi tidak ditebak. Lihat tab **Rincian** dan lengkapi nama di File 2 bila perlu.")

    with st.expander("🔍 Uji satu teks dari File 2 (lihat cara aplikasi membacanya)"):
        probe = st.text_input("Ketik atau tempel teks, mis. BKT - Budi", key="ck_f_probe")
        if probe.strip():
            oc = matcher.match_text(probe)
            st.markdown(f"**Status:** {oc.status}" + (f" · **Keyakinan:** {oc.confidence}" if oc.confidence else ""))
            if oc.tokens:
                st.caption("Kata yang terbaca: " + ", ".join(esc_md(t) for t in oc.tokens))
            if oc.note:
                st.caption("Catatan: " + esc_md(oc.note))
            if oc.ranked:
                rows = []
                for pi, key in oc.ranked:
                    person = matcher.persons[pi]
                    rows.append({"Unit": checklib.key_text(df1[unit_col].iloc[person.row]) if unit_col else "",
                                 "Nama": checklib.key_text(df1[name_col].iloc[person.row]),
                                 "Skor": f"unit {key[0]:.0%} · {'nama persis' if key[1] else 'nama sebagian'} · kata asing {-key[2]}"})
                show_df(pd.DataFrame(rows))

    t1, t2, t3, t4, t5 = st.tabs([f"✅ Hadir ({s['hadir']:,})", f"❓ Belum pasti ({s['belum_pasti']:,})",
                                  f"❌ Belum hadir ({s['belum_hadir']:,})", f"➕ Tidak cocok ({s['tidak_cocok']:,})",
                                  "🔎 Rincian"])
    with t1:
        st.caption("Diurutkan dari keyakinan terendah agar mudah diperiksa. **Tinggi** = unit dan nama persis; "
                   "**Sedang** = nama sebagian; **Rendah** = unit tak terbaca atau ada kata asing.")
        flex_tab(res.hadir, sig, "hadir", "Belum ada yang tercatat hadir.", sort_low_conf_first=True)
    with t2:
        flex_tab(res.belum_pasti, sig, "belum_pasti", "Tidak ada yang ambigu.")
    with t3:
        flex_tab(res.belum_hadir, sig, "belum_hadir", "🎉 Semua baris File 1 sudah tercatat hadir.")
    with t4:
        st.caption("Entri File 2 yang tidak cocok dengan siapa pun di File 1 (tamu, salah ketik berat, atau unit tak dikenal).")
        flex_tab(res.only_in_2, sig, "tidak_cocok", "Semua entri File 2 berhasil dicocokkan.")
    with t5:
        st.caption("Satu baris per entri File 2: dicocokkan ke siapa, dengan keyakinan apa. Gunakan untuk memeriksa hasil.")
        pick = st.multiselect("Filter status", ["Cocok", "Ambigu", "Tidak cocok", "Kosong"], default=["Ambigu", "Tidak cocok"],
                              key="ck_f_detail_filter")
        flex_tab(res.detail[res.detail["Status"].isin(pick)].reset_index(drop=True), sig, "rincian",
                 "Tidak ada entri dengan status yang dipilih.")

    st.markdown("#### 📗 Laporan lengkap")
    if st.button("Siapkan laporan Excel", key="ck_f_report_btn"):
        st.session_state["ck_f_report"] = (sig, sl.to_excel_sheets(matchlib.build_report(res)))
    ready = st.session_state.get("ck_f_report")
    if ready and ready[0] == sig:
        st.download_button("📥 Unduh laporan (.xlsx)", ready[1], file_name="laporan_cek_absensi_fleksibel.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="ck_f_report_dl")


def render_howto() -> None:
    with st.expander("📘 Cara pakai", expanded=True):
        st.markdown(
            "1. **File 1** = daftar acuan, mis. semua pegawai/peserta yang seharusnya hadir.\n"
            "2. **File 2** = daftar pembanding, mis. daftar hadir (rekap Zoom, absensi, dll.).\n"
            "3. Pilih **kolom kunci** di masing-masing file (NIP, NIK, ID, email, atau nama). Nama kolomnya boleh berbeda.\n"
            "4. Hasil muncul otomatis: **sudah ada** (hadir), **belum ada** (belum hadir), **hanya di File 2** "
            "(hadir tapi tidak terdaftar), dan **perlu dicek** (kunci kosong/kembar). Semuanya bisa diunduh dalam satu laporan Excel.\n\n"
            "5. **Mode fleksibel:** jika File 1 punya kolom **Unit** dan **Nama** terpisah sedangkan File 2 hanya satu kolom teks "
            "seperti `BKT - Budi` (nama boleh tidak lengkap), pilih *Cara mencocokkan → fleksibel*.\n\n"
            "**Tips:** kolom NIP/NIK (16–18 digit) harus berformat *Teks* di Excel. Jika berformat Angka, Excel menyimpan "
            "hanya 15 digit sehingga digit belakangnya berubah menjadi nol."
        )
    st.caption("🔒 Berkas hanya diproses di memori server selama sesi ini dan tidak disimpan permanen.")


def table_tab(df: pd.DataFrame, sig: str, name: str, fname: str, empty_msg: str) -> None:
    if df.empty:
        st.success(empty_msg)
        return
    if len(df) > PREVIEW_ROWS:
        st.caption(f"Menampilkan {PREVIEW_ROWS:,} dari {len(df):,} baris. Unduh untuk melihat semuanya.")
    show_df(df.head(PREVIEW_ROWS))
    st.download_button("📥 Unduh CSV", cached_csv(sig, name, df), file_name=fname, mime="text/csv",
                       key=f"ck_csv_{name}_{sig[:8]}")


def render() -> None:
    st.title("✅ Cek Absensi")
    st.caption("Bandingkan dua daftar lewat kolom kunci: siapa yang sudah ada, belum ada, atau hanya ada di salah satu file.")

    c1, c2 = st.columns(2)
    p1 = file_panel("ck_1", "1️⃣ File 1 — daftar acuan (mis. semua pegawai/peserta)", c1)
    p2 = file_panel("ck_2", "2️⃣ File 2 — daftar pembanding (mis. yang hadir)", c2)
    if not (p1 and p2):
        render_howto()
        return

    mode = st.radio("Cara mencocokkan", [MODE_EXACT, MODE_FLEX], horizontal=True, key="ck_mode",
                    help="Persis: satu kolom kunci yang sama di kedua file. Fleksibel: File 1 punya kolom Unit dan Nama "
                         "terpisah, File 2 hanya satu kolom teks seperti 'Unit - Nama' (nama boleh tidak lengkap).")
    if mode == MODE_FLEX:
        render_flexible(p1, p2)
    else:
        render_exact(p1, p2)


def render_exact(p1: dict, p2: dict) -> None:
    df1, df2 = p1["df"], p2["df"]
    cols1, cols2 = list(df1.columns), list(df2.columns)

    st.markdown("#### 🔑 Kolom kunci")
    g1, g2 = checklib.guess_key_columns(cols1, cols2)
    k1c, k2c = st.columns(2)
    key1 = k1c.selectbox("Kolom kunci di File 1", cols1, index=cols1.index(g1), key=f"ck_k1_{p1['fk']}")
    key2 = k2c.selectbox("Kolom kunci di File 2", cols2, index=cols2.index(g2), key=f"ck_k2_{p2['fk']}")

    with st.expander("⚙️ Aturan pencocokan & kolom tambahan"):
        o1, o2 = st.columns(2)
        ignore_case = o1.checkbox("Abaikan huruf besar/kecil", True, key="ck_o_case")
        ignore_spaces = o1.checkbox("Abaikan spasi", True, key="ck_o_space")
        digits_only = o2.checkbox("Hanya angka (mis. NIP/NIK yang ditulis dengan pemisah)", False, key="ck_o_digits")
        strip_zeros = o2.checkbox("Abaikan angka nol di depan (00123 = 123)", False, key="ck_o_zero")
        extras = st.multiselect("Tambahkan kolom dari File 2 ke hasil 'Sudah ada'",
                                [c for c in cols2 if c != key2], key=f"ck_extra_{p2['fk']}_{key2}")

    n_bad = max(checklib.precision_warning_count(df1[key1]), checklib.precision_warning_count(df2[key2]))
    if n_bad:
        st.warning(f"⚠️ Ada {n_bad:,} nilai pada kolom kunci berupa **angka ≥ 16 digit** (mis. NIP/NIK) yang tersimpan "
                   "sebagai angka di Excel. Excel hanya menyimpan 15 digit signifikan, jadi digit belakangnya sudah "
                   "berubah dan pencocokan bisa salah. Ubah kolom menjadi format **Teks**, lalu isi ulang dari sumber aslinya.")

    sig = hashlib.md5(repr((p1["fk"], p1["header"], p1["drop"], p2["fk"], p2["header"], p2["drop"], key1, key2,
                            ignore_case, ignore_spaces, digits_only, strip_zeros, tuple(extras))).encode()).hexdigest()
    try:
        res = cached_compare(sig, key1, key2, ignore_case, ignore_spaces, digits_only, strip_zeros,
                             tuple(extras), df1, df2)
    except sl.SpreadsheetError as e:
        st.error(f"❌ {esc_md(e)}")
        return
    except Exception:
        st.error("❌ Terjadi kesalahan saat membandingkan data. Periksa kolom kunci dan coba lagi.")
        return
    s = res.summary

    st.markdown("#### 📊 Ringkasan")
    m = st.columns(4)
    m[0].metric("Baris File 1", f"{s['total_1']:,}")
    m[1].metric("✅ Sudah ada", f"{s['found']:,}")
    m[2].metric("❌ Belum ada", f"{s['missing']:,}")
    m[3].metric("➕ Hanya di File 2", f"{s['only_2']:,}")
    st.progress(min(s["pct_found"] / 100, 1.0), text=f"{s['pct_found']}% baris File 1 (yang punya kunci) sudah ada di File 2")
    if s["blank_1"] or s["dup_keys_1"] or s["dup_keys_2"]:
        st.caption(f"⚠️ Perlu dicek: {s['blank_1']:,} baris File 1 tanpa kunci · {s['dup_keys_1']:,} kunci kembar di File 1 "
                   f"· {s['dup_keys_2']:,} kunci kembar di File 2 (lihat tab **Perlu dicek**).")

    t1, t2, t3, t4 = st.tabs([f"✅ Sudah ada ({s['found']:,})", f"❌ Belum ada ({s['missing']:,})",
                              f"➕ Hanya di File 2 ({s['only_2']:,})",
                              f"⚠️ Perlu dicek ({s['blank_1'] + s['dup_keys_1'] + s['dup_keys_2']:,})"])
    with t1:
        st.caption("Baris File 1 yang kuncinya ditemukan di File 2. Kolom *Jumlah di File 2* = berapa kali kunci itu muncul.")
        table_tab(res.found, sig, "sudah_ada", "sudah_ada.csv", "Tidak ada baris yang kuncinya ditemukan di File 2.")
    with t2:
        st.caption("Baris File 1 yang kuncinya tidak ditemukan di File 2 (mis. belum hadir).")
        table_tab(res.missing, sig, "belum_ada", "belum_ada.csv", "🎉 Semua baris File 1 sudah ada di File 2.")
    with t3:
        st.caption("Baris File 2 yang kuncinya tidak ada di File 1 (mis. hadir tetapi tidak terdaftar).")
        table_tab(res.only_in_2, sig, "hanya_di_file_2", "hanya_di_file_2.csv", "Semua baris File 2 terdaftar di File 1.")
    with t4:
        if not (s["blank_1"] or s["dup_keys_1"] or s["dup_keys_2"] or s["blank_2"]):
            st.success("Tidak ada kunci kosong maupun kunci kembar.")
        if s["blank_1"]:
            st.markdown(f"**Baris File 1 tanpa kunci** ({s['blank_1']:,}) — tidak ikut dicocokkan:")
            show_df(res.blank_1.head(PREVIEW_ROWS))
        if s["blank_2"]:
            st.caption(f"File 2 memiliki {s['blank_2']:,} baris tanpa kunci (diabaikan).")
        if s["dup_keys_1"]:
            st.markdown(f"**Kunci kembar di File 1** ({s['dup_keys_1']:,}):")
            show_df(res.dup_1.head(PREVIEW_ROWS))
        if s["dup_keys_2"]:
            st.markdown(f"**Kunci kembar di File 2** ({s['dup_keys_2']:,}) — wajar jika satu orang tercatat beberapa kali:")
            show_df(res.dup_2.head(PREVIEW_ROWS))

    st.markdown("#### 📗 Laporan lengkap")
    report_key = "ck_report"
    if st.button("Siapkan laporan Excel", key="ck_report_btn",
                 help="Satu file Excel berisi Ringkasan, Sudah ada, Belum ada, Hanya di File 2, dan Perlu dicek."):
        st.session_state[report_key] = (sig, cached_report(sig, res))
    ready = st.session_state.get(report_key)
    if ready and ready[0] == sig:
        st.download_button("📥 Unduh laporan (.xlsx)", ready[1], file_name="laporan_cek_absensi.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="ck_report_dl")
    st.caption("Laporan dibuat saat tombol ditekan agar data besar tidak memperlambat halaman.")
