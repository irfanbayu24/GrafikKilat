"""
views/cek_absensi.py — menu "Cek Absensi": bandingkan dua daftar lewat kolom kunci (NIP/NIK/ID/email/nama).

File 1 = daftar acuan (mis. semua pegawai/peserta)   File 2 = daftar pembanding (mis. yang hadir)
Hasil: sudah ada · belum ada · hanya di File 2 · perlu dicek (kunci kosong/kembar) + laporan Excel.
"""
import hashlib
from datetime import time

import pandas as pd
import streamlit as st

import checklib
import durlib
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
def cached_matcher(sig, unit_col, name_col, noise_text, alias_text, overlap, fallback, max_extra, typo, _df1):
    return matchlib.Matcher(_df1, unit_col, name_col, noise=matchlib.parse_noise(noise_text), unit_overlap=overlap,
                            name_only_fallback=fallback, max_extra=max_extra, typo=typo,
                            unit_aliases=matchlib.parse_aliases(alias_text))


@st.cache_data(show_spinner="Mencocokkan nama…", max_entries=4, ttl=1800)
def cached_flex_run(sig, text_col, unit_col, name_col, _matcher, _df1, _df2):
    return matchlib.run(_matcher, _df1, _df2, text_col, unit_col, name_col)


@st.cache_data(show_spinner=False, max_entries=8, ttl=1800)
def cached_frame_raw(file_key, header_row, drop_summary, _raw):
    """Tabel File 2 tanpa penebakan tipe (teks tanggal-jam tetap asli, agar format bisa dibaca sesuai pilihan)."""
    return sl.build_dataframe(_raw, header_row, False, drop_summary, infer=False)


@st.cache_data(show_spinner="Menghitung durasi…", max_entries=6, ttl=1800)
def cached_durations(sig, _rule, source, join_col, leave_col, dur_col, order, n1, _df2raw, _row_persons):
    if source == "times":
        join, notes1 = durlib.parse_times(_df2raw[join_col], order)
        leave, notes2 = durlib.parse_times(_df2raw[leave_col], order)
        table, info = durlib.attendance_table(n1, _row_persons, join, leave, _rule)
        return table, info, list(dict.fromkeys(notes1 + notes2))
    table, info = durlib.attendance_from_duration(n1, _row_persons, _df2raw[dur_col], _rule)
    return table, info, []


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
        session_hdr = durlib.find_session_header(raw)
        guess = session_hdr if session_hdr is not None else sl.guess_header_row(raw)
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
    return {"name": name, "fk": fk, "df": df, "header": header_row, "drop": drop_summary, "raw": raw}


MODE_EXACT = "Satu kolom kunci (persis) — NIP/NIK/ID/email"
MODE_FLEX = "Unit + Nama → satu kolom teks di File 2 (fleksibel, nama boleh tidak lengkap)"
NO_UNIT = "(Tidak ada — cocokkan nama saja)"
_CONF_RANK = {"Rendah": 0, "Sedang": 1, "Tinggi": 2}
REASON_TIPS = {
    "Unit tidak dikenali di teks": "Coba isi kolom *Singkatan unit* di Aturan pencocokan, atau aktifkan 'cocokkan nama saja' "
                                   "jika sebagian peserta tidak menulis unit. Lihat kolom *Petunjuk terdekat*.",
    "Nama terlalu pendek (hanya inisial) atau kosong": "Peserta perlu menulis minimal satu kata nama (≥ 2 huruf).",
    "Ada kata di teks yang tidak dikenali": "Ubah 'Kata di File 2 yang tidak dikenali' menjadi 'Boleh 1 kata', atau "
                                            "tambahkan kata itu ke daftar kata yang diabaikan.",
    "Nama ada, tetapi di unit lain": "Unit yang ditulis peserta berbeda dengan unit di File 1. Periksa unitnya, atau aktifkan "
                                     "'cocokkan nama saja'.",
    "Nama tidak ditemukan di unit tersebut": "Nama mungkin belum terdaftar di File 1 atau ejaannya jauh berbeda. "
                                             "Lihat *Petunjuk terdekat*.",
    "Nama tidak ditemukan di File 1": "Nama mungkin belum terdaftar di File 1 atau ejaannya jauh berbeda.",
}


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
        unit_rule = o1.selectbox("Pencocokan unit", ["Longgar — cukup separuh kata unit (disarankan)",
                                                     "Ketat — semua kata unit harus ada di teks"], key="ck_f_unitrule",
                                 disabled=unit_col is None,
                                 help="Nama unit di data sering panjang atau berkurung, mis. 'Sekretariat Badan (Setban)', "
                                      "sedangkan peserta menulis singkat. Aturan ketat menolak yang tidak menulis semua kata.")
        extra_rule = o2.selectbox("Kata di File 2 yang tidak dikenali", ["Tidak boleh ada (ketat)",
                                                                         "Boleh 1 kata (longgar)"], key="ck_f_extra")
        fallback = o1.checkbox("Jika unit tidak terbaca di File 2, coba cocokkan nama saja (keyakinan rendah)", False,
                               key="ck_f_fallback", disabled=unit_col is None,
                               help="Berguna jika sebagian peserta menulis singkatan unit yang tidak dikenal. "
                                    "Nama lengkap di File 1 harus ikut tercocokkan seluruhnya.")
        typo = o2.checkbox("Toleransi salah ketik 1 huruf (kata ≥ 5 huruf)", True, key="ck_f_typo")
        noise_text = st.text_area("Kata yang diabaikan (sapaan, gelar, label perangkat) — pisahkan dengan spasi",
                                  matchlib.DEFAULT_NOISE, height=110, key="ck_f_noise")
        alias_text = st.text_area("Singkatan unit (opsional) — satu baris per unit", "", height=100, key="ck_f_alias",
                                  disabled=unit_col is None,
                                  placeholder="Pusdatin, PDI = Pusat Data dan Informasi\nSetban = Sekretariat Badan (Setban)",
                                  help="Format: singkatan, singkatan lain = nama unit persis seperti di File 1. "
                                       "Berguna bila peserta menulis singkatan yang tidak bisa ditebak otomatis.")
    overlap = 0.5 if unit_rule.startswith("Longgar") else 1.0
    max_extra = 1 if extra_rule.startswith("Boleh") else 0

    sig_m = hashlib.md5(repr((p1["fk"], p1["header"], p1["drop"], unit_col, name_col, noise_text, alias_text, overlap, fallback,
                              max_extra, typo)).encode()).hexdigest()
    try:
        matcher = cached_matcher(sig_m, unit_col, name_col, noise_text, alias_text, overlap, fallback, max_extra, typo, df1)
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
    if matcher.unknown_alias_units:
        st.warning("⚠️ Singkatan unit mengacu ke unit yang **tidak ditemukan di File 1**: "
                   + " · ".join(esc_md(u) for u in matcher.unknown_alias_units)
                   + ". Tulis nama unit persis seperti di File 1 (huruf besar/kecil dan tanda baca tidak berpengaruh).")
    if s["tidak_cocok"]:
        bad = res.detail[res.detail["Status"] == "Tidak cocok"]
        reasons = bad["Alasan tidak cocok"].value_counts()
        with st.expander(f"🩺 Mengapa {s['tidak_cocok']:,} entri tidak cocok? (klik untuk melihat penyebab dan saran)",
                         expanded=s["tidak_cocok"] >= max(s["cocok"], 1)):
            for reason, n in reasons.items():
                tip = REASON_TIPS.get(reason, "")
                st.markdown(f"**{n:,}×** {reason}" + (f"  \n_{tip}_" if tip else ""))
            show_df(bad[["Teks di File 2", "Alasan tidak cocok", "Petunjuk terdekat"]].head(25))
            st.caption("Menampilkan 25 contoh pertama. Daftar lengkap ada di tab **Rincian**.")
    if s["belum_pasti"] or s["ambigu"]:
        st.info("❓ **Belum pasti** = nama di File 2 cocok dengan lebih dari satu orang di File 1 (mis. dua 'Budi' di unit yang sama), "
                "jadi tidak ditebak. Lihat tab **Rincian** dan lengkapi nama di File 2 bila perlu.")

    with st.expander("🔍 Uji satu teks dari File 2 (lihat cara aplikasi membacanya)"):
        probe = st.text_input("Ketik atau tempel teks, mis. Keuangan - Budi", key="ck_f_probe")
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

    render_duration(p1, p2, df1, res.row_persons, res.row_ambiguous, matcher.blank_rows, "flex", sig)


def _free_col(existing, name: str) -> str:
    cand, n = name, 2
    while cand in set(existing):
        cand, n = f"{name} ({n})", n + 1
    return cand


METHODS = ["Total waktu terhubung — jeda reconnect ikut dihitung (disarankan)",
           "Total waktu terhubung — jeda reconnect tidak dihitung",
           "Rentang dari join pertama sampai leave terakhir"]


def render_duration(p1: dict, p2: dict, df1: pd.DataFrame, row_persons: list, row_ambiguous: list,
                    blank_pos: list, scope: str, base_sig: str) -> None:
    """Aturan durasi untuk log join/leave (mis. laporan Zoom). Dipakai oleh mode persis maupun fleksibel."""
    df2 = p2["df"]
    k = f"{scope}_{p2['fk']}"
    st.markdown("#### ⏱️ Aturan durasi kehadiran")
    if not st.checkbox("Aktifkan aturan durasi (mis. laporan Zoom: hadir bila ≥ 60 menit, reconnect dihitung)", False,
                       key=f"du_on_{scope}"):
        st.caption("Aktifkan untuk menghitung lama kehadiran tiap orang dari kolom Join/Leave di File 2. Sesi yang putus "
                   "lalu masuk lagi (reconnect) digabung, dan sesi yang tumpang tindih tidak dihitung dobel.")
        return
    cols2 = list(df2.columns)
    g_join, g_leave, g_dur = durlib.guess_session_columns(cols2)
    sources = ["Jam join & leave (disarankan)", "Kolom durasi (menit) saja"]
    source = st.radio("Sumber waktu di File 2", sources, horizontal=True, index=0 if (g_join and g_leave) else 1,
                      key=f"du_src_{k}")
    c1, c2, c3 = st.columns(3)
    join_col = leave_col = dur_col = None
    order = "auto"
    if source == sources[0]:
        join_col = c1.selectbox("Kolom Join (bergabung)", cols2, index=cols2.index(g_join) if g_join else 0, key=f"du_join_{k}")
        leave_col = c2.selectbox("Kolom Leave (keluar)", cols2, index=cols2.index(g_leave) if g_leave else min(1, len(cols2) - 1),
                                 key=f"du_leave_{k}")
        order = {"Otomatis": "auto", "Hari/Bulan/Tahun": "dmy", "Bulan/Hari/Tahun": "mdy"}[
            c3.selectbox("Format tanggal", ["Otomatis", "Hari/Bulan/Tahun", "Bulan/Hari/Tahun"], key=f"du_order_{k}")]
    else:
        dur_col = c1.selectbox("Kolom durasi (menit)", cols2, index=cols2.index(g_dur) if g_dur else 0, key=f"du_dur_{k}")
        st.caption("Tanpa jam join/leave, durasi tiap baris dijumlahkan per orang. Sesi tumpang tindih dan reconnect "
                   "tidak bisa dikenali, serta jam acara tidak dapat diterapkan.")
    r1, r2, r3 = st.columns(3)
    threshold = r1.number_input("Batas durasi hadir (menit)", 1.0, 1440.0, 60.0, step=5.0, key=f"du_thr_{k}")
    inclusive = r2.selectbox("Ketentuan batas", ["Minimal sama dengan batas (≥)", "Lebih dari batas (>)"],
                             key=f"du_incl_{k}").startswith("Minimal")
    tolerance = r3.number_input("Toleransi jeda reconnect (menit)", 0.0, 240.0, 10.0, step=1.0, key=f"du_tol_{k}",
                                help="Orang yang putus lalu masuk lagi dalam jeda sebesar ini dianggap tersambung.")
    with st.expander("⚙️ Opsi lanjutan"):
        method_label = st.selectbox("Cara menghitung durasi", METHODS, key=f"du_method_{k}")
        use_window = st.checkbox("Batasi ke jam acara (waktu di ruang tunggu atau setelah acara tidak dihitung)", False,
                                 key=f"du_win_{k}", disabled=source != sources[0])
        w1, w2 = st.columns(2)
        win_start = w1.time_input("Jam mulai acara", time(8, 0), key=f"du_ws_{k}", disabled=not use_window)
        win_end = w2.time_input("Jam selesai acara", time(10, 0), key=f"du_we_{k}", disabled=not use_window)
    window = use_window and source == sources[0]
    rule = durlib.Rule(threshold_min=float(threshold), inclusive=inclusive, tolerance_min=float(tolerance),
                       count_gaps=method_label != METHODS[1], method="span" if method_label == METHODS[2] else "merged",
                       win_start=win_start if window else None, win_end=win_end if window else None)
    st.caption("📐 " + rule.describe())

    df2raw, _ = cached_frame_raw(p2["fk"], p2["header"], p2["drop"], p2["raw"])
    if len(df2raw) != len(df2):
        st.error("❌ Jumlah baris File 2 tidak sinkron. Muat ulang halaman lalu coba lagi.")
        return
    sig = hashlib.md5(repr((base_sig, source, join_col, leave_col, dur_col, order, rule)).encode()).hexdigest()
    try:
        table, info, notes = cached_durations(sig, rule, "times" if source == sources[0] else "duration", join_col,
                                              leave_col, dur_col, order, len(df1), df2raw, row_persons)
    except Exception:
        st.error("❌ Gagal menghitung durasi. Periksa kolom waktu yang dipilih (format tanggal/jam).")
        return
    for n in notes:
        st.warning("⚠️ " + esc_md(n))
    if info["rows_matched"] == 0:
        st.warning("Belum ada baris File 2 yang cocok dengan File 1, sehingga durasi belum bisa dihitung. "
                   "Periksa pengaturan pencocokan di atas.")
        return
    skipped = [f"{info['unreadable']:,} sesi dengan waktu tidak terbaca" if info["unreadable"] else "",
               f"{info['negative']:,} sesi dengan leave sebelum join" if info["negative"] else "",
               f"{info['outside_window']:,} sesi di luar jam acara" if info["outside_window"] else ""]
    if any(skipped):
        st.caption("Dilewati: " + " · ".join(x for x in skipped if x))

    ambiguous, blank = {p for lst in row_ambiguous for p in lst}, set(blank_pos)
    people = df1.copy()
    status_col = _free_col(people.columns, "Status akhir")
    people[status_col] = [durlib.final_status(t, i in ambiguous, i in blank) for i, t in enumerate(table["Status durasi"])]
    added = {}
    for c in durlib.COLS:
        added[c] = _free_col(people.columns, c)
        people[added[c]] = table[c].to_numpy()
    by = lambda v: people[people[status_col] == v].reset_index(drop=True)   # noqa: E731
    hadir, short = by("Hadir"), by(durlib.STATUS_SHORT)
    belum, pasti = by("Belum hadir"), by("Hadir (belum pasti)")
    short = short.sort_values(added["Total durasi (menit)"], ascending=False, kind="stable").reset_index(drop=True)

    st.markdown("##### 📊 Hasil berdasarkan durasi")
    valid = len(people) - len(blank)
    m = st.columns(4)
    m[0].metric("✅ Hadir (memenuhi durasi)", f"{len(hadir):,}")
    m[1].metric("⏳ Kurang dari batas", f"{len(short):,}")
    m[2].metric("❌ Belum hadir", f"{len(belum):,}")
    m[3].metric("❓ Hadir (belum pasti)", f"{len(pasti):,}")
    st.progress(min(len(hadir) / valid, 1.0) if valid else 0.0,
                text=f"{round(len(hadir) / valid * 100, 1) if valid else 0.0}% memenuhi durasi dari baris File 1 yang punya kunci")
    reconnects = int((table["Reconnect (jeda kecil)"] > 0).sum())
    if reconnects:
        st.caption(f"🔁 {reconnects:,} orang tercatat reconnect (putus lalu masuk lagi dengan jeda ≤ {tolerance:g} menit).")
    if len(pasti):
        st.info("❓ **Belum pasti** = nama di File 2 cocok dengan lebih dari satu orang, sehingga durasinya tidak dikreditkan "
                "ke siapa pun. Lengkapi nama di File 2 atau periksa tab Rincian pada hasil pencocokan.")

    t1, t2, t3, t4, t5 = st.tabs([f"✅ Hadir ({len(hadir):,})", f"⏳ Kurang dari batas ({len(short):,})",
                                  f"❌ Belum hadir ({len(belum):,})", f"❓ Belum pasti ({len(pasti):,})", "📋 Semua"])
    with t1:
        flex_tab(hadir, sig, "durasi_hadir", "Belum ada yang memenuhi batas durasi.")
    with t2:
        st.caption("Terhubung tetapi belum memenuhi batas; yang paling dekat dengan batas ada di atas.")
        flex_tab(short, sig, "durasi_kurang", "Tidak ada yang kurang dari batas.")
    with t3:
        flex_tab(belum, sig, "durasi_belum_hadir", "🎉 Semua baris File 1 sudah terhubung.")
    with t4:
        flex_tab(pasti, sig, "durasi_belum_pasti", "Tidak ada yang ambigu.")
    with t5:
        flex_tab(people, sig, "durasi_semua", "Tidak ada data.")

    rkey = f"du_report_{scope}"
    if st.button("Siapkan laporan Excel (durasi)", key=f"du_report_btn_{k}"):
        summary = pd.DataFrame([
            ("Aturan", rule.describe()), ("Total baris File 1", len(people)), ("Baris tanpa nama/kunci", len(blank)),
            ("Hadir (memenuhi durasi)", len(hadir)), ("Kurang dari batas", len(short)), ("Belum hadir", len(belum)),
            ("Hadir (belum pasti)", len(pasti)), ("Orang yang reconnect", reconnects)], columns=["Keterangan", "Nilai"])
        sheets = {"Ringkasan durasi": summary.astype(str), "Hadir": hadir, "Kurang dari batas": short,
                  "Belum hadir": belum, "Belum pasti": pasti, "Semua": people}
        st.session_state[rkey] = (sig, sl.to_excel_sheets(sheets))
    ready = st.session_state.get(rkey)
    if ready and ready[0] == sig:
        st.download_button("📥 Unduh laporan durasi (.xlsx)", ready[1], file_name="laporan_durasi_kehadiran.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"du_report_dl_{k}")


def render_howto() -> None:
    with st.expander("📘 Cara pakai", expanded=True):
        st.markdown(
            "1. **File 1** = daftar acuan, mis. semua pegawai/peserta yang seharusnya hadir.\n"
            "2. **File 2** = daftar pembanding, mis. daftar hadir (rekap Zoom, absensi, dll.).\n"
            "3. Pilih **kolom kunci** di masing-masing file (NIP, NIK, ID, email, atau nama). Nama kolomnya boleh berbeda.\n"
            "4. Hasil muncul otomatis: **sudah ada** (hadir), **belum ada** (belum hadir), **hanya di File 2** "
            "(hadir tapi tidak terdaftar), dan **perlu dicek** (kunci kosong/kembar). Semuanya bisa diunduh dalam satu laporan Excel.\n\n"
            "5. **Mode fleksibel:** jika File 1 punya kolom **Unit** dan **Nama** terpisah sedangkan File 2 hanya satu kolom teks "
            "seperti `Keuangan - Budi` (nama boleh tidak lengkap), pilih *Cara mencocokkan → fleksibel*.\n\n"
            "6. **Aturan durasi (mis. laporan Zoom):** aktifkan di bagian bawah halaman hasil untuk menghitung lama hadir dari kolom "
            "Join/Leave. Orang yang putus lalu masuk lagi digabung, dan hadir bila total durasinya memenuhi batas (mis. 60 menit).\n\n"
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

    opts = checklib.Options(ignore_case, ignore_spaces, digits_only, strip_zeros)
    blank_pos = [i for i, kk in enumerate(checklib.normalize_keys(df1[key1], opts).tolist()) if not kk]
    render_duration(p1, p2, df1, checklib.row_assignments(df1, key1, df2, key2, opts), [], blank_pos, "exact", sig)
