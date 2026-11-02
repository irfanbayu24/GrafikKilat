"""
views/visualisasi.py — menu "Visualisasi Spreadsheet".

Siapa pun dapat mengunggah spreadsheet-nya sendiri (.xlsx .xlsm .xls .ods .csv) dan langsung
mendapat ringkasan, grafik interaktif, analisis sebaran/korelasi, serta tabel yang bisa difilter.
"""
import hashlib

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

import sheetlib as sl
from views.common import DF_KW, PLOT_KW, esc_md, remembered_upload


MAX_DISPLAY_ROWS = 5000
EXCEL_MAX_ROWS = 50_000


def show_plot(fig, key: str, height: int = 440) -> None:
    fig.update_layout(height=height, margin=dict(t=50, b=20, l=10, r=10), legend_title_text="")
    st.plotly_chart(fig, key=key, **PLOT_KW)


def show_df(data: pd.DataFrame, **kw) -> None:
    st.dataframe(data, hide_index=kw.pop("hide_index", True), **DF_KW, **kw)


def fmt(x) -> str:
    if pd.isna(x):
        return "-"
    return f"{x:,.0f}" if abs(x) >= 1000 or float(x).is_integer() else f"{x:,.2f}"


# ---------------------------------------------------------------------------
# Cache (hanya di memori server; kedaluwarsa otomatis)
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Membaca spreadsheet…", max_entries=3, ttl=1800)
def cached_load(filename: str, data: bytes):
    kind = sl.validate_upload(filename, data)      # tolak file yang bukan spreadsheet asli
    return sl.load_workbook(kind, data)


@st.cache_data(show_spinner=False, max_entries=8, ttl=1800)
def cached_build(file_id, sheet, header_row, drop_dups, drop_summary, _raw):
    return sl.build_dataframe(_raw, header_row, drop_dups, drop_summary)


@st.cache_data(show_spinner=False)
def cached_sample() -> bytes:
    return sl.make_sample_xlsx()


@st.cache_data(show_spinner=False, max_entries=4, ttl=900)
def cached_csv(sig, _df) -> bytes:
    return sl.to_csv_bytes(_df)


@st.cache_data(show_spinner=False, max_entries=2, ttl=900)
def cached_xlsx(sig, _df) -> bytes:
    return sl.to_excel_bytes(_df)


# ---------------------------------------------------------------------------
# Sidebar & halaman awal
# ---------------------------------------------------------------------------
def sidebar_upload():
    with st.sidebar:
        st.markdown("## 📊 Visualisasi Spreadsheet")
        remembered = remembered_upload(
            "vis", "Unggah spreadsheet Anda", list(sl.ALLOWED_EXTENSIONS),
            help="Hanya file spreadsheet: " + ", ".join("." + e for e in sl.ALLOWED_EXTENSIONS),
        )
        st.caption(f"Format: {', '.join('.' + e for e in sl.ALLOWED_EXTENSIONS)} · maks. {sl.MAX_FILE_MB} MB")
        if remembered is not None:
            st.session_state.pop("use_sample", None)
            return remembered
        if st.session_state.get("use_sample"):
            st.info("Sedang menampilkan **data contoh**.")
            if st.button("Tutup data contoh"):
                st.session_state.pop("use_sample", None)
                st.rerun()
            return "data_contoh.xlsx", cached_sample()
    return None


def render_welcome() -> None:
    st.title("📊 Visualisasi Spreadsheet")
    st.markdown(
        "Unggah spreadsheet Anda dan dapatkan **ringkasan, grafik interaktif, dan analisis data** "
        "secara otomatis — tanpa rumus, tanpa kode."
    )
    c1, c2, c3 = st.columns(3)
    c1.markdown("#### 1️⃣ Unggah\nPilih file di panel kiri. Judul di atas tabel dan baris kosong dikenali otomatis.")
    c2.markdown("#### 2️⃣ Jelajahi\nLihat ringkasan, buat grafik sendiri, filter data, dan cek korelasi antar kolom.")
    c3.markdown("#### 3️⃣ Unduh\nSimpan data hasil filter sebagai CSV atau Excel, dan grafik sebagai gambar PNG.")
    st.divider()
    left, right = st.columns([1, 2])
    with left:
        st.markdown("##### Belum punya file?")
        if st.button("🧪 Coba dengan data contoh", type="primary"):
            st.session_state["use_sample"] = True
            st.rerun()
    with right:
        with st.expander("🔒 Format yang didukung & privasi", expanded=True):
            st.markdown(
                f"- **Format:** {', '.join('`.' + e + '`' for e in sl.ALLOWED_EXTENSIONS)} (maks. {sl.MAX_FILE_MB} MB, "
                f"{sl.MAX_ROWS:,} baris per sheet).\n"
                "- **Hanya spreadsheet:** isi file diperiksa, bukan sekadar ekstensinya. File lain yang diganti "
                "namanya akan ditolak.\n"
                "- **Privasi:** file hanya diproses di memori server, tidak ditulis ke disk, dan otomatis "
                "dibuang dari cache (±30 menit). Jangan unggah data rahasia ke server yang tidak Anda percaya.\n"
                "- **Tips:** file dengan password harus dibuka kuncinya terlebih dahulu."
            )


def build_filter(df: pd.DataFrame, col: str, role: str, key: str) -> pd.Series:
    """Widget filter sesuai jenis kolom; mengembalikan mask boolean."""
    s = df[col]
    all_rows = pd.Series(True, index=df.index)
    if role == "kategori":
        options = sorted(s.dropna().astype(str).unique().tolist())
        picked = st.multiselect(esc_md(col), options, key=f"f_{key}")
        return s.astype(str).isin(picked) if picked else all_rows
    if role == "numerik":
        lo, hi = s.min(), s.max()
        if pd.isna(lo) or lo == hi:
            return all_rows
        lo, hi = (int(lo), int(hi)) if pd.api.types.is_integer_dtype(s) else (float(lo), float(hi))
        sel = st.slider(esc_md(col), lo, hi, (lo, hi), key=f"f_{key}")
        mask = s.between(sel[0], sel[1])
        return mask | s.isna() if sel == (lo, hi) else mask
    if role == "tanggal":
        lo, hi = s.min(), s.max()
        if pd.isna(lo):
            return all_rows
        sel = st.date_input(esc_md(col), value=(lo.date(), hi.date()), min_value=lo.date(), max_value=hi.date(),
                            key=f"f_{key}")
        if isinstance(sel, (tuple, list)) and len(sel) == 2:
            return (s >= pd.Timestamp(sel[0])) & (s < pd.Timestamp(sel[1]) + pd.Timedelta(days=1))
    return all_rows


# ---------------------------------------------------------------------------
# Tab 1: Ringkasan
# ---------------------------------------------------------------------------
def auto_figures(df: pd.DataFrame, roles: dict) -> list:
    figs = []
    cats = sorted([c for c in df.columns if roles[c] == "kategori" and 2 <= df[c].nunique() <= 20],
                  key=lambda c: df[c].nunique())
    nums = [c for c in df.columns if roles[c] == "numerik"]
    dates = [c for c in df.columns if roles[c] == "tanggal"]

    if dates:
        d = dates[0]
        gran = sl.auto_granularity(df[d])
        work = df.assign(**{d: sl.bucket_dates(df[d], gran)}) if gran != "Asli" else df
        if nums:
            label = sl.default_agg(nums[0])
            ylab = f"{label} {nums[0]}"
            t = sl.aggregate(work, [d], nums[0], sl.AGG_CHOICES[label])
        else:
            ylab = "Jumlah baris"
            t = sl.aggregate(work, [d])
        t = t.sort_values(d).rename(columns={"Nilai": ylab})
        figs.append((f"{ylab} per {d} ({gran.lower()})", px.line(t, x=d, y=ylab, markers=len(t) <= 60)))
    for c in cats[:2]:
        t = sl.aggregate(df, [c]).sort_values("Nilai", ascending=False).rename(columns={"Nilai": "Jumlah baris"})
        fig = px.bar(t, x=c, y="Jumlah baris", text_auto=True)
        figs.append((f"Jumlah baris per {c}", fig))
    for v in nums[:2]:
        figs.append((f"Sebaran nilai {v}", px.histogram(df, x=v, nbins=30)))
    return figs[:4]


def tab_summary(df: pd.DataFrame, roles: dict, key: str) -> None:
    n, n_cols = len(df), len(df.columns)
    missing = int(df.isna().sum().sum())
    m = st.columns(5)
    m[0].metric("Jumlah baris", f"{n:,}")
    m[1].metric("Jumlah kolom", n_cols)
    m[2].metric("Kolom angka", sum(r == "numerik" for r in roles.values()))
    m[3].metric("Sel kosong", f"{missing / max(n * n_cols, 1) * 100:.1f}%")
    m[4].metric("Baris duplikat", f"{int(df.duplicated().sum()):,}")

    st.markdown("#### ✨ Visual otomatis")
    figs = auto_figures(df, roles)
    if not figs:
        st.info("Belum ada visual otomatis untuk data ini. Coba tab **Studio Grafik** untuk membuat grafik sendiri.")
    for i in range(0, len(figs), 2):
        cols = st.columns(2)
        for j, (title, fig) in enumerate(figs[i:i + 2]):
            with cols[j]:
                st.markdown(f"**{esc_md(title)}**")
                show_plot(fig, key=f"{key}_auto_{i + j}", height=340)

    with st.expander("🧪 Kualitas data per kolom"):
        show_df(sl.quality_report(df, roles))
    nums = [c for c in df.columns if roles[c] == "numerik"]
    if nums:
        with st.expander("📐 Statistik deskriptif kolom angka"):
            show_df(sl.describe_numeric(df, nums))


# ---------------------------------------------------------------------------
# Tab 2: Studio Grafik
# ---------------------------------------------------------------------------
KINDS = ["Batang", "Garis", "Area", "Pie / Donut", "Sebaran (Scatter)", "Histogram", "Box Plot"]


def chart_grouped(df, roles, kind, num_cols, color_cols, key):
    cols = list(df.columns)
    r1 = st.columns(4)
    default_x = next((c for c in cols if roles[c] in ("tanggal", "kategori")), cols[0])
    x = r1[0].selectbox("Sumbu X / kelompok", cols, index=cols.index(default_x), key=f"{key}_x")
    y = r1[1].selectbox("Nilai (Y)", ["Jumlah baris"] + num_cols, key=f"{key}_y")
    if y == "Jumlah baris":
        agg, ylabel, value = "size", "Jumlah baris", None
    else:
        choices = ["Jumlah (sum)"] if kind == "Pie / Donut" else list(sl.AGG_CHOICES)
        dflt = sl.default_agg(y)
        agg_label = r1[2].selectbox("Cara menghitung", choices,
                                    index=choices.index(dflt) if dflt in choices else 0, key=f"{key}_agg_{y}")
        agg, ylabel, value = sl.AGG_CHOICES[agg_label], f"{agg_label} {y}", y
    color = None
    if kind != "Pie / Donut":
        pick = r1[3].selectbox("Pisahkan berdasarkan", ["(Tidak ada)"] + [c for c in color_cols if c != x],
                               key=f"{key}_color")
        color = None if pick == "(Tidak ada)" else pick

    is_date = roles[x] == "tanggal"
    num_like = pd.api.types.is_numeric_dtype(df[x])
    cat_x = not is_date and not num_like
    period_x = num_like and roles[x] == "kategori"          # mis. kolom Tahun / Bulan berupa angka

    r2 = st.columns(3)
    work = df
    if is_date:
        opts = ["Asli"] + list(sl.GRANULARITY)
        gran = r2[0].selectbox("Kelompokkan tanggal per", opts,
                               index=opts.index(sl.auto_granularity(df[x])), key=f"{key}_gran_{x}")
        if gran != "Asli":
            work = df.assign(**{x: sl.bucket_dates(df[x], gran)})
    agg_df = sl.aggregate(work, [x] + ([color] if color else []), value, agg).rename(columns={"Nilai": ylabel})

    topn, order_mode, horizontal = 15, "Nilai tertinggi", False
    if cat_x:
        topn = int(r2[0].number_input("Tampilkan N kategori teratas", 3, 200, 15, key=f"{key}_topn"))
        if kind == "Batang":
            order_mode = r2[1].selectbox("Urutkan", ["Nilai tertinggi", "Nilai terendah", "Abjad"], key=f"{key}_ord")
    if kind == "Batang" and not is_date and not num_like:
        horizontal = r2[2].checkbox("Batang horizontal", key=f"{key}_h")

    if cat_x and kind != "Pie / Donut":
        totals = agg_df.groupby(x)[ylabel].sum()
        if len(totals) > topn:
            agg_df = agg_df[agg_df[x].isin(totals.nlargest(topn).index)]
            st.caption(f"Menampilkan {topn} dari {len(totals):,} kategori dengan nilai terbesar.")
    elif not cat_x and not is_date and agg_df[x].nunique() > 300:
        keep = sorted(agg_df[x].unique())[:300]
        agg_df = agg_df[agg_df[x].isin(keep)]
        st.warning("Sumbu X memiliki >300 nilai berbeda; hanya 300 pertama yang ditampilkan.")

    title = f"{ylabel} per {x}"
    if kind == "Pie / Donut":
        agg_df = agg_df[agg_df[ylabel] > 0].sort_values(ylabel, ascending=False)
        if is_date:
            agg_df[x] = agg_df[x].dt.strftime("%Y-%m-%d")
        if len(agg_df) > topn:
            rest = agg_df.iloc[topn:][ylabel].sum()
            agg_df = pd.concat([agg_df.head(topn), pd.DataFrame({x: ["Lainnya"], ylabel: [rest]})], ignore_index=True)
        agg_df[x] = agg_df[x].astype(str)
        if agg_df.empty:
            st.warning("Tidak ada nilai positif untuk ditampilkan sebagai pie.")
            return None
        fig = px.pie(agg_df, names=x, values=ylabel, hole=0.45, title=title)
        fig.update_traces(textinfo="percent+label", sort=False)
        return fig, agg_df

    if kind == "Batang":
        text = ".3~s" if len(agg_df) <= 40 else False
        if cat_x:
            tot = agg_df.groupby(x, sort=False)[ylabel].sum()
            order = (sorted(tot.index) if order_mode == "Abjad"
                     else tot.sort_values(ascending=(order_mode == "Nilai terendah")).index.tolist())
        else:
            agg_df = agg_df.sort_values(x)
        if horizontal:
            fig = px.bar(agg_df, x=ylabel, y=x, color=color, orientation="h", barmode="group", text_auto=text, title=title)
            fig.update_yaxes(categoryorder="array", categoryarray=order, autorange="reversed")
        else:
            fig = px.bar(agg_df, x=x, y=ylabel, color=color, barmode="group", text_auto=text, title=title)
            if cat_x:
                fig.update_xaxes(categoryorder="array", categoryarray=order)
    else:  # Garis / Area
        if cat_x:
            st.caption("💡 Grafik garis/area paling cocok bila sumbu X berupa tanggal atau angka.")
        agg_df = agg_df.sort_values(x)
        if kind == "Garis":
            fig = px.line(agg_df, x=x, y=ylabel, color=color, markers=len(agg_df) <= 80, title=title)
        else:
            fig = px.area(agg_df, x=x, y=ylabel, color=color, title=title)
    if period_x:
        fig.update_xaxes(type="category")
    return fig, agg_df


def chart_scatter(df, roles, num_cols, color_cols, key):
    if len(num_cols) < 2:
        st.info("Grafik sebaran membutuhkan minimal 2 kolom angka.")
        return None
    r = st.columns(4)
    x = r[0].selectbox("Sumbu X", num_cols, key=f"{key}_sx")
    y = r[1].selectbox("Sumbu Y", num_cols, index=1, key=f"{key}_sy")
    color = r[2].selectbox("Warna berdasarkan", ["(Tidak ada)"] + color_cols, key=f"{key}_sc")
    size = r[3].selectbox("Ukuran titik", ["(Tidak ada)"] + num_cols, key=f"{key}_ss")
    trend = st.checkbox("Tampilkan garis tren", key=f"{key}_trend")
    color = None if color == "(Tidak ada)" else color
    size = None if size == "(Tidak ada)" else size

    data = df[list(dict.fromkeys(c for c in (x, y, color, size) if c))].dropna()
    if size:
        data = data[data[size] >= 0]
    if data.empty:
        st.warning("Tidak ada baris lengkap untuk kombinasi kolom ini.")
        return None
    if len(data) > 5000:
        st.caption(f"Menampilkan 5.000 titik acak dari {len(data):,} baris.")
        data = data.sample(5000, random_state=0)
    fig = px.scatter(data, x=x, y=y, color=color, size=size, opacity=0.7, size_max=18, title=f"{y} vs {x}")
    if trend and len(data) >= 2 and data[x].nunique() > 1:
        slope, intercept = np.polyfit(data[x], data[y], 1)
        xs = np.array([data[x].min(), data[x].max()])
        fig.add_scatter(x=xs, y=slope * xs + intercept, mode="lines", name="Tren linear",
                        line=dict(dash="dash", color="gray"))
        st.caption(f"Korelasi (Pearson) antara X dan Y: **{data[x].corr(data[y]):.3f}**")
    return fig, data


def chart_hist(df, roles, num_cols, color_cols, key):
    if not num_cols:
        st.info("Histogram membutuhkan minimal 1 kolom angka.")
        return None
    r = st.columns(3)
    x = r[0].selectbox("Kolom angka", num_cols, key=f"{key}_hx")
    bins = int(r[1].number_input("Jumlah interval (bins)", 5, 100, 20, key=f"{key}_bins"))
    color = r[2].selectbox("Warna berdasarkan", ["(Tidak ada)"] + color_cols, key=f"{key}_hc")
    color = None if color == "(Tidak ada)" else color
    fig = px.histogram(df, x=x, nbins=bins, color=color, marginal="box", barmode="overlay",
                       title=f"Sebaran {x}")
    if color:
        fig.update_traces(opacity=0.75)
    return fig, sl.describe_numeric(df, [x])


def chart_box(df, roles, num_cols, key):
    if not num_cols:
        st.info("Box plot membutuhkan minimal 1 kolom angka.")
        return None
    groups = [c for c in df.columns if roles[c] == "kategori" and df[c].nunique() <= 30]
    r = st.columns(2)
    y = r[0].selectbox("Kolom angka", num_cols, key=f"{key}_by")
    g = r[1].selectbox("Kelompokkan berdasarkan", ["(Tidak ada)"] + groups, key=f"{key}_bg")
    if g == "(Tidak ada)":
        fig = px.box(df, y=y, points="outliers", title=f"Box plot {y}")
    else:
        fig = px.box(df, x=g, y=y, color=g, points="outliers", title=f"{y} per {g}")
        fig.update_layout(showlegend=False)
    return fig, sl.describe_numeric(df, [y])


def tab_studio(df: pd.DataFrame, roles: dict, key: str) -> None:
    num_cols = [c for c in df.columns if roles[c] == "numerik"]
    color_cols = [c for c in df.columns if roles[c] == "kategori" and df[c].nunique() <= 12]
    kind = st.selectbox("Jenis grafik", KINDS, key=f"{key}_kind")
    if kind in ("Batang", "Garis", "Area", "Pie / Donut"):
        result = chart_grouped(df, roles, kind, num_cols, color_cols, key)
    elif kind.startswith("Sebaran"):
        result = chart_scatter(df, roles, num_cols, color_cols, key)
    elif kind == "Histogram":
        result = chart_hist(df, roles, num_cols, color_cols, key)
    else:
        result = chart_box(df, roles, num_cols, key)
    if result is None:
        return
    fig, table = result
    show_plot(fig, key=f"{key}_main_{kind}", height=480)
    st.caption("💡 Arahkan kursor ke grafik untuk melihat nilai. Ikon kamera di pojok kanan atas menyimpan grafik sebagai PNG.")
    with st.expander("📄 Data di balik grafik"):
        show_df(table)
        st.download_button("📥 Unduh data grafik (CSV)", sl.to_csv_bytes(table), "data_grafik.csv", "text/csv",
                           key=f"{key}_dl_chart")


# ---------------------------------------------------------------------------
# Tab 3: Distribusi & Korelasi
# ---------------------------------------------------------------------------
def tab_distribution(df: pd.DataFrame, roles: dict, key: str) -> None:
    num_cols = [c for c in df.columns if roles[c] == "numerik"]
    cat_cols = [c for c in df.columns if roles[c] == "kategori"]

    st.markdown("#### 📈 Sebaran kolom angka")
    if num_cols:
        col = st.selectbox("Pilih kolom angka", num_cols, key=f"{key}_dn")
        s = df[col].dropna()
        m = st.columns(5)
        for c, (label, val) in zip(m, [("Rata-rata", s.mean()), ("Median", s.median()), ("Std. deviasi", s.std()),
                                       ("Minimum", s.min()), ("Maksimum", s.max())]):
            c.metric(label, fmt(val))
        show_plot(px.histogram(df, x=col, nbins=30, marginal="box"), key=f"{key}_dhist", height=380)
    else:
        st.info("Tidak ada kolom angka pada tabel ini.")

    st.markdown("#### 🏷️ Komposisi kategori")
    if cat_cols:
        col = st.selectbox("Pilih kolom kategori", cat_cols, key=f"{key}_dc")
        vc = df[col].fillna("(Kosong)").astype(str).value_counts().head(20)
        t = pd.DataFrame({"Kategori": vc.index, "Jumlah": vc.to_numpy()})
        t["Persentase"] = (t["Jumlah"] / len(df) * 100).round(1)
        fig = px.bar(t, x="Jumlah", y="Kategori", orientation="h",
                     text=t["Persentase"].map(lambda p: f"{p:.1f}%"))
        fig.update_yaxes(autorange="reversed")
        show_plot(fig, key=f"{key}_dcat", height=max(300, 28 * len(t) + 80))
    else:
        st.info("Tidak ada kolom kategori yang terdeteksi.")

    st.markdown("#### 🔗 Korelasi antar kolom angka")
    if len(num_cols) >= 2:
        r = st.columns([3, 1])
        picked = r[0].multiselect("Kolom yang dibandingkan", num_cols, default=num_cols[:12], key=f"{key}_cc")
        method = r[1].radio("Metode", ["pearson", "spearman"], horizontal=True, key=f"{key}_cm")
        if len(picked) >= 2:
            corr = df[picked].corr(method=method)
            fig = px.imshow(corr, text_auto=".2f", color_continuous_scale="RdBu_r", zmin=-1, zmax=1, aspect="auto")
            show_plot(fig, key=f"{key}_cheat", height=max(380, 40 * len(picked) + 120))
            st.markdown("**Pasangan dengan korelasi terkuat**")
            show_df(sl.top_correlations(corr))
            st.caption("Korelasi menunjukkan keterkaitan, bukan sebab-akibat.")
        else:
            st.info("Pilih minimal 2 kolom.")
    else:
        st.info("Korelasi membutuhkan minimal 2 kolom angka.")


# ---------------------------------------------------------------------------
# Tab 4: Data & Unduh
# ---------------------------------------------------------------------------
def tab_data(df: pd.DataFrame, key: str, file_stem: str) -> None:
    c1, c2 = st.columns([2, 3])
    query = c1.text_input("🔎 Cari teks di tabel", key=f"{key}_q")
    shown = c2.multiselect("Kolom ditampilkan", list(df.columns), default=list(df.columns), key=f"{key}_cols")
    view = df[shown] if shown else df
    if query:
        mask = pd.Series(False, index=view.index)
        for c in view.columns:
            mask |= view[c].astype(str).str.contains(query, case=False, regex=False, na=False)
        view = view[mask]

    st.caption(f"{len(view):,} baris" + (f" — menampilkan {MAX_DISPLAY_ROWS:,} baris pertama" if len(view) > MAX_DISPLAY_ROWS else ""))
    show_df(view.head(MAX_DISPLAY_ROWS))

    sig = (int(pd.util.hash_pandas_object(view, index=False).sum()), tuple(view.columns))
    d1, d2, _ = st.columns([1, 1, 3])
    d1.download_button("📥 Unduh CSV", cached_csv(sig, view), f"{file_stem}.csv", "text/csv", key=f"{key}_csv")
    if len(view) <= EXCEL_MAX_ROWS:
        d2.download_button("📗 Unduh Excel", cached_xlsx(sig, view), f"{file_stem}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"{key}_xlsx")
    else:
        d2.caption(f"Excel hanya untuk ≤ {EXCEL_MAX_ROWS:,} baris; gunakan CSV.")


# ---------------------------------------------------------------------------
# Alur utama
# ---------------------------------------------------------------------------
def render() -> None:
    source = sidebar_upload()
    if source is None:
        render_welcome()
        return
    name, data = source
    file_id = hashlib.md5(data).hexdigest()[:10]

    try:
        sheets, notes = cached_load(name, data)
    except sl.SpreadsheetError as e:
        st.error(f"❌ {esc_md(e)}")
        st.stop()
    except Exception:
        st.error("❌ Terjadi kesalahan saat membaca file. Pastikan file tidak rusak, lalu coba lagi.")
        st.stop()

    usable = [n for n, d in sheets.items() if not d.dropna(how="all").empty]
    if not usable:
        st.warning("File ini tidak memiliki sheet yang berisi data.")
        st.stop()

    with st.sidebar:
        st.divider()
        sheet = st.selectbox("📑 Sheet", usable, key=f"sheet_{file_id}") if len(usable) > 1 else usable[0]
        fk = f"{file_id}_{sheet}"
        raw = sheets[sheet]
        guess = sl.guess_header_row(raw)
        with st.expander("⚙️ Pengaturan tabel"):
            header_row = guess
            if st.checkbox("Atur baris header manual", key=f"man_{fk}"):
                header_row = int(st.number_input("Baris header (1 = baris pertama)", 1, len(raw), guess + 1,
                                                 key=f"hdr_{fk}")) - 1
            else:
                st.caption(f"Header terdeteksi otomatis di baris {guess + 1}.")
            drop_summary = st.checkbox("Buang baris Total/Jumlah", True, key=f"sum_{fk}")
            drop_dups = st.checkbox("Buang baris duplikat", False, key=f"dup_{fk}")
            st.caption("Pratinjau baris awal file:")
            prev = raw.head(12)
            prev = prev.astype(object).where(prev.notna(), "").astype(str)
            prev.index = range(1, len(prev) + 1)
            show_df(prev, hide_index=False)

    df, removed = cached_build(file_id, sheet, header_row, drop_dups, drop_summary, raw)
    if df.empty or len(df.columns) == 0:
        st.warning("Tabel kosong setelah dibersihkan. Coba atur baris header secara manual di sidebar "
                   "(⚙️ Pengaturan tabel).")
        st.stop()
    roles = sl.column_roles(df)

    with st.sidebar:
        with st.expander("🔎 Filter data"):
            filterable = [c for c in df.columns if roles[c] in ("kategori", "numerik", "tanggal")]
            chosen = st.multiselect("Kolom yang difilter", filterable, key=f"fcols_{fk}")
            mask = pd.Series(True, index=df.index)
            for c in chosen:
                mask &= build_filter(df, c, roles[c], f"{fk}_{c}")
    dfv = df[mask]

    st.title(f"📊 {esc_md(name)}")
    st.caption(f"Sheet **{esc_md(sheet)}** · menampilkan {len(dfv):,} dari {len(df):,} baris · {len(df.columns)} kolom")
    for note in notes:
        st.warning(esc_md(note))
    if removed:
        st.caption(f"ℹ️ {removed} baris 'Total/Jumlah' dibuang otomatis (bisa dimatikan di ⚙️ Pengaturan tabel).")
    if dfv.empty:
        st.warning("Tidak ada baris yang cocok dengan filter. Longgarkan filter di sidebar.")
        st.stop()

    t1, t2, t3, t4 = st.tabs(["📌 Ringkasan", "📊 Studio Grafik", "📈 Distribusi & Korelasi", "📋 Data & Unduh"])
    with t1:
        tab_summary(dfv, roles, f"{fk}_t1")
    with t2:
        tab_studio(dfv, roles, f"{fk}_t2")
    with t3:
        tab_distribution(dfv, roles, f"{fk}_t3")
    with t4:
        stem = (name.rsplit(".", 1)[0] or "data") + "_hasil"
        tab_data(dfv, f"{fk}_t4", stem)
