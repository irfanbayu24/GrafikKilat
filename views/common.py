"""views/common.py — helper yang dipakai bersama oleh semua menu."""
import inspect
import re

import streamlit as st


def stretch_kwargs(fn) -> dict:
    """Kompatibel dengan Streamlit lama (use_container_width) maupun baru (width="stretch")."""
    try:
        params = inspect.signature(fn).parameters
    except (TypeError, ValueError):
        params = {}
    return {"width": "stretch"} if "width" in params else {"use_container_width": True}


PLOT_KW = stretch_kwargs(st.plotly_chart)
DF_KW = stretch_kwargs(st.dataframe)


def esc_md(text) -> str:
    """Escape markdown: nama file/kolom/sheet berasal dari pengguna dan tidak boleh diinterpretasi sebagai format/tautan."""
    return re.sub(r"([\\`*_{}\[\]()#+\-.!|<>~$:])", r"\\\1", str(text))


def remembered_upload(slot: str, label: str, types: list, container=None, help: str = None):
    """file_uploader yang mengingat berkas terakhir selama sesi.

    Streamlit membuang state widget saat pengguna pindah menu; tanpa ini berkas harus diunggah ulang
    setiap kali kembali ke menu. Mengembalikan (nama, bytes) atau None.
    """
    box = container if container is not None else st
    uploaded = box.file_uploader(label, type=types, key=f"up_{slot}", help=help)
    if uploaded is not None:
        st.session_state[f"mem_{slot}"] = (uploaded.name, uploaded.getvalue())
    remembered = st.session_state.get(f"mem_{slot}")
    if remembered and uploaded is None:
        box.caption(f"📎 Berkas tersimpan: **{esc_md(remembered[0])}**")
        if box.button("Hapus berkas tersimpan", key=f"clr_{slot}"):
            st.session_state.pop(f"mem_{slot}", None)
            st.rerun()
    return remembered
