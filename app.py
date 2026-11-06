"""
app.py — pintu masuk aplikasi (menu navigasi antar fitur).

Menu:
  📊 Visualisasi Spreadsheet  -> views/visualisasi.py
  🎓 AUTOCert                 -> views/autocert.py
  ✅ Cek Absensi              -> views/cek_absensi.py

Jalankan:  python -m streamlit run app.py
"""
import streamlit as st

st.set_page_config(page_title="Visualisasi Spreadsheet · AUTOCert", page_icon="📊", layout="wide")

# CSS statis saja (tidak ada konten pengguna yang disisipkan ke HTML).
st.markdown(
    """
    <style>
      .block-container { padding-top: 1.6rem; }
      div[data-testid="stMetric"] {
          background: rgba(127,127,127,.08); border-radius: 10px; padding: 10px 14px;
          border-left: 4px solid #2563EB;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

from views import autocert, cek_absensi, visualisasi  # noqa: E402  (diimpor setelah set_page_config)

pages = [
    st.Page(visualisasi.render, title="Visualisasi Spreadsheet", icon="📊", url_path="visualisasi", default=True),
    st.Page(autocert.render, title="AUTOCert", icon="🎓", url_path="autocert"),
    st.Page(cek_absensi.render, title="Cek Absensi", icon="✅", url_path="cek-absensi"),
]
st.navigation(pages).run()
