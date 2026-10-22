import os
import streamlit as st

st.set_page_config(page_title="Sug'urta Portali", layout="wide")

col_logo, col_title = st.columns([1, 6])
with col_logo:
    if os.path.exists("logo.png"):
        st.image("logo.png", width=100)
    else:
        st.write("🏛️")

with col_title:
    st.title("Sug'urta monitoringi va hisoboti")
    st.caption("Bo'limni tanlash uchun chapdagi menyudan foydalaning.")

st.markdown("---")
st.info("""
### Xush kelibsiz!
Kerakli modulni chap menyudan (Sidebar) tanlang:
- **🔍 Kuzatuv:** 75% va 25% ulushlar hisob-kitobi.
- **🚗 OSAGO:** 80% va 20% ulushlar hisob-kitobi.

Har bir sahifa alohida ishlaydi, hisoblar va fayllar bir-biriga xalaqit bermaydi.
""")
