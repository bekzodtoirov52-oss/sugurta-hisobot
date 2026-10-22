import io
import os
import streamlit as st
import pandas as pd
from PIL import Image

# Logotip fayli nomi
LOGO_PATH = "logo.jpg"

if os.path.exists(LOGO_PATH):
    app_icon = Image.open(LOGO_PATH)
else:
    app_icon = "📊"

st.set_page_config(
    page_title="Bojxona Hisobot Tizimi",
    page_icon=app_icon,
    layout="wide"
)

# Yuqori qism: Logotip va Sarlavha
col_logo, col_title = st.columns([1, 10])
with col_logo:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=80)
with col_title:
    st.title("Davlat Bojxona Xizmati — Hisobotlar moduli")
    st.caption("OSAGO hisob-kitobi va Kuzatuv ma'lumotlarini tahlil qilish")

# Yordamchi hisoblash funksiyalari
def clean_numeric(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).replace(" ", "").replace("\xa0", "").replace(",", ".")
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def format_currency(val):
    return f"{val:,.2f}".replace(",", " ")

# TABLARNI YARATISH
tab_osago, tab_kuzatuv = st.tabs(["📑 OSAGO Hisoboti", "🔍 Kuzatuv (Monitoring)"])

# ==========================================
# 1-TAB: OSAGO HISOBOTI
# ==========================================
with tab_osago:
    st.subheader("OSAGO sug'urta hisobotini shakllantirish")
    uploaded_osago = st.file_uploader(
        "OSAGO faylini yuklang (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "csv"],
        key="osago_uploader"
    )

    OSAGO_COLUMNS = [
        "sugurta_firma", "declaration", "response_vaqti",
        "seria_polis", "number_polis", "ins_srok",
        "firstname", "lastname", "res_ins_prem"
    ]

    if uploaded_osago is not None:
        try:
            if uploaded_osago.name.endswith(".csv"):
                df_raw = pd.read_csv(uploaded_osago)
            else:
                df_raw = pd.read_excel(uploaded_osago)

            df_raw.columns = df_raw.columns.astype(str).str.strip().str.lower()
            missing = [c for c in OSAGO_COLUMNS if c not in df_raw.columns]

            if missing:
                st.error(f"Faylda quyidagi ustunlar topilmadi: {', '.join(missing)}")
            else:
                df = df_raw.copy()

                # Birlashtirishlar
                df["polis_full"] = (
                    df["seria_polis"].fillna("").astype(str).str.strip() + " " +
                    df["number_polis"].fillna("").astype(str).str.strip()
                ).str.strip()

                df["full_name"] = (
                    df["lastname"].fillna("").astype(str).str.strip() + " " +
                    df["firstname"].fillna("").astype(str).str.strip()
                ).str.strip()

                # Hisob-kitoblar (80% va 20%)
                df["res_ins_prem"] = df["res_ins_prem"].apply(clean_numeric)
                df["prem_80"] = (df["res_ins_prem"] * 0.80).round(2)
                df["prem_20"] = (df["res_ins_prem"] * 0.20).round(2)

                detail_cols = [
                    "sugurta_firma", "declaration", "response_vaqti",
                    "polis_full", "ins_srok", "full_name",
                    "res_ins_prem", "prem_80", "prem_20"
                ]
                df_details = df[detail_cols]

                # Umumiy statistika kartalari
                st.markdown("#### 📌 Umumiy ko‘rsatkichlar (Jami)")
                m1, m2, m3, m4, m5 = st.columns(5)
                m1.metric("Firmalar soni", df["sugurta_firma"].nunique())
                m2.metric("Polislar soni", f"{len(df):,}".replace(",", " "))
                m3.metric("Jami summa (100%)", format_currency(df["res_ins_prem"].sum()))
                m4.metric("Jami 80%", format_currency(df["prem_80"].sum()))
                m5.metric("Jami 20%", format_currency(df["prem_20"].sum()))

                st.markdown("---")

                # Firma bo'yicha alohida xisob-kitob jadvali
                st.markdown("#### 🏢 Sug'urta firmalari kesimida hisobot")
                summary_df = df.groupby("sugurta_firma", as_index=False).agg(
                    soni=("declaration", "count"),
                    jami_prem=("res_ins_prem", "sum"),
                    jami_80=("prem_80", "sum"),
                    jami_20=("prem_20", "sum")
                ).rename(columns={
                    "sugurta_firma": "Sug'urta firmasi",
                    "soni": "Polis soni",
                    "jami_prem": "Umumiy summa (100%)",
                    "jami_80": "Ulush 80%",
                    "jami_20": "Ulush 20%"
                }).sort_values(by="Polis soni", ascending=False)

                st.dataframe(
                    summary_df.style.format({
                        "Polis soni": "{:,}",
                        "Umumiy summa (100%)": "{:,.2f}",
                        "Ulush 80%": "{:,.2f}",
                        "Ulush 20%": "{:,.2f}"
                    }),
                    use_container_width=True,
                    hide_index=True
                )

                st.markdown("---")

                # Tafsilotlar jadvali va filtr
                st.markdown("#### 📋 Polislarning to'liq reestri")
                firm_list = ["Barcha firmalar"] + sorted(df["sugurta_firma"].dropna().unique().tolist())
                chosen_firm = st.selectbox("Sug'urta kompaniyasini tanlang:", firm_list)

                if chosen_firm != "Barcha firmalar":
                    view_df = df_details[df_details["sugurta_firma"] == chosen_firm]
                else:
                    view_df = df_details

                st.dataframe(
                    view_df.style.format({
                        "res_ins_prem": "{:,.2f}",
                        "prem_80": "{:,.2f}",
                        "prem_20": "{:,.2f}"
                    }),
                    use_container_width=True,
                    hide_index=True
                )

                # Excelga yuklab olish
                excel_buff = io.BytesIO()
                with pd.ExcelWriter(excel_buff, engine="openpyxl") as writer:
                    summary_df.to_excel(writer, sheet_name="Firma_Svodi", index=False)
                    df_details.to_excel(writer, sheet_name="Batafsil_Polislar", index=False)
                excel_buff.seek(0)

                st.download_button(
                    label="📥 OSAGO hisobotini Excelda yuklab olish",
                    data=excel_buff,
                    file_name="osago_hisobot.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Xatolik yuz berdi: {e}")
    else:
        st.info("Hisoblashni boshlash uchun OSAGO faylini yuklang.")

# ==========================================
# 2-TAB: KUZATUV (MONITORING)
# ==========================================
with tab_kuzatuv:
    st.subheader("Bojxona kuzatuvi (Monitoring) hisoboti")
    uploaded_kuzatuv = st.file_uploader(
        "Kuzatuv faylini yuklang (.xlsx, .xls, .csv)",
        type=["xlsx", "xls", "csv"],
        key="kuzatuv_uploader"
    )

    if uploaded_kuzatuv is not None:
        try:
            if uploaded_kuzatuv.name.endswith(".csv"):
                df_kuzatuv = pd.read_csv(uploaded_kuzatuv)
            else:
                df_kuzatuv = pd.read_excel(uploaded_kuzatuv)

            df_kuzatuv.columns = df_kuzatuv.columns.astype(str).str.strip().str.lower()

            st.success(f"Fayl muvaffaqiyatli yuklandi! Jami yozuvlar: {len(df_kuzatuv):,}".replace(",", " "))

            # Kuzatuv ma'lumotlari jadvali
            st.dataframe(df_kuzatuv, use_container_width=True)

            # Kuzatuv faylini ham qayta yuklab olish imkoniyati
            kuzatuv_buff = io.BytesIO()
            with pd.ExcelWriter(kuzatuv_buff, engine="openpyxl") as writer:
                df_kuzatuv.to_excel(writer, sheet_name="Kuzatuv", index=False)
            kuzatuv_buff.seek(0)

            st.download_button(
                label="📥 Kuzatuv ma'lumotlarini Excelda yuklab olish",
                data=kuzatuv_buff,
                file_name="kuzatuv_hisobot.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Kuzatuv faylini o‘qishda xatolik: {e}")
    else:
        st.info("Kuzatuv ma'lumotlarini ko‘rish uchun fayl yuklang.")
