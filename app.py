import io
import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Хисобот ОСАГО",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Автоматический расчет отчетов ОСАГО")
st.caption("Группировка по страховым компаниям, сплит 80/20 и экспорт в Excel")

uploaded_file = st.file_uploader(
    "Загрузите исходный файл отчета (.xlsx, .xls, .csv)",
    type=["xlsx", "xls", "csv"]
)

REQUIRED_COLUMNS = [
    "sugurta_firma",
    "declaration",
    "response_vaqti",
    "seria_polis",
    "number_polis",
    "ins_srok",
    "firstname",
    "lastname",
    "res_ins_prem"
]

def clean_numeric(val):
    """Очищает строку от пробелов и спецсимволов и переводит в float."""
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
    """Форматирует число с разделением тысяч."""
    return f"{val:,.2f}".replace(",", " ")

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df_raw = pd.read_csv(uploaded_file)
        else:
            df_raw = pd.read_excel(uploaded_file)

        # Нормализация названий колонок (нижний регистр, убираем пробелы)
        df_raw.columns = df_raw.columns.astype(str).str.strip().str.lower()

        missing_cols = [c for c in REQUIRED_COLUMNS if c not in df_raw.columns]
        if missing_cols:
            st.error(f"В загруженном файле отсутствуют обязательные колонки: `{', '.join(missing_cols)}`")
            st.stop()

        df = df_raw.copy()

        # 1. Объединение серии и номера полиса
        df["polis_full"] = (
            df["seria_polis"].fillna("").astype(str).str.strip() + " " +
            df["number_polis"].fillna("").astype(str).str.strip()
        ).str.strip()

        # 2. Объединение Фамилии и Имени
        df["full_name"] = (
            df["lastname"].fillna("").astype(str).str.strip() + " " +
            df["firstname"].fillna("").astype(str).str.strip()
        ).str.strip()

        # 3. Приведение суммы к float и расчет 80% / 20%
        df["res_ins_prem"] = df["res_ins_prem"].apply(clean_numeric)
        df["prem_80"] = (df["res_ins_prem"] * 0.80).round(2)
        df["prem_20"] = (df["res_ins_prem"] * 0.20).round(2)

        # Выравнивание итогового порядка колонок для детальной таблицы
        detail_cols = [
            "sugurta_firma",
            "declaration",
            "response_vaqti",
            "polis_full",
            "ins_srok",
            "full_name",
            "res_ins_prem",
            "prem_80",
            "prem_20"
        ]
        df_details = df[detail_cols]

        # --- БЛОК 1: ОБЩИЕ ИТОГОВЫЕ МЕТРИКИ ---
        st.subheader("📌 Общие показатели (Итого)")
        total_policies = len(df)
        total_companies = df["sugurta_firma"].nunique()
        total_prem = df["res_ins_prem"].sum()
        total_80 = df["prem_80"].sum()
        total_20 = df["prem_20"].sum()

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Компаний", total_companies)
        m2.metric("Всего полисов", f"{total_policies:,}".replace(",", " "))
        m3.metric("Общая сумма (100%)", format_currency(total_prem))
        m4.metric("Доля 80%", format_currency(total_80))
        m5.metric("Доля 20%", format_currency(total_20))

        st.markdown("---")

        # --- БЛОК 2: СВОДНАЯ СТАТИСТИКА ПО КОМПАНИЯМ ---
        st.subheader("🏢 Сводка по страховым компаниям")

        summary_df = df.groupby("sugurta_firma", as_index=False).agg(
            soni=("declaration", "count"),
            jami_prem=("res_ins_prem", "sum"),
            jami_80=("prem_80", "sum"),
            jami_20=("prem_20", "sum")
        ).rename(columns={
            "sugurta_firma": "Страховая компания",
            "soni": "Кол-во полисов",
            "jami_prem": "Общая сумма (100%)",
            "jami_80": "Сумма 80%",
            "jami_20": "Сумма 20%"
        }).sort_values(by="Кол-во полисов", ascending=False)

        # Отображение сводной таблицы с форматированием
        st.dataframe(
            summary_df.style.format({
                "Кол-во полисов": "{:,}",
                "Общая сумма (100%)": "{:,.2f}",
                "Сумма 80%": "{:,.2f}",
                "Сумма 20%": "{:,.2f}"
            }),
            use_container_width=True,
            hide_index=True
        )

        st.markdown("---")

        # --- БЛОК 3: ДЕТАЛИЗИРОВАННАЯ ТАБЛИЦА С ФИЛЬТРОМ ---
        st.subheader("📋 Детализированный реестр полисов")

        all_firms = ["Все компании"] + sorted(df["sugurta_firma"].dropna().unique().tolist())
        selected_firm = st.selectbox("Фильтр по страховой компании:", all_firms)

        if selected_firm != "Все компании":
            filtered_df = df_details[df_details["sugurta_firma"] == selected_firm]
        else:
            filtered_df = df_details

        st.write(f"Отображено записей: **{len(filtered_df)}**")
        st.dataframe(
            filtered_df.style.format({
                "res_ins_prem": "{:,.2f}",
                "prem_80": "{:,.2f}",
                "prem_20": "{:,.2f}"
            }),
            use_container_width=True,
            hide_index=True
        )

        # --- БЛОК 4: ЭКСПОРТ В EXCEL ---
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            summary_df.to_excel(writer, sheet_name="Сводка_Компании", index=False)
            df_details.to_excel(writer, sheet_name="Детализация", index=False)

        excel_buffer.seek(0)

        st.download_button(
            label="📥 Скачать полный отчет в Excel (.xlsx)",
            data=excel_buffer,
            file_name="xisobot_osago_hisoblangan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Произошла ошибка при обработке файла: {e}")
else:
    st.info("Пожалуйста, загрузите файл с отчетом для начала расчета.")
