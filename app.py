import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Ҳисобот кузатув", layout="wide")
st.title("📊 Страховой мониторинг: Ҳисобот кузатув")

uploaded_file = st.file_uploader("Загрузите Excel-файл реестра", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # Приведение названий колонок к нижнему регистру для надежности
    df.columns = [str(col).strip().lower() for col in df.columns]

    # Объединение полей
    df["polis_full"] = (
        df.get("seria_polis", "").astype(str).str.strip()
        + " "
        + df.get("number_polis", "").astype(str).str.strip()
    )
    df["full_name"] = (
        df.get("firstname", "").astype(str).str.strip()
        + " "
        + df.get("lastname", "").astype(str).str.strip()
    )

    # Приведение премии к числу и вычисление долей
    df["res_ins_prem"] = pd.to_numeric(
        df["res_ins_prem"].astype(str).str.replace(" ", "").str.replace(",", "."),
        errors="coerce",
    ).fillna(0)
    df["prem_75"] = df["res_ins_prem"] * 0.75
    df["prem_25"] = df["res_ins_prem"] * 0.25

    # Общие метрики
    total_policies = len(df)
    total_prem = df["res_ins_prem"].sum()
    total_75 = df["prem_75"].sum()
    total_25 = df["prem_25"].sum()

    st.subheader("📌 Общие итоги (Grand Total)")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Всего полисов", f"{total_policies:,}")
    col2.metric("Общая премия", f"{total_prem:,.2f}")
    col3.metric("Доля 75%", f"{total_75:,.2f}")
    col4.metric("Доля 25%", f"{total_25:,.2f}")

    # Сводка по каждой компании
    st.subheader("🏢 Итоги по каждой страховой компании")
    summary_df = (
        df.groupby("sugurta_firma")
        .agg(
            Полисы=("polis_full", "count"),
            Общая_премия=("res_ins_prem", "sum"),
            Доля_75=("prem_75", "sum"),
            Доля_25=("prem_25", "sum"),
        )
        .reset_index()
    )

    st.dataframe(
        summary_df.style.format(
            {
                "Полисы": "{:,}",
                "Общая_премия": "{:,.2f}",
                "Доля_75": "{:,.2f}",
                "Доля_25": "{:,.2f}",
            }
        ),
        use_container_width=True,
    )

    # Детальная таблица
    st.subheader("📋 Детальный реестр")
    selected_cols = [
        "sugurta_firma",
        "declaration",
        "polis_full",
        "full_name",
        "response_vaqti",
        "ins_srok",
        "res_ins_prem",
        "prem_75",
        "prem_25",
    ]
    existing_cols = [c for c in selected_cols if c in df.columns]

    filter_company = st.selectbox(
        "Фильтр по компании", ["Все"] + list(df["sugurta_firma"].dropna().unique())
    )
    view_df = (
        df
        if filter_company == "Все"
        else df[df["sugurta_firma"] == filter_company]
    )

    st.dataframe(
        view_df[existing_cols].style.format(
            {"res_ins_prem": "{:,.2f}", "prem_75": "{:,.2f}", "prem_25": "{:,.2f}"}
        ),
        use_container_width=True,
    )

    # Экспорт в Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Сводка_по_фирмам", index=False)
        view_df[existing_cols].to_excel(
            writer, sheet_name="Реестр", index=False
        )

    st.download_button(
        label="📥 Скачать готовый отчет в Excel",
        data=output.getvalue(),
        file_name="xisobot_kuzatuv_hisoblangan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
