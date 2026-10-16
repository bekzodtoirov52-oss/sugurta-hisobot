import streamlit as st
import pandas as pd
import io

# Настройка страницы
st.set_page_config(
    page_title="Sug'urta Polislari Hisoboti",
    page_icon="📋",
    layout="wide"
)

st.title("📋 Har bir polis va uning summasi bo'yicha hisobot")
st.write("Excel faylni yuklang. Dastur har bir polisning summasi va uning 75% hamda 25% qismlarini chiqarib, eng oxirida umumiy JAMI summani hisoblab beradi.")

# Боковое меню
st.sidebar.header("Fayl va Filtrlar")
uploaded_file = st.sidebar.file_uploader("Excel faylni tanlang (.xlsx, .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Чтение Excel
        df = pd.read_excel(uploaded_file)
        
        # Приведение колонок к нижнему регистру и удаление лишних пробелов
        df.columns = df.columns.astype(str).str.strip().str.lower()
        
        # Проверка обязательных колонок
        required_cols = ['sugurta_firma', 'res_ins_prem', 'polis']
        missing_cols = [c for c in required_cols if c not in df.columns]
        
        if missing_cols:
            st.error(f"Xatolik! Excel jadvalida quyidagi majburiy ustunlar topilmadi: {', '.join(missing_cols)}")
            st.info("Kutilayotgan ustunlar: sugurta_firma, declaration, polis, response_vaqti, res_ins_prem")
        else:
            # Очистка числового поля суммы премии
            df['res_ins_prem_clean'] = (
                df['res_ins_prem']
                .astype(str)
                .str.replace(r'\s+', '', regex=True)
                .str.replace(',', '.')
            )
            df['res_ins_prem_clean'] = pd.to_numeric(df['res_ins_prem_clean'], errors='coerce').fillna(0)

            # Расчет долей 75% и 25% для каждого отдельного полиса
            df['summa_75'] = df['res_ins_prem_clean'] * 0.75
            df['summa_25'] = df['res_ins_prem_clean'] * 0.25

            # Фильтр по страховым фирмам в сайдбаре
            companies = ['Barchasi (Hammasi)'] + sorted(df['sugurta_firma'].dropna().unique().tolist())
            selected_company = st.sidebar.selectbox("Kompaniya bo'yicha saralash:", companies)

            # Применение фильтра
            if selected_company != 'Barchasi (Hammasi)':
                filtered_df = df[df['sugurta_firma'] == selected_company].copy()
            else:
                filtered_df = df.copy()

            # Формирование колонок для итогового отображения
            cols_to_show = []
            rename_dict = {}

            # Добавляем колонки, если они присутствуют в файле
            if 'sugurta_firma' in filtered_df.columns:
                cols_to_show.append('sugurta_firma')
                rename_dict['sugurta_firma'] = 'Sug‘urta firmasi'
            if 'declaration' in filtered_df.columns:
                cols_to_show.append('declaration')
                rename_dict['declaration'] = 'Deklaratsiya'
            if 'polis' in filtered_df.columns:
                cols_to_show.append('polis')
                rename_dict['polis'] = 'Polis raqami'
            if 'response_vaqti' in filtered_df.columns:
                cols_to_show.append('response_vaqti')
                rename_dict['response_vaqti'] = 'Vaqti'

            # Финансовые колонки
            cols_to_show.extend(['res_ins_prem_clean', 'summa_75', 'summa_25'])
            rename_dict['res_ins_prem_clean'] = 'Polis summasi (100%)'
            rename_dict['summa_75'] = 'Ulush (75%)'
            rename_dict['summa_25'] = 'Ulush (25%)'

            detail_df = filtered_df[cols_to_show].rename(columns=rename_dict)

            # Подсчет общих сумм
            total_100 = detail_df['Polis summasi (100%)'].sum()
            total_75 = detail_df['Ulush (75%)'].sum()
            total_25 = detail_df['Ulush (25%)'].sum()
            total_count = len(detail_df)

            # Верхние карточки с ключевыми метриками
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Polislar soni", f"{total_count:,}".replace(',', ' '))
            col2.metric("Jami summa (100%)", f"{total_100:,.2f}".replace(',', ' '))
            col3.metric("Jami 75%", f"{total_75:,.2f}".replace(',', ' '))
            col4.metric("Jami 25%", f"{total_25:,.2f}".replace(',', ' '))

            st.write("---")

            # Формирование финальной строки «JAMI / ITOGO»
            total_row_data = {col: '' for col in detail_df.columns}
            
            # Назначаем надпись "JAMI / ITOGO" в первую текстовую колонку
            if 'Sug‘urta firmasi' in detail_df.columns:
                total_row_data['Sug‘urta firmasi'] = 'JAMI / ITOGO'
            elif 'Polis raqami' in detail_df.columns:
                total_row_data['Polis raqami'] = 'JAMI / ITOGO'

            total_row_data['Polis summasi (100%)'] = total_100
            total_row_data['Ulush (75%)'] = total_75
            total_row_data['Ulush (25%)'] = total_25

            total_row = pd.DataFrame([total_row_data])

            # Добавляем строку итогов в самый конец таблицы
            final_table = pd.concat([detail_df, total_row], ignore_index=True)

            # Таблица для красивого отображения с форматированием чисел
            display_table = final_table.copy()
            for col in ['Polis summasi (100%)', 'Ulush (75%)', 'Ulush (25%)']:
                display_table[col] = display_table[col].apply(
                    lambda x: f"{x:,.2f}".replace(',', ' ') if pd.notnull(x) and isinstance(x, (int, float)) else str(x)
                )

            st.subheader(f"📑 Har bir polis bo'yicha to'liq jadval ({selected_company})")
            st.dataframe(display_table, use_container_width=True)

            # Кнопка скачивания итоговой детальной таблицы
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_table.to_excel(writer, index=False, sheet_name="Polislar_Hisoboti")
            output.seek(0)

            st.download_button(
                label="📥 Jadvalni Excel formatida yuklab olish",
                data=output,
                file_name="polislar_va_jami_summa.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Дополнительный блок: Сводка по компаниям
            with st.expander("🏢 Kompaniyalar bo'yicha qisqacha xulosa (Svodka)"):
                summary_df = df.groupby('sugurta_firma', as_index=False).agg(
                    Polislar_soni=('polis', 'nunique'),
                    Jami_100=('res_ins_prem_clean', 'sum')
                )
                summary_df['Ulush_75'] = summary_df['Jami_100'] * 0.75
                summary_df['Ulush_25'] = summary_df['Jami_100'] * 0.25
                st.dataframe(summary_df, use_container_width=True)

    except Exception as e:
        st.error(f"Faylni qayta ishlashda xatolik: {e}")
else:
    st.info("Iltimos, hisoblashni boshlash uchun chap menyudan Excel faylni yuklang.")
