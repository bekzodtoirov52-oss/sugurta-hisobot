import streamlit as st
import pandas as pd
import io

# Sahifa sozlamalari
st.set_page_config(
    page_title="Sug'urta Polis Hisoboti",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Sug'urta polislarini hisoblash tizimi")
st.write("Excel faylni yuklang. Dastur har bir sug'urta firmasi bo'yicha polislar soni va summalarini (100%, 75%, 25%) hisoblab beradi.")

# Fayl yuklash bo'limi
uploaded_file = st.file_uploader("Excel faylni tanlang (.xlsx yoki .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Excelni o'qish
        df = pd.read_excel(uploaded_file)
        
        # Ustun nomlarini kichik harfga o'tkazish va bo'sh joylarni tozalash
        df.columns = df.columns.astype(str).str.strip().str.lower()
        
        required_cols = ['sugurta_firma', 'res_ins_prem']
        missing_cols = [c for c in required_cols if c not in df.columns]
        
        if missing_cols:
            st.error(f"Xatolik! Excel jadvalida quyidagi ustunlar topilmadi: {', '.join(missing_cols)}")
            st.info("Kutilayotgan ustunlar: sugurta_firma, res_ins_prem, polis")
        else:
            # res_ins_prem ustunidagi sonlarni tozalash (probellar yoki vergullarni to'g'rilash)
            df['res_ins_prem_clean'] = (
                df['res_ins_prem']
                .astype(str)
                .str.replace(r'\s+', '', regex=True)
                .str.replace(',', '.')
            )
            df['res_ins_prem_clean'] = pd.to_numeric(df['res_ins_prem_clean'], errors='coerce').fillna(0)

            # Polis ustunini aniqlash
            polis_col = 'polis' if 'polis' in df.columns else df.columns[0]

            # Firmalar bo'yicha guruhlash
            summary_df = df.groupby('sugurta_firma', as_index=False).agg(
                polis_soni=(polis_col, 'nunique'),
                summa_100=('res_ins_prem_clean', 'sum')
            )

            # 75% va 25% hisoblash
            summary_df['summa_75'] = summary_df['summa_100'] * 0.75
            summary_df['summa_25'] = summary_df['summa_100'] * 0.25

            # Ustun nomlarini chiroyli qilish
            summary_df = summary_df.rename(columns={
                'sugurta_firma': 'Sug‘urta firmasi',
                'polis_soni': 'Polislar soni (dona)',
                'summa_100': 'Jami summa (100%)',
                'summa_75': 'Ulush (75%)',
                'summa_25': 'Ulush (25%)'
            })

            # JAMI (ITOGO) qatorini yasash
            total_row = pd.DataFrame([{
                'Sug‘urta firmasi': 'JAMI / ITOGO',
                'Polislar soni (dona)': summary_df['Polislar soni (dona)'].sum(),
                'Jami summa (100%)': summary_df['Jami summa (100%)'].sum(),
                'Ulush (75%)': summary_df['Ulush (75%)'].sum(),
                'Ulush (25%)': summary_df['Ulush (25%)'].sum()
            }])

            # Asosiy jadvalga Jami qatorini qo'shish
            final_df = pd.concat([summary_df, total_row], ignore_index=True)

            # Yuqori ko'rsatkichlar (Vidjetlar)
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Jami polislar", f"{total_row['Polislar soni (dona)'][0]:,}".replace(',', ' '))
            col2.metric("Umumiy summa (100%)", f"{total_row['Jami summa (100%)'][0]:,.2f}".replace(',', ' '))
            col3.metric("75% qismi", f"{total_row['Ulush (75%)'][0]:,.2f}".replace(',', ' '))
            col4.metric("25% qismi", f"{total_row['Ulush (25%)'][0]:,.2f}".replace(',', ' '))

            st.write("---")
            st.subheader("📋 Hisoblangan natija jadvali")

            # Ekranda chiroyli formatda ko'rsatish
            display_df = final_df.copy()
            for col in ['Jami summa (100%)', 'Ulush (75%)', 'Ulush (25%)']:
                display_df[col] = display_df[col].apply(lambda x: f"{x:,.2f}".replace(',', ' '))
            display_df['Polislar soni (dona)'] = display_df['Polislar soni (dona)'].apply(lambda x: f"{x:,}".replace(',', ' '))

            st.dataframe(display_df, use_container_width=True)

            # Tayyor Excel fayl qilib yuklab olish tugmasi
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name="Hisobot")
            output.seek(0)

            st.download_button(
                label="📥 Natijani Excel fayl qilib yuklab olish",
                data=output,
                file_name="sugurta_hisobot_natija.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Faylni o'qishda xatolik yuz berdi: {e}")
else:
    st.info("Iltimos, hisoblash uchun Excel faylni yuklang.")
