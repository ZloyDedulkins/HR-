import pandas as pd
import streamlit as st

st.title("📊 Расчет текучести персонала")

file = st.file_uploader("Загрузи Excel файл", type=["xlsx"])

if file:
    try:
        # Чтение данных
        fired = pd.read_excel(file, sheet_name=0)
        staff = pd.read_excel(file, sheet_name=1)
        exclude = pd.read_excel(file, sheet_name=2)

        # Проверка колонок
        required_fired = ['ФИО', 'подразделение', 'дата увольнения']
        required_staff = ['подразделение', 'штат']
        required_exclude = ['ФИО']

        for col in required_fired:
            if col not in fired.columns:
                st.error(f"❌ В листе 'Уволенные' нет колонки: {col}")
                st.stop()

        for col in required_staff:
            if col not in staff.columns:
                st.error(f"❌ В листе 'Штатка' нет колонки: {col}")
                st.stop()

        for col in required_exclude:
            if col not in exclude.columns:
                st.error(f"❌ В листе 'Исключения' нет колонки: {col}")
                st.stop()

        # Удаляем исключенных
        fired_clean = fired[~fired['ФИО'].isin(exclude['ФИО'])]

        # Группировка
        result = (
            fired_clean.groupby('подразделение')
            .size()
            .reset_index(name='Уволенные')
        )

        # Объединение со штаткой
        result = result.merge(staff, on='подразделение', how='left')

        # Проверка на пропуски
        if result['штат'].isna().any():
            st.warning("⚠️ Есть подразделения без штатной численности")

        # Расчет текучести
        result['Текучесть %'] = (result['Уволенные'] / result['штат']) * 100
        result['Текучесть %'] = result['Текучесть %'].round(2)

        # Вывод
        st.subheader("Результат")
        st.dataframe(result)

        # Скачать
        output = result.to_excel(index=False, engine='openpyxl')

        st.download_button(
            label="📥 Скачать результат",
            data=output,
            file_name="turnover_result.xlsx"
        )

    except Exception as e:
        st.error(f"Ошибка обработки файла: {e}")
