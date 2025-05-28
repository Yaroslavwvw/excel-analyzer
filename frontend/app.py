# frontend/app.py

import streamlit as st
import requests
import pandas as pd
import io

BACKEND_URL = "https://backservice-l6s9.onrender.com/convert"

st.set_page_config(page_title="🌍 Конвертер координат", layout="centered")

st.title("🌍 Конвертер координат между системами")
st.markdown("Загрузите Excel-файл с координатами и выберите системы для преобразования.")

# Загрузка файла
uploaded_file = st.file_uploader("Выберите Excel-файл (.xlsx)", type=["xlsx", "xls"])

# Выбор систем
systems = ["СК-42", "СК-95", "ПЗ-90", "ПЗ-90.02", "ПЗ-90.11", "WGS-84", "ITRF-2008"]
from_system = st.selectbox("Исходная система:", systems)
to_system = st.selectbox("Целевая система:", ["ГСК-2011"])

# Кнопка запуска преобразования
if uploaded_file and st.button("🚀 Выполнить преобразование"):
    with st.spinner("Преобразование данных... Это может занять несколько секунд"):
        try:
            files = {"file": (uploaded_file.name, uploaded_file.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
            data = {"from_system": from_system, "to_system": to_system}

            response = requests.post(BACKEND_URL, data=data, files=files)

            if response.status_code == 200:
                result = response.json()
                
                # Отображение отчета
                st.markdown("### 📄 Отчет о преобразовании:")
                st.markdown(result["markdown"])
                
                # Кнопка загрузки
                st.download_button(
                    label="📥 Скачать отчет в Markdown",
                    data=result["markdown"],
                    file_name=result["filename"],
                    mime="text/markdown"
                )
            else:
                error = response.json().get("detail", "Неизвестная ошибка")
                st.error(f"❌ Ошибка при обработке данных: {error}")

        except Exception as e:
            st.error(f"⚠️ Произошла ошибка: {str(e)}")