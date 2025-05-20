import streamlit as st
import pandas as pd
import requests
import io
from urllib.parse import urljoin

# Конфигурация приложения
st.set_page_config(
    page_title="Excel Analyzer",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# URL бэкенда на render.com
# Замените на свой URL после развертывания на render.com
BACKEND_URL = "https://excel-to-markdown-api.onrender.com"

# Функции для взаимодействия с бэкендом
def check_api_status():
    """Проверяем доступность API"""
    try:
        response = requests.get(BACKEND_URL, timeout=10)
        return response.status_code == 200
    except requests.RequestException:
        return False

def process_excel(file):
    """Отправляем Excel-файл на обработку в API и получаем Markdown-отчет"""
    url = urljoin(BACKEND_URL, "/process-excel/")
    files = {"file": file}
    try:
        response = requests.post(url, files=files)
        if response.status_code == 200:
            return response.content.decode('utf-8')
        else:
            st.error(f"Ошибка при обработке файла: {response.text}")
            return None
    except requests.RequestException as e:
        st.error(f"Ошибка соединения с API: {str(e)}")
        return None

# Интерфейс приложения
def main():
    st.title("📊 Анализатор Excel файлов")
    st.markdown("""
    Этот инструмент позволяет загрузить Excel-файл и получить аналитический отчет
    в формате Markdown. Просто загрузите файл и нажмите кнопку "Анализировать".
    """)

    # Проверка статуса API
    if not check_api_status():
        st.error("⚠️ Не удалось подключиться к API. Пожалуйста, проверьте соединение или попробуйте позже.")
        return

    # Загрузка файла
    uploaded_file = st.file_uploader("Выберите Excel файл", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        # Показываем предварительный просмотр данных
        try:
            df = pd.read_excel(uploaded_file)
            st.subheader("Предварительный просмотр данных")
            st.dataframe(df.head(5))

            # Получение основных статистик для информации
            st.subheader("Базовая информация")
            col1, col2, col3 = st.columns(3)
            col1.metric("Строки", df.shape[0])
            col2.metric("Столбцы", df.shape[1])
            col3.metric("Пропущенные значения", df.isna().sum().sum())

            # Сбрасываем указатель файла для повторного чтения
            uploaded_file.seek(0)

            if st.button("Анализировать"):
                with st.spinner("Обрабатываем данные..."):
                    markdown_report = process_excel(uploaded_file)

                if markdown_report:
                    st.success("Отчет успешно создан!")

                    # Показываем отчет в интерфейсе
                    st.subheader("Отчет")
                    st.markdown(markdown_report)

                    # Предоставляем возможность скачать отчет
                    st.download_button(
                        label="Скачать отчет",
                        data=markdown_report,
                        file_name="report.md",
                        mime="text/markdown",
                    )

        except Exception as e:
            st.error(f"Ошибка при чтении файла: {str(e)}")

if __name__ == "__main__":
    main()