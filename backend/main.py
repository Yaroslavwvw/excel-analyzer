from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from io import BytesIO
import pandas as pd
import numpy as np
from typing import List
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime

# Инициализация FastAPI приложения
app = FastAPI(
    title="Excel to Markdown API",
    description="API для обработки Excel файлов и генерации отчетов в формате Markdown",
    version="1.0.0"
)

@app.get("/")
def read_root():
    """Корневой маршрут, возвращающий информацию о сервисе"""
    return {
        "message": "Excel to Markdown API работает",
        "endpoints": {
            "/process-excel/": "Загрузка и обработка Excel-файла с генерацией отчета в формате Markdown"
        }
    }

@app.post("/process-excel/")
async def process_excel(file: UploadFile = File(...)):
    """
    Обрабатывает загруженный Excel-файл и возвращает отчет в формате Markdown
    """
    # Проверка формата файла
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(
            status_code=400,
            detail="Поддерживаются только файлы Excel (.xlsx, .xls)"
        )

    try:
        # Чтение содержимого файла
        contents = await file.read()
        buffer = BytesIO(contents)

        # Загрузка Excel-файла в DataFrame
        df = pd.read_excel(buffer)

        # Генерация отчета в формате Markdown на основе данных
        report = generate_markdown_report(df)

        # Создание байтового объекта для хранения отчета
        output = BytesIO(report.encode())
        output.seek(0)

        # Возвращаем отчет как скачиваемый файл
        filename = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.md"
        return StreamingResponse(
            output,
            media_type="text/markdown",
            headers={"Content-Disposition": f"attachment; filename={filename}"}
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка обработки файла: {str(e)}")

def generate_markdown_report(df: pd.DataFrame) -> str:
    """
    Генерирует отчет в формате Markdown на основе предоставленного DataFrame
    """
    # Создаем заголовок отчета
    report = f"# Отчет по анализу данных\n\n"
    report += f"Дата создания: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"

    # Основная информация о данных
    report += f"## Общая информация\n\n"
    report += f"- **Количество строк**: {df.shape[0]}\n"
    report += f"- **Количество столбцов**: {df.shape[1]}\n"
    report += f"- **Столбцы**: {', '.join(df.columns)}\n\n"

    # Статистика по числовым столбцам
    report += f"## Статистический анализ\n\n"

    # Проверяем наличие числовых столбцов
    numeric_columns = df.select_dtypes(include=['number']).columns.tolist()
    if numeric_columns:
        report += f"### Числовые данные\n\n"
        stats_df = df[numeric_columns].describe().transpose()
        # Форматируем статистику в виде таблицы Markdown
        stats_table = "| Столбец | Количество | Среднее | Ст. отклонение | Мин | 25% | 50% | 75% | Макс |\n"
        stats_table += "| --- | --- | --- | --- | --- | --- | --- | --- | --- |\n"

        for column, row in stats_df.iterrows():
            stats_table += f"| {column} | {row['count']:.0f} | {row['mean']:.2f} | {row['std']:.2f} | {row['min']:.2f} | {row['25%']:.2f} | {row['50%']:.2f} | {row['75%']:.2f} | {row['max']:.2f} |\n"

        report += stats_table + "\n\n"

    # Анализ категориальных данных
    categorical_columns = df.select_dtypes(include=['object', 'category']).columns.tolist()
    if categorical_columns:
        report += f"### Категориальные данные\n\n"

        for column in categorical_columns:
            value_counts = df[column].value_counts().head(5)  # Топ-5 значений
            report += f"#### {column}\n\n"

            # Создаем таблицу с частотами значений
            report += "| Значение | Количество | Процент |\n"
            report += "| --- | --- | --- |\n"

            for value, count in value_counts.items():
                percentage = (count / len(df)) * 100
                report += f"| {value} | {count} | {percentage:.2f}% |\n"

            report += "\n"

    # Анализ пропущенных значений
    report += f"## Анализ пропущенных значений\n\n"
    missing_values = df.isnull().sum()
    if missing_values.sum() > 0:
        report += "| Столбец | Пропущенные значения | Процент пропущенных |\n"
        report += "| --- | --- | --- |\n"

        for column, missing in missing_values.items():
            if missing > 0:
                percentage = (missing / len(df)) * 100
                report += f"| {column} | {missing} | {percentage:.2f}% |\n"

        report += "\n"
    else:
        report += "Пропущенные значения отсутствуют.\n\n"

    # Заключение
    report += "## Выводы\n\n"
    report += "На основе анализа данных можно сделать следующие выводы:\n\n"
    report += "1. Данные содержат информацию о " + str(df.shape[0]) + " записях с " + str(df.shape[1]) + " характеристиками.\n"

    if numeric_columns:
        # Находим столбец с наибольшим средним значением
        max_mean_column = df[numeric_columns].mean().idxmax()
        max_mean_value = df[numeric_columns].mean().max()
        report += f"2. Столбец '{max_mean_column}' имеет наибольшее среднее значение ({max_mean_value:.2f}).\n"

    if missing_values.sum() > 0:
        most_missing = missing_values.idxmax()
        most_missing_count = missing_values.max()
        report += f"3. Столбец '{most_missing}' имеет наибольшее количество пропущенных значений ({most_missing_count}).\n"

    report += "\n"

    return report