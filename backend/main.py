# backend/main.py

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import JSONResponse
import pandas as pd
import numpy as np
import json
import sympy as sp
import io
from math import radians
from sympy import Matrix, symbols, Eq, pi
from sympy.printing.latex import latex
from fastapi.middleware.cors import CORSMiddleware
import base64

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

with open("parameters.json", "r") as f:
    parameters = json.load(f)

# Символы
Xe, Ye, Ze = sp.symbols('X_e Y_e Z_e')
Xs, Ys, Zs = sp.symbols('X_s Y_s Z_s')
m, wx, wy, wz = sp.symbols('m ω_x ω_y ω_z')
dX, dY, dZ = sp.symbols('dX dY dZ')

def create_formula_matrix():
    """Создает матрицы преобразования"""
    to_GSK_matrix = Matrix([
        [1, wz, -wy],
        [-wz, 1, wx],
        [wy, -wx, 1]
    ])
    
    from_GSK_matrix = Matrix([
        [1, -wz, wy],
        [wz, 1, -wx],
        [-wy, wx, 1]
    ])
    
    return to_GSK_matrix, from_GSK_matrix

def generate_formula_latex(matrix, symbol_X, symbol_Y, symbol_Z, m_expr, dX, dY, dZ):
    """Генерирует LaTeX-формулу для отчёта"""
    try:
        formula = Eq(
            Matrix([Xe, Ye, Ze]),
            (1 + m_expr) * matrix * Matrix([symbol_X, symbol_Y, symbol_Z]) + Matrix([dX, dY, dZ])
        )
        return f"$$ {latex(formula)} $$"
    except Exception as e:
        print(f"Ошибка генерации формулы: {str(e)}")
        return ""

def create_markdown_report(start_system, end_system, start_df, transformed_df, parameters):
    """Создает отчет в формате Markdown"""
    report = "# Отчет по преобразованию координат\n\n"
    report += "## Общая формула по которой производились вычисления\n\n"
    
    to_GSK_matrix, from_GSK_matrix = create_formula_matrix()
    
    if start_system != "ГСК-2011":
        report += "### Формула для перевода в систему ГСК\n"
        formula = generate_formula_latex(to_GSK_matrix, Xs, Ys, Zs, 'm', dX, dY, dZ)
        report += formula + "\n"
    
    if end_system != "ГСК-2011":
        report += "### Формула для перевода из системы ГСК\n"
        formula = generate_formula_latex(from_GSK_matrix, Xs, Ys, Zs, 'm', dX, dY, dZ)
        report += formula + "\n"
    
    report += "## Формулы с подставленными параметрами\n\n"
    
    if start_system != "ГСК-2011":
        p_start = parameters[start_system]
        wz_val = round(radians(p_start['wz']/3600)*pi/180, 10)
        wy_val = round(radians(p_start['wy']/3600)*pi/180, 10)
        wx_val = round(radians(p_start['wx']/3600)*pi/180, 10)
        
        matrix = Matrix([
            [1, wz_val, -wy_val],
            [-wz_val, 1, wx_val],
            [wy_val, -wx_val, 1]
        ])
        
        m_val = p_start['m']/(10**6)
        formula = generate_formula_latex(matrix, Xs, Ys, Zs, m_val, p_start['dX'], p_start['dY'], p_start['dZ'])
        report += f"### Формула для перевода {start_system} в ГСК\n"
        report += formula + "\n"
    
    if end_system != "ГСК-2011":
        p_end = parameters[end_system]
        wz_val = round(radians(p_end['wz']/3600)*pi/180, 10)
        wy_val = round(radians(p_end['wy']/3600)*pi/180, 10)
        wx_val = round(radians(p_end['wx']/3600)*pi/180, 10)
        
        matrix = Matrix([
            [1, -wz_val, wy_val],
            [wz_val, 1, -wx_val],
            [-wy_val, wx_val, 1]
        ])
        
        m_val = p_end['m']/(10**6)
        formula = generate_formula_latex(matrix, Xs, Ys, Zs, m_val, p_end['dX'], p_end['dY'], p_end['dZ'])
        report += f"### Формула для перевода ГСК в {end_system}\n"
        report += formula + "\n"
    
    # Добавляем таблицы
    report += "## Исходные данные\n"
    report += "| X | Y | Z |\n| --- | --- | --- |\n"
    for _, row in start_df.iterrows():
        report += f"| {row['X']} | {row['Y']} | {row['Z']} |\n"
    
    report += "\n## Преобразованные данные\n"
    report += "| X | Y | Z |\n| --- | --- | --- |\n"
    for _, row in transformed_df.iterrows():
        report += f"| {row['X']} | {row['Y']} | {row['Z']} |\n"
    
    # Добавляем вывод
    report += "\n## Вывод\n"
    report += "Процесс преобразования координат был успешно выполнен, с результатами, представленными выше."
    
    return report

@app.post("/convert")
async def convert(
    file: UploadFile = File(...),
    from_system: str = "СК-42",
    to_system: str = "ГСК-2011"
):
    # Проверка формата файла
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="Требуется файл Excel (.xlsx или .xls)")
    
    try:
        contents = await file.read()
        df = pd.read_excel(io.BytesIO(contents))
        
        # Проверка наличия нужных колонок
        required_columns = ['X', 'Y', 'Z']
        if not all(col in df.columns for col in required_columns):
            raise HTTPException(status_code=400, 
                              detail=f"Файл должен содержать колонки: {required_columns}")
        
        converted = []
        
        for _, row in df.iterrows():
            X, Y, Z = row['X'], row['Y'], row['Z']
            
            if to_system == "ГСК-2011":
                p = parameters[from_system]
                res = convert_coordinates(X, Y, Z,
                                          p["dX"], p["dY"], p["dZ"],
                                          np.radians(p["wx"] / 3600),
                                          np.radians(p["wy"] / 3600),
                                          np.radians(p["wz"] / 3600),
                                          p["m"],
                                          to_gsk=True)
            elif from_system == "ГСК-2011":
                p = parameters[to_system]
                res = convert_coordinates(X, Y, Z,
                                          p["dX"], p["dY"], p["dZ"],
                                          np.radians(p["wx"] / 3600),
                                          np.radians(p["wy"] / 3600),
                                          np.radians(p["wz"] / 3600),
                                          p["m"],
                                          to_gsk=False)
            else:
                # Переход через ГСК-2011
                p_from = parameters[from_system]
                X1, Y1, Z1 = convert_coordinates(X, Y, Z,
                                                p_from["dX"], p_from["dY"], p_from["dZ"],
                                                np.radians(p_from["wx"] / 3600),
                                                np.radians(p_from["wy"] / 3600),
                                                np.radians(p_from["wz"] / 3600),
                                                p_from["m"],
                                                to_gsk=True)
                
                p_to = parameters[to_system]
                res = convert_coordinates(X1, Y1, Z1,
                                          p_to["dX"], p_to["dY"], p_to["dZ"],
                                          np.radians(p_to["wx"] / 3600),
                                          np.radians(p_to["wy"] / 3600),
                                          np.radians(p_to["wz"] / 3600),
                                          p_to["m"],
                                          to_gsk=False)
            
            converted.append(res)
        
        result_df = pd.DataFrame(converted, columns=["X", "Y", "Z"])
        
        # Создаем отчет в формате Markdown
        report = create_markdown_report(from_system, to_system, df, result_df, parameters)
        
        # Сохраняем в буфер
        buffer = io.BytesIO(report.encode())
        
        return JSONResponse(content={
            "markdown": report,
            "filename": f"report_{from_system}_to_{to_system}.md"
        })
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

def convert_coordinates(X, Y, Z, dX, dY, dZ, wx, wy, wz, m, to_gsk):
    """Преобразует координаты между системами"""
    if not to_gsk:
        m = -m
        wx, wy, wz = -wx, -wy, -wz
        dX, dY, dZ = -dX, -dY, -dZ
    
    R = np.array([
        [1, wz, -wy],
        [-wz, 1, wx],
        [wy, -wx, 1]
    ])
    
    input_coords = np.array([X, Y, Z])
    transformed = (1 + m) * R @ input_coords + np.array([dX, dY, dZ])
    return transformed[0], transformed[1], transformed[2]