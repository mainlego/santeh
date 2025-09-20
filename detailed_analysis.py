import pandas as pd
import openpyxl
from openpyxl import load_workbook

# Читаем products.xlsx
print("=== ДЕТАЛЬНЫЙ АНАЛИЗ products.xlsx ===")
products_df = pd.read_excel('products.xlsx', header=None)
print(f"Размер: {products_df.shape}")
print("\nПервые 15 строк (все колонки):")
print(products_df.head(15).to_string())

# Читаем template.xlsx - первый лист
print("\n\n=== ДЕТАЛЬНЫЙ АНАЛИЗ template.xlsx (Sheet1) ===")
template_df = pd.read_excel('template.xlsx', sheet_name='Sheet1', header=None)
print(f"Размер: {template_df.shape}")
print("\nПервые 10 строк:")
print(template_df.head(10).to_string())

# Анализируем структуру для маппинга
print("\n\n=== АНАЛИЗ СООТВЕТСТВИЙ ===")
print("\nЗаголовки в products (строка 1):")
for i, val in enumerate(products_df.iloc[0]):
    if pd.notna(val):
        print(f"Колонка {i}: {val}")

print("\nЗаголовки в template (строка 1):")
for i, val in enumerate(template_df.iloc[0]):
    if pd.notna(val):
        print(f"Колонка {i}: {val}")