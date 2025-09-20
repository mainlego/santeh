import pandas as pd
import re

# Читаем products.xlsx
products_df = pd.read_excel('products.xlsx', header=None)

print("=== АНАЛИЗ ОПИСАНИЙ ТОВАРОВ ===")

# Смотрим первые 10 товаров и их описания
for idx in range(3, min(13, len(products_df))):
    product_row = products_df.iloc[idx]

    if pd.notna(product_row[2]) and pd.notna(product_row[4]):
        print(f"\n--- ТОВАР {idx-2} ---")
        print(f"Название: {product_row[2]}")
        print(f"Модель: {product_row[3]}")
        print(f"Артикул: {product_row[5]}")
        print(f"ПОЛНОЕ ОПИСАНИЕ:")
        print(product_row[4])
        print("-" * 80)