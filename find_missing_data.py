import pandas as pd
import re

# Читаем products.xlsx
products_df = pd.read_excel('products.xlsx', header=None)

print("=== АНАЛИЗ ПРОПУЩЕННЫХ ДАННЫХ ===")

# Проверяем все товары на наличие размеров и весов
missing_data_count = 0

for idx in range(3, len(products_df)):
    product_row = products_df.iloc[idx]

    if pd.isna(product_row[2]) or str(product_row[2]).strip() == '-':
        continue

    name = str(product_row[2])
    desc = str(product_row[4]) if pd.notna(product_row[4]) else ""

    # Ищем различные типы размеров в описании
    found_data = []

    # Высота смесителя
    height_matches = re.findall(r'Высота смесителя:\s*(\d+)\s*мм', desc)
    if height_matches:
        found_data.append(f"Высота смесителя: {height_matches[0]}мм")

    # Высота излива
    height_spout_matches = re.findall(r'Высота излива:\s*(\d+)\s*мм', desc)
    if height_spout_matches:
        found_data.append(f"Высота излива: {height_spout_matches[0]}мм")

    # Длина излива
    length_matches = re.findall(r'Длина излива:\s*(\d+)\s*мм', desc)
    if length_matches:
        found_data.append(f"Длина излива: {length_matches[0]}мм")

    # Ширина лейки
    width_matches = re.findall(r'Ширина лейки:\s*(\d+)\s*мм', desc)
    if width_matches:
        found_data.append(f"Ширина лейки: {width_matches[0]}мм")

    # Высота держателя
    holder_height_matches = re.findall(r'Высота держателя:\s*(\d+)\s*мм', desc)
    if holder_height_matches:
        found_data.append(f"Высота держателя: {holder_height_matches[0]}мм")

    # Размеры упаковки
    package_matches = re.findall(r'ДхШхВ упаковки\s*(\d+)х(\d+)х(\d+)\s*мм', desc)
    if package_matches:
        found_data.append(f"Упаковка: {package_matches[0][0]}x{package_matches[0][1]}x{package_matches[0][2]}мм")

    # Размеры упаковки (альтернативный формат)
    package_matches2 = re.findall(r'Размер упаковки[:\s]*(\d+)[х×](\d+)[х×](\d+)', desc)
    if package_matches2:
        found_data.append(f"Размер упаковки: {package_matches2[0][0]}x{package_matches2[0][1]}x{package_matches2[0][2]}мм")

    # Габариты упаковки (еще один формат)
    package_matches3 = re.findall(r'Габариты упаковки[:\s]*(\d+)[х×](\d+)[х×](\d+)', desc)
    if package_matches3:
        found_data.append(f"Габариты упаковки: {package_matches3[0][0]}x{package_matches3[0][1]}x{package_matches3[0][2]}мм")

    # Простые габариты
    simple_dimensions = re.findall(r'(\d+)[х×](\d+)[х×](\d+)\s*мм', desc)
    if simple_dimensions:
        for dim in simple_dimensions:
            found_data.append(f"Размеры: {dim[0]}x{dim[1]}x{dim[2]}мм")

    # Вес брутто
    weight_matches = re.findall(r'Вес брутто:\s*(\d+)\s*г', desc)
    if weight_matches:
        found_data.append(f"Вес брутто: {weight_matches[0]}г")

    # Вес товара
    weight_matches2 = re.findall(r'Вес товара:\s*(\d+)\s*г', desc)
    if weight_matches2:
        found_data.append(f"Вес товара: {weight_matches2[0]}г")

    # Просто вес
    weight_matches3 = re.findall(r'Вес:\s*(\d+)\s*г', desc)
    if weight_matches3:
        found_data.append(f"Вес: {weight_matches3[0]}г")

    if found_data:
        print(f"\nТовар {idx-2}: {name[:50]}")
        for data in found_data:
            print(f"  {data}")
        missing_data_count += 1
    else:
        print(f"\nТовар {idx-2}: {name[:50]} - НЕТ РАЗМЕРОВ!")

print(f"\nНайдено товаров с размерными данными: {missing_data_count}")

# Теперь покажем примеры описаний с размерами
print("\n=== ПРИМЕРЫ ОПИСАНИЙ С РАЗМЕРАМИ ===")

sample_descriptions = [
    products_df.iloc[3][4],  # Первый товар
    products_df.iloc[7][4],  # Товар с размерами упаковки
    products_df.iloc[10][4]  # Еще один пример
]

for i, desc in enumerate(sample_descriptions):
    if pd.notna(desc):
        print(f"\nПример {i+1}:")
        print(desc)
        print("-" * 80)