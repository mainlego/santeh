import pandas as pd
import json
import re

# Читаем файл products
products_df = pd.read_excel('products.xlsx', header=None)

# Функция для извлечения всех параметров из описания
def extract_parameters(description):
    params = {
        'length': None,
        'width': None,
        'height': None,
        'weight': None,
        'package_length': None,
        'package_width': None,
        'package_height': None,
        'package_weight': None,
        'material': None,
        'color': None,
        'connection_size': None,
        'type': None
    }

    if pd.isna(description):
        return params

    desc = str(description)

    # Длина трубки/лейки
    length_patterns = [
        r'Длина трубки:\s*(\d+)\s*мм',
        r'макс\.\s*(\d+)\s*мм',
        r'(\d+)\s*мм\s*\(макс\.\s*\d+\s*мм\)'
    ]
    for pattern in length_patterns:
        match = re.search(pattern, desc)
        if match:
            params['length'] = int(match.group(1))
            break

    # Ширина лейки
    width_patterns = [
        r'Ширина лейки:\s*(\d+)\s*мм',
        r'Ширина:\s*(\d+)\s*мм'
    ]
    for pattern in width_patterns:
        match = re.search(pattern, desc)
        if match:
            params['width'] = int(match.group(1))
            break

    # Высота держателя
    height_patterns = [
        r'Высота держателя:\s*(\d+)\s*мм',
        r'Высота:\s*(\d+)\s*мм'
    ]
    for pattern in height_patterns:
        match = re.search(pattern, desc)
        if match:
            params['height'] = int(match.group(1))
            break

    # Размеры упаковки
    package_match = re.search(r'Размер упаковки[:\s]*(\d+)[х×](\d+)[х×](\d+)', desc)
    if package_match:
        params['package_length'] = int(package_match.group(1))
        params['package_width'] = int(package_match.group(2))
        params['package_height'] = int(package_match.group(3))

    # Вес товара
    weight_match = re.search(r'Вес товара:\s*(\d+)\s*г', desc)
    if weight_match:
        params['package_weight'] = float(weight_match.group(1)) / 1000  # переводим в кг

    # Материал
    if 'пластик' in desc.lower():
        params['material'] = 'Пластик'
    elif 'металл' in desc.lower():
        params['material'] = 'Металл'
    elif 'латунь' in desc.lower():
        params['material'] = 'Латунь'

    # Цвет
    colors = ['зеленый', 'зелёный', 'синий', 'белый', 'серый', 'оранжевый']
    for color in colors:
        if color in desc.lower():
            params['color'] = color.capitalize()
            break

    # Размер соединения
    connection_match = re.search(r'Присоединительный размер:\s*([^\\n]+)', desc)
    if connection_match:
        params['connection_size'] = connection_match.group(1).strip()

    return params

# Создаем список всех товаров
products_list = []

for idx in range(3, len(products_df)):
    product_row = products_df.iloc[idx]

    # Пропускаем пустые строки
    if pd.isna(product_row[2]):
        continue

    # Извлекаем параметры
    params = extract_parameters(product_row[4])

    # Создаем объект товара
    product = {
        'id': int(product_row[1]) if pd.notna(product_row[1]) and str(product_row[1]).isdigit() else idx-2,
        'name': str(product_row[2]) if pd.notna(product_row[2]) else '',
        'model': str(product_row[3]) if pd.notna(product_row[3]) else '',
        'full_description': str(product_row[4]) if pd.notna(product_row[4]) else '',
        'article': str(product_row[5]) if pd.notna(product_row[5]) else '',
        'stock_quantity': int(float(product_row[6])) if pd.notna(product_row[6]) and isinstance(product_row[6], (int, float)) else 0,
        'price': float(product_row[7]) if pd.notna(product_row[7]) and isinstance(product_row[7], (int, float)) else 0,
        'total_sum': float(product_row[8]) if pd.notna(product_row[8]) and isinstance(product_row[8], (int, float)) else 0,
        'parameters': params
    }

    products_list.append(product)

# Сохраняем в JSON
with open('products_data.json', 'w', encoding='utf-8') as f:
    json.dump(products_list, f, ensure_ascii=False, indent=2)

print(f"Создан JSON файл с {len(products_list)} товарами")
print("Файл сохранен как: products_data.json")

# Выводим первые 3 товара для проверки
print("\nПример данных (первые 3 товара):")
for i, product in enumerate(products_list[:3]):
    print(f"\nТовар {i+1}:")
    print(f"  Название: {product['name']}")
    print(f"  Модель: {product['model']}")
    print(f"  Артикул: {product['article']}")
    print(f"  Параметры: {product['parameters']}")