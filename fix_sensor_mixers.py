import pandas as pd
import re

# Читаем products.xlsx
products_df = pd.read_excel('products.xlsx', header=None)

print("=== ПОИСК И АНАЛИЗ СЕНСОРНЫХ СМЕСИТЕЛЕЙ ===")

sensor_mixers = []

# Ищем все сенсорные смесители
for idx in range(3, len(products_df)):
    product_row = products_df.iloc[idx]

    if pd.notna(product_row[2]):
        name = str(product_row[2]).lower()
        if 'сенсорный' in name or 'sensor' in name:
            article = str(product_row[5]) if pd.notna(product_row[5]) else "НЕТ АРТИКУЛА"
            description = str(product_row[4]) if pd.notna(product_row[4]) else "НЕТ ОПИСАНИЯ"

            sensor_mixers.append({
                'row': idx,
                'article': article,
                'name': str(product_row[2]),
                'model': str(product_row[3]) if pd.notna(product_row[3]) else "",
                'description': description
            })

print(f"Найдено сенсорных смесителей: {len(sensor_mixers)}")

# Анализируем каждый сенсорный смеситель
for i, mixer in enumerate(sensor_mixers, 1):
    print(f"\n{'='*80}")
    print(f"СЕНСОРНЫЙ СМЕСИТЕЛЬ {i}: {mixer['article']}")
    print(f"{'='*80}")
    print(f"Название: {mixer['name']}")
    print(f"Модель: {mixer['model']}")
    print(f"\nПОЛНОЕ ОПИСАНИЕ:")
    print("-" * 40)
    print(mixer['description'])
    print("-" * 40)

    desc = mixer['description']

    print(f"\nИЗВЛЕЧЕННЫЕ ПАРАМЕТРЫ:")

    # Давление
    pressure_matches = re.findall(r'(\d+[.,]\d+)\s*-\s*(\d+[.,]\d+)\s*[Мм][Пп][Аа]', desc)
    if pressure_matches:
        for p in pressure_matches:
            print(f"  Рабочее давление: {p[0]}-{p[1]} МПа")

    # Температура
    temp_matches = re.findall(r'([+-]?\d+)\s*[°C°С]\s*до\s*([+-]?\d+)\s*[°C°С]', desc)
    if temp_matches:
        for t in temp_matches:
            print(f"  Температура: от {t[0]}°C до {t[1]}°C")

    # Питание
    voltage_matches = re.findall(r'(\d+)\s*[Вв]ольт', desc)
    if voltage_matches:
        print(f"  Питание: {', '.join(voltage_matches)} Вольт")

    # Зона срабатывания
    zone_matches = re.findall(r'(\d+)-(\d+)\s*см', desc)
    if zone_matches:
        for z in zone_matches:
            print(f"  Зона срабатывания: {z[0]}-{z[1]} см")

    # Задержка срабатывания
    delay_matches = re.findall(r'(\d+[.,]?\d*)\s*сек', desc)
    if delay_matches:
        print(f"  Задержка срабатывания: {', '.join(delay_matches)} сек")

    # Размеры излива - ОСНОВНОЕ!
    length_matches = re.findall(r'Длина излива[,\s]*мм:\s*(\d+)\s*\((\d+)\)', desc)
    if length_matches:
        for l in length_matches:
            print(f"  Длина излива: {l[0]} мм (габ. {l[1]} мм)")

    # Высота излива - ОСНОВНОЕ!
    height_matches = re.findall(r'Высота излива[,\s]*мм:\s*(\d+)\s*\((\d+)\)', desc)
    if height_matches:
        for h in height_matches:
            print(f"  Высота излива: {h[0]} мм (габ. {h[1]} мм)")

    # Высота смесителя
    mixer_height = re.findall(r'Высота смесителя:\s*(\d+)\s*мм', desc)
    if mixer_height:
        print(f"  Высота смесителя: {mixer_height[0]} мм")

    # Размеры упаковки
    package_matches = re.findall(r'ДхШхВ упаковки:\s*(\d+)[х×](\d+)[х×](\d+)\s*мм', desc)
    if package_matches:
        for p in package_matches:
            print(f"  Размер упаковки: {p[0]}×{p[1]}×{p[2]} мм")

    # Вес
    weight_matches = re.findall(r'Вес брутто:\s*(\d+)\s*г', desc)
    if weight_matches:
        print(f"  Вес брутто: {weight_matches[0]} г")

    # Керамический картридж
    cartridge_matches = re.findall(r'керамический\s*(\d+)\s*мм', desc)
    if cartridge_matches:
        print(f"  Керамический картридж: {cartridge_matches[0]} мм")

    # Гибкая подводка
    hose_matches = re.findall(r'гибкая подводка\s*(\d+)\s*см', desc)
    if hose_matches:
        print(f"  Гибкая подводка: {hose_matches[0]} см")

    # Максимальное давление
    max_pressure = re.findall(r'Максимальное давление:\s*(\d+)\s*бар', desc)
    if max_pressure:
        print(f"  Максимальное давление: {max_pressure[0]} бар")

def parse_sensor_mixer(description):
    """Специальный парсер для сенсорных смесителей"""
    params = {
        'product_length': None,
        'product_width': None,
        'product_height': None,
        'product_weight': None,
        'package_length': None,
        'package_width': None,
        'package_height': None,
        'package_weight': None,
        'pressure_min': None,
        'pressure_max': None,
        'temperature_min': None,
        'temperature_max': None,
        'voltage_main': None,
        'voltage_backup': None,
        'zone_min': None,
        'zone_max': None,
        'delay': None,
        'max_pressure_bar': None,
        'cartridge_size': None,
        'hose_length': None
    }

    if pd.isna(description):
        return params

    desc = str(description)

    # Длина излива (берем значение без скобок - рабочее)
    length_match = re.search(r'Длина излива[,\s]*мм:\s*(\d+)', desc)
    if length_match:
        params['product_length'] = int(length_match.group(1))

    # Высота излива (берем значение без скобок - рабочее)
    height_match = re.search(r'Высота излива[,\s]*мм:\s*(\d+)', desc)
    if height_match:
        params['product_height'] = int(height_match.group(1))

    # Высота смесителя (если есть)
    mixer_height_match = re.search(r'Высота смесителя:\s*(\d+)\s*мм', desc)
    if mixer_height_match and not params['product_height']:
        params['product_height'] = int(mixer_height_match.group(1))

    # Рабочее давление
    pressure_match = re.search(r'(\d+[.,]\d+)\s*-\s*(\d+[.,]\d+)\s*[Мм][Пп][Аа]', desc)
    if pressure_match:
        params['pressure_min'] = float(pressure_match.group(1).replace(',', '.'))
        params['pressure_max'] = float(pressure_match.group(2).replace(',', '.'))

    # Максимальное давление в барах
    max_pressure_match = re.search(r'Максимальное давление:\s*(\d+)\s*бар', desc)
    if max_pressure_match:
        params['max_pressure_bar'] = int(max_pressure_match.group(1))

    # Температура
    temp_match = re.search(r'([+-]?\d+)\s*[°C°С]\s*до\s*([+-]?\d+)\s*[°C°С]', desc)
    if temp_match:
        params['temperature_min'] = int(temp_match.group(1))
        params['temperature_max'] = int(temp_match.group(2))

    # Питание
    voltage_main = re.search(r'(\d+)\s*[Вв]ольт', desc)
    if voltage_main:
        params['voltage_main'] = int(voltage_main.group(1))

    voltage_backup = re.search(r'резервное питание\s*(\d+)\s*[Вв]ольт', desc)
    if voltage_backup:
        params['voltage_backup'] = int(voltage_backup.group(1))

    # Зона срабатывания
    zone_match = re.search(r'(\d+)-(\d+)\s*см', desc)
    if zone_match:
        params['zone_min'] = int(zone_match.group(1))
        params['zone_max'] = int(zone_match.group(2))

    # Задержка срабатывания
    delay_match = re.search(r'(\d+[.,]?\d*)\s*сек', desc)
    if delay_match:
        params['delay'] = float(delay_match.group(1).replace(',', '.'))

    # Размеры упаковки
    package_match = re.search(r'ДхШхВ упаковки:\s*(\d+)[х×](\d+)[х×](\d+)\s*мм', desc)
    if package_match:
        params['package_length'] = int(package_match.group(1))
        params['package_width'] = int(package_match.group(2))
        params['package_height'] = int(package_match.group(3))

    # Вес брутто
    weight_match = re.search(r'Вес брутто:\s*(\d+)\s*г', desc)
    if weight_match:
        weight_g = int(weight_match.group(1))
        params['package_weight'] = round(weight_g / 1000, 3)
        params['product_weight'] = round(weight_g / 1000, 3)

    # Керамический картридж
    cartridge_match = re.search(r'керамический\s*(\d+)\s*мм', desc)
    if cartridge_match:
        params['cartridge_size'] = int(cartridge_match.group(1))

    # Гибкая подводка
    hose_match = re.search(r'гибкая подводка\s*(\d+)\s*см', desc)
    if hose_match:
        params['hose_length'] = int(hose_match.group(1))

    return params

print(f"\n{'='*80}")
print("ТЕСТИРОВАНИЕ ПАРСЕРА НА СЕНСОРНЫХ СМЕСИТЕЛЯХ:")
print(f"{'='*80}")

for mixer in sensor_mixers:
    print(f"\nТест парсера для {mixer['article']}:")
    params = parse_sensor_mixer(mixer['description'])

    for key, value in params.items():
        if value is not None:
            print(f"  {key}: {value}")

print("\nАНАЛИЗ ЗАВЕРШЕН!")