import pandas as pd
import re

# Читаем products.xlsx
products_df = pd.read_excel('products.xlsx', header=None)

# Артикулы для глубокого анализа
target_articles = [
    'FL-01-800-L',
    'FL-700-SL',
    'FL-701-SL',
    'FL-702-SL',
    'FL-703-SL',
    'FL-704-SL',
    'FL-705-SL'
]

print("=== ГЛУБОКИЙ АНАЛИЗ КОНКРЕТНЫХ АРТИКУЛОВ ===")

found_products = []

# Ищем эти артикулы в файле
for idx in range(3, len(products_df)):
    product_row = products_df.iloc[idx]

    if pd.notna(product_row[5]):  # Колонка с артикулами
        article = str(product_row[5]).strip()

        if article in target_articles:
            name = str(product_row[2]) if pd.notna(product_row[2]) else "НЕТ НАЗВАНИЯ"
            model = str(product_row[3]) if pd.notna(product_row[3]) else "НЕТ МОДЕЛИ"
            description = str(product_row[4]) if pd.notna(product_row[4]) else "НЕТ ОПИСАНИЯ"

            found_products.append({
                'article': article,
                'name': name,
                'model': model,
                'description': description,
                'row_index': idx
            })

print(f"Найдено товаров: {len(found_products)}")

# Анализируем каждый найденный товар
for i, product in enumerate(found_products, 1):
    print(f"\n{'='*80}")
    print(f"ТОВАР {i}: {product['article']}")
    print(f"{'='*80}")
    print(f"Название: {product['name']}")
    print(f"Модель: {product['model']}")
    print(f"\nПОЛНОЕ ОПИСАНИЕ:")
    print("-" * 40)
    print(product['description'])
    print("-" * 40)

    # Глубокий анализ описания
    desc = product['description']

    print(f"\nИЗВЛЕЧЕННЫЕ ПАРАМЕТРЫ:")

    # Ищем все возможные размеры
    all_dimensions = re.findall(r'(\d+)\s*[х×]\s*(\d+)\s*[х×]\s*(\d+)', desc)
    if all_dimensions:
        print(f"  Найденные размеры:")
        for j, dim in enumerate(all_dimensions):
            print(f"    {j+1}. {dim[0]}x{dim[1]}x{dim[2]} мм")

    # Ищем веса
    weights = re.findall(r'(\d+)\s*г', desc)
    if weights:
        print(f"  Найденные веса: {', '.join(weights)} г")

    # Ищем длину
    lengths = re.findall(r'[Дд]лина[:\s]*(\d+)\s*мм', desc)
    if lengths:
        print(f"  Длина: {', '.join(lengths)} мм")

    # Ищем ширину
    widths = re.findall(r'[Шш]ирина[:\s]*(\d+)\s*мм', desc)
    if widths:
        print(f"  Ширина: {', '.join(widths)} мм")

    # Ищем высоту
    heights = re.findall(r'[Вв]ысота[:\s]*(\d+)\s*мм', desc)
    if heights:
        print(f"  Высота: {', '.join(heights)} мм")

    # Ищем материалы
    materials = re.findall(r'[Мм]атериал[:\s]*([^\n]+)', desc)
    if materials:
        print(f"  Материал: {', '.join(materials)}")

    # Ищем покрытие
    coatings = re.findall(r'[Пп]окрытие[:\s]*([^\n]+)', desc)
    if coatings:
        print(f"  Покрытие: {', '.join(coatings)}")

    # Ищем страну
    countries = re.findall(r'[Сс]трана[:\s]*([^\n]+)', desc)
    if countries:
        print(f"  Страна: {', '.join(countries)}")

    # Ищем тип
    types = re.findall(r'[Тт]ип[:\s]*([^\n]+)', desc)
    if types:
        print(f"  Тип: {', '.join(types)}")

    # Ищем подключение
    connections = re.findall(r'[Пп]рисоединительный[:\s]*([^\n]+)', desc)
    if connections:
        print(f"  Подключение: {', '.join(connections)}")

# Теперь ищем эти артикулы среди всех товаров для проверки
print(f"\n{'='*80}")
print("ПОИСК СРЕДИ ВСЕХ ТОВАРОВ:")
print(f"{'='*80}")

for idx in range(3, len(products_df)):
    product_row = products_df.iloc[idx]

    if pd.notna(product_row[5]):
        article = str(product_row[5]).strip()
        name = str(product_row[2]) if pd.notna(product_row[2]) else ""

        # Ищем похожие артикулы
        for target in target_articles:
            if target.lower() in article.lower() or article.lower() in target.lower():
                print(f"\nСтрока {idx}: {article} - {name[:50]}")
                if pd.notna(product_row[4]):
                    desc_preview = str(product_row[4])[:100] + "..."
                    print(f"  Описание: {desc_preview}")

print(f"\nАНАЛИЗ ЗАВЕРШЕН!")