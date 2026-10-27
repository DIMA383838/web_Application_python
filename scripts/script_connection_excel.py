import pandas as pd
import os
from glob import glob

current_dir = os.getcwd()  # Рабочая папка задаётся приложением/текущей директорией
files = glob(os.path.join(current_dir, "*.xlsx"))

print(f"Найдено файлов для соединения: {len(files)}")

all_data = []
all_dates = []  # Список для сбора всех дат из файлов

for file in files:
    df = pd.read_excel(file)
    all_data.append(df)
    print(f"Добавлен: {os.path.basename(file)} (строк: {len(df)})")
    
    # Проверяем, есть ли колонка 'Начало счёта'
    if 'Начало счёта' in df.columns:
        # Преобразуем в datetime (корректно обрабатывает даты с временем)
        dates = pd.to_datetime(df['Начало счёта'], errors='coerce')
        # Убираем пустые значения
        dates = dates.dropna()
        if not dates.empty:
            all_dates.extend(dates)
            print(f"  - Найдено дат: {len(dates)}, диапазон: {dates.min()} до {dates.max()}")

# Объединяем все данные
if all_data:
    res = pd.concat(all_data, ignore_index=True)
    print(f"\nВсего объединено строк: {len(res)}")
else:
    print("Нет данных для объединения!")
    exit()

# Определяем самую раннюю и самую позднюю дату
if all_dates:
    min_date = min(all_dates).strftime('%Y-%m-%d')
    max_date = max(all_dates).strftime('%Y-%m-%d')
    date_range = f"{min_date}-{max_date}"
    print(f"\nДиапазон дат в файлах: {min_date} до {max_date}")
else:
    date_range = "no_dates_found"
    print("\nВНИМАНИЕ: Колонка 'Начало счёта' не найдена или не содержит дат!")

# Генерируем имя файла автоматически
output_filename = f"statistic_{date_range}.xlsx"
output_file = os.path.join(current_dir, output_filename)

# Сохраняем результат
res.to_excel(output_file, index=False)
print(f"\nГотово! Сохранено в: {output_file}")