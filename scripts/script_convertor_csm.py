import pandas as pd
from datetime import datetime
import re
import os

LOGIN_TO_FIO = {
    'dostanko': 'ОСТАНКО Д.А.',
    'dgudilin': 'ГУДИЛИН Д.А.',
    'akalinin': 'КАЛИНИН А.С.',
    'kelovikov': 'ЕЛОВИКОВ К.Н.',
    'dpanasenkov': 'ПАНАСЕНКОВ Д.А.',
    'imalyarov': 'МАЛЯРОВ И.Д.',
}

DEFAULT_RESOURCE = 'ЦСМ'


def convert_date_format(date_str):
    try:
        dt = datetime.strptime(date_str, '%d/%m/%Y %H:%M:%S')
        return dt.strftime('%Y-%m-%d %H:%M:%S')
    except:
        return date_str


def parse_log_line(line):
    parts = line.split('\t')

    # Находим логин
    login = None
    for p in parts[:4]:
        if p in LOGIN_TO_FIO:
            login = p
            break

    if not login:
        return None

    fio = LOGIN_TO_FIO[login]

    # Находим даты
    dates = []
    for p in parts:
        if re.search(r'\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}', p):
            dates.append(p)

    if len(dates) < 2:
        return None

    start_str = convert_date_format(dates[0])
    end_str = convert_date_format(dates[1])

    # Вычисляем разницу времени в часах
    try:
        start_dt = datetime.strptime(dates[0], '%d/%m/%Y %H:%M:%S')
        end_dt = datetime.strptime(dates[1], '%d/%m/%Y %H:%M:%S')
        duration_hours = (end_dt - start_dt).total_seconds() / 3600
    except:
        duration_hours = 0

    # Ищем ядра и время ЦП
    cores = 0  # первое число (реальные ядра)
    cpu_minutes = 0  # второе число (проц.ядер)

    # Ищем числа, исключая ID задачи (7 цифр)
    numbers = []
    for p in parts:
        if p.isdigit() and len(p) != 7 and len(p) != 8 and len(p) != 9:
            numbers.append(int(p))

    # Ищем пару: ядра (<= 100) и время ЦП (> 100)
    for i in range(len(numbers) - 1):
        if numbers[i] <= 100 and numbers[i + 1] > 100:
            cores = numbers[i]  # реальные ядра
            cpu_minutes = numbers[i + 1]  # проц.ядер (второе число)
            break

    # Если не нашли через type-1, пробуем по-другому
    if cores == 0 or cpu_minutes == 0:
        for i, p in enumerate(parts):
            if p == 'type-1':
                if i + 2 < len(parts):
                    if parts[i + 1].isdigit() and int(parts[i + 1]) <= 100:
                        cores = int(parts[i + 1])
                    if parts[i + 2].isdigit():
                        cpu_minutes = int(parts[i + 2])
                break

    # Если все еще не нашли
    if cpu_minutes == 0:
        for p in parts:
            if p.isdigit() and len(p) != 7 and int(p) > 100:
                cpu_minutes = int(p)
                break

    if cores == 0:
        for p in parts:
            if p.isdigit() and len(p) != 7 and int(p) <= 100 and int(p) > 0:
                cores = int(p)
                break

    # ВАЖНО:
    # - "Число проц.ядер" = cpu_minutes (второе число)
    # - "Время ЦП" = cpu_minutes × duration_hours
    cpu_time = cpu_minutes * duration_hours if cpu_minutes > 0 else 0

    # Определяем методику
    metodika = 'LOGOS_SIMPLE'
    for p in parts:
        if 'ЛОГОС' in p:
            metodika = 'LOGOS_AERO-GIDRO'
            break
        elif 'STARCCM' in p:
            metodika = 'STAR-CCM+'
            break
        elif 'FLUENT' in p:
            metodika = 'ANSYS_FLUENT'
            break

    # Статус
    status = parts[-1].strip() if parts else ''
    comment = status if status in ['FAILED', 'CANCELLED', 'TIMEOUT', 'NODE_FAIL', 'COMPLETED', 'RUNNING'] else ''

    result = {
        'Методика': metodika,
        'ФИО': fio,
        'Начало счёта': start_str,
        'Окончание счёта': end_str,
        'Число проц.ядер': cpu_minutes,  
        'Время ЦП, проц.ядер/ч': round(cpu_time, 1), 
        'Коментарий': comment,
        'Ресурс': DEFAULT_RESOURCE
    }

    return result


def parse_txt_to_excel(input_file, output_file=None):
    if output_file is None:
        base_name = os.path.splitext(input_file)[0]
        output_file = base_name + '_statistic_csm.xlsx'

    rows = []
    seen_records = set()

    with open(input_file, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()

    lines = content.split('\n')
    current_record = ''

    for line in lines:
        line = line.strip()
        if not line:
            continue

        if re.match(r'^\d{7}', line):
            if current_record:
                record = parse_log_line(current_record)
                if record:
                    key = (record['Начало счёта'], record['Окончание счёта'], record['ФИО'])
                    if key not in seen_records:
                        seen_records.add(key)
                        rows.append(record)
            current_record = line
        else:
            current_record += '\t' + line

    if current_record:
        record = parse_log_line(current_record)
        if record:
            key = (record['Начало счёта'], record['Окончание счёта'], record['ФИО'])
            if key not in seen_records:
                seen_records.add(key)
                rows.append(record)

    if not rows:
        print(" Не найдено записей для обработки.")
        return

    df = pd.DataFrame(rows)

    df['Начало счёта'] = pd.to_datetime(df['Начало счёта'])
    df = df.sort_values('Начало счёта')
    df['Начало счёта'] = df['Начало счёта'].dt.strftime('%Y-%m-%d %H:%M:%S')

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='ЦСМ статистика сырая', index=False)

    print(f" Файл успешно создан: {output_file}")
    print(f"Всего записей: {len(rows)}")


if __name__ == '__main__':
    file_name = input("Введите имя файла, который надо обработать (без расширения): ").strip()
    input_file = file_name + '.txt'

    if not os.path.exists(input_file):
        print(f" Файл {input_file} не найден!")
        exit(1)

    output_name = input(
        "Введите имя для выходного файла (без расширения) или нажмите Enter для автоматического: ").strip()

    if output_name:
        output_file = output_name + '.xlsx'
    else:
        output_file = file_name + '_statistic_csm.xlsx'

    parse_txt_to_excel(input_file, output_file)