import os
import re
import sys
import subprocess
import warnings
warnings.filterwarnings('ignore')

# --- БЛОК УСТАНОВКИ ЗАВИСИМОСТЕЙ ---

def install_dependencies():
    """
    Проверяет и устанавливает необходимые библиотеки.
    """
    required_packages = ['pandas', 'openpyxl']
    print("Проверка необходимых библиотек...")
    for package in required_packages:
        try:
            __import__(package)
            print(f"✓ '{package}' уже установлен.")
        except ImportError:
            print(f"✗ '{package}' не найден. Попытка установки...")
            try:
                subprocess.check_call([sys.executable, "-m", "pip", "install", package])
                print(f"✓ '{package}' успешно установлен.")
            except Exception as e:
                print(f"ОШИБКА: Не удалось установить '{package}'.")
                print(f"Пожалуйста, установите его вручную: pip install {package}")
                print(f"Ошибка: {e}")
                sys.exit(1)

# --- НАСТРОЙКИ ---

# Папки для исходных и обработанных файлов
INPUT_DIR = 'xslx'
OUTPUT_DIR = 'processed'

# --- СЛОВАРИ ДЛЯ ПРЕОБРАЗОВАНИЙ ---

SPEC_MAP = {
    "Системы и модели морских подвижных объектов": "КСУ",
    "Управление судовыми электроэнергетическими системами и автоматика судов": "САУ",
    "Системы и технические средства автоматизации и управления": "САУ",
    "Корабельные системы управления": "КСУ",
    "Возобновляемая энергетика": "РАПС",
    "Электрооборудование и автоматика судов": "САУ",
    "Автоматизированные электротехнологические установки и системы": "ЭТПТ",
    "Электропривод и автоматика": "РАПС",
    "Электротехнические системы и технологии": "ЭТПТ",
}

CODE_MAP = {
    "27.03.04": "УТС",
    "27.04.04": "УТС",  # Добавил возможный вариант магистратуры
}

DEPT_MAP = {
    "Робототехники и автоматизации": "РАПС",
    "Систем автоматического управления": "САУ",
    "Корабельных систем управления": "КСУ",
    "Электротехнологической и": "ЭТПТ",
}

# --- ОСНОВНЫЕ ФУНКЦИИ ---

def find_header_info(filepath):
    """
    Находит информацию в 'шапке' Excel файла для нового имени.
    """
    import openpyxl
    
    print(f"  Анализ заголовка файла...")
    
    try:
        workbook = openpyxl.load_workbook(filepath, data_only=True)
        sheet = workbook.active
    except Exception as e:
        raise ValueError(f"Не удалось открыть файл: {e}")
    
    header_info = {
        "plx_string": None,
        "specialization_name": None
    }
    
    # Ищем в первых 30 строках
    for row in sheet.iter_rows(min_row=1, max_row=30):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                # Ищем строку с .plx
                if '.plx' in cell.value.lower():
                    header_info["plx_string"] = cell.value
                    print(f"  ✓ Найден PLX: {cell.value}")
                
                # Ищем название специальности
                for spec_name in SPEC_MAP.keys():
                    if spec_name.lower() in cell.value.lower():
                        header_info["specialization_name"] = spec_name
                        print(f"  ✓ Найдена специальность: {spec_name}")
    
    workbook.close()
    
    if not header_info["plx_string"]:
        raise ValueError(f"Не найден шифр .plx в файле")
    if not header_info["specialization_name"]:
        print(f"  ⚠ Не найдена специальность, будет использовано UNKNOWN")
        
    return header_info

def generate_new_filename(header_info):
    """
    Генерирует новое имя файла на основе найденной информации.
    """
    # Пробуем несколько паттернов для извлечения информации из plx строки
    patterns = [
        r"(\d{2}\.\d{2}\.\d{2})_(\d+)_(\d+)\.plx",  # 27.03.04_23_391.plx
        r"(\d{2}\.\d{2}\.\d{2})_(\d+)_(\d+)",        # 27.03.04_23_391
        r"(\d{2}\.\d{2}\.\d{2}).*?(\d{2}).*?(\d{3})" # более гибкий паттерн
    ]
    
    match = None
    for pattern in patterns:
        match = re.search(pattern, header_info["plx_string"])
        if match:
            break
    
    if not match:
        print(f"  ⚠ Не удалось разобрать PLX строку: {header_info['plx_string']}")
        # Возвращаем имя по умолчанию
        return f"UNKNOWN_{header_info.get('specialization_name', 'UNKNOWN')}_000_00.xlsx"
    
    code, part1, part2 = match.groups()
    
    # Получаем коды из словарей
    uts_code = CODE_MAP.get(code, f"CODE_{code.replace('.', '_')}")
    spec_code = SPEC_MAP.get(header_info.get("specialization_name"), "UNKNOWN_SPEC")
    
    new_name = f"{uts_code}_{spec_code}_{part2}_{part1}.xlsx"
    print(f"  → Новое имя: {new_name}")
    return new_name

def find_table_start(filepath):
    """
    Находит начало таблицы и анализирует структуру файла.
    """
    import pandas as pd
    
    # Читаем первые 50 строк для анализа
    df_preview = pd.read_excel(filepath, header=None, nrows=50)
    
    table_info = {
        'header_row': 0,
        'data_start_row': 0,
        'has_multiheader': False
    }
    
    # Ищем строку с "Наименование"
    for idx, row in df_preview.iterrows():
        row_str = row.astype(str).str.lower()
        if row_str.str.contains('наименование').any():
            table_info['header_row'] = idx
            # Проверяем, есть ли многострочный заголовок
            next_row = df_preview.iloc[idx + 1] if idx + 1 < len(df_preview) else None
            if next_row is not None and next_row.astype(str).str.contains('1|2|3').any():
                table_info['has_multiheader'] = True
                table_info['data_start_row'] = idx + 2
            else:
                table_info['data_start_row'] = idx + 1
            break
    
    return table_info

def process_excel_file(filepath, new_filename):
    """
    Основная функция обработки Excel файла.
    """
    import pandas as pd
    
    try:
        print(f"  Обработка данных...")
        
        # Находим структуру таблицы
        table_info = find_table_start(filepath)
        
        # Читаем файл целиком для анализа структуры
        df_full = pd.read_excel(filepath, header=None)
        
        # Если нашли заголовок, используем его
        if table_info['header_row'] > 0:
            # Читаем с правильной строки заголовка
            if table_info['has_multiheader']:
                df = pd.read_excel(filepath, header=[table_info['header_row'], table_info['header_row'] + 1])
                # Объединяем многоуровневые заголовки
                df.columns = ['_'.join(col).strip() if col[1] != '' else col[0] 
                             for col in df.columns.values]
            else:
                df = pd.read_excel(filepath, header=table_info['header_row'])
        else:
            df = pd.read_excel(filepath)
        
        print(f"  Загружено строк: {len(df)}")
        print(f"  Столбцов: {len(df.columns)}")
        
        # --- Очистка данных ---
        
        # Убираем полностью пустые строки
        df = df.dropna(how='all')
        
        # Пытаемся найти ключевые столбцы
        columns_mapping = {}
        for col in df.columns:
            col_str = str(col).lower()
            if 'наименование' in col_str:
                columns_mapping['Наименование'] = col
            elif 'индекс' in col_str or col == 'Unnamed: 1':
                columns_mapping['Индекс'] = col
            elif 'семестр' in col_str and 'Семестр' not in columns_mapping:
                columns_mapping['Семестр'] = col
            elif 'кафедра' in col_str:
                columns_mapping['Кафедра'] = col
            elif 'зач' in col_str and 'оц' in col_str:
                columns_mapping['Зачет с оц.'] = col
        
        # Переименовываем найденные столбцы
        df = df.rename(columns=columns_mapping)
        
        # Проверяем наличие ключевого столбца
        if 'Наименование' not in df.columns:
            print(f"  ⚠ Столбец 'Наименование' не найден, пытаемся использовать второй столбец")
            if len(df.columns) > 2:
                df.rename(columns={df.columns[2]: 'Наименование'}, inplace=True)
        
        # Убираем строки без наименования
        if 'Наименование' in df.columns:
            df = df[df['Наименование'].notna()]
            df = df[df['Наименование'].astype(str).str.strip() != '']
        
        # --- Фильтрация по диапазону ---
        try:
            start_idx = df[df['Наименование'].astype(str).str.contains(
                "Часть, формируемая|формируемая участниками", 
                case=False, na=False
            )].index
            
            end_idx = df[df['Наименование'].astype(str).str.contains(
                "Блок 2|Практика|Государственная итоговая", 
                case=False, na=False
            )].index
            
            if len(start_idx) > 0 and len(end_idx) > 0:
                start = start_idx[0]
                end = end_idx[0]
                df = df.loc[start:end-1]
                print(f"  Отфильтровано строк: {len(df)}")
        except Exception as e:
            print(f"  ⚠ Не удалось отфильтровать по блокам: {e}")
        
        # --- Удаление строк с "-" в индексе ---
        if 'Индекс' in df.columns:
            df = df[~df['Индекс'].astype(str).str.contains("-", na=False)]
        
        # --- Обработка модулей ---
        processed_rows = []
        current_module = None
        
        for _, row in df.iterrows():
            if 'Наименование' not in row:
                continue
                
            name = str(row['Наименование'])
            
            # Проверяем, является ли это модулем
            if 'модуль' in name.lower() and (
                'Индекс' not in row or 
                pd.isna(row.get('Индекс')) or 
                '+' not in str(row.get('Индекс', ''))
            ):
                current_module = name.strip()
                continue
            
            # Если есть активный модуль и это дисциплина в модуле
            if current_module and 'Индекс' in row and '+' in str(row.get('Индекс', '')):
                row = row.copy()
                row['Наименование'] = f"{current_module}. {name}"
                current_module = None
            
            processed_rows.append(row)
        
        if processed_rows:
            df = pd.DataFrame(processed_rows)
        
        # --- Обработка семестра "34" ---
        if 'Семестр' in df.columns:
            df['Семестр'] = df['Семестр'].astype(str)
            mask_34 = df['Семестр'].isin(['34', '34.0'])
            df.loc[mask_34, 'Наименование'] = df.loc[mask_34, 'Наименование'] + ' (3 и 4 семестр)'
            df.loc[mask_34, 'Семестр'] = '3'
        
        # --- Сокращение названий кафедр ---
        if 'Кафедра' in df.columns:
            for full_name, short_name in DEPT_MAP.items():
                df.loc[df['Кафедра'].astype(str).str.contains(full_name, case=False, na=False), 'Кафедра'] = short_name
        
        # --- Извлечение данных по семестрам ---
        
        # Определяем какие столбцы есть в файле
        available_cols = df.columns.tolist()
        
        # Базовые столбцы для результата
        result_columns = ['Наименование']
        if 'Зачет с оц.' in df.columns:
            result_columns.append('Зачет с оц.')
        if 'Кафедра' in df.columns:
            result_columns.append('Кафедра')
        
        # Ищем столбцы с данными о часах
        semester_cols = ['КП', 'КР', 'з.е.', 'Лек', 'Лаб', 'Пр', 'ИКР', 'СР']
        
        # Находим столбцы семестров в данных
        found_sem_cols = []
        for col in available_cols:
            for sem_col in semester_cols:
                if sem_col in str(col):
                    found_sem_cols.append(col)
                    break
        
        # Если нашли столбцы семестров, добавляем их
        if found_sem_cols:
            result_columns.extend(found_sem_cols[:8])  # Берем первые 8 столбцов
        
        # Создаем итоговый датафрейм
        final_df = df[result_columns].copy() if all(col in df.columns for col in result_columns) else df
        
        # Сохраняем результат
        output_path = os.path.join(OUTPUT_DIR, new_filename)
        final_df.to_excel(output_path, index=False)
        print(f"  ✓ Сохранено строк: {len(final_df)}")
        print(f"  ✓ Файл сохранен: {new_filename}")
        
        return True

    except Exception as e:
        print(f"  ✗ ОШИБКА при обработке: {e}")
        import traceback
        traceback.print_exc()
        return False

# --- ГЛАВНЫЙ БЛОК ---

if __name__ == "__main__":
    print("=" * 60)
    print("ОБРАБОТКА EXCEL ФАЙЛОВ")
    print("=" * 60)
    
    # Установка зависимостей
    install_dependencies()
    
    # Проверка и создание папок
    if not os.path.exists(INPUT_DIR):
        os.makedirs(INPUT_DIR)
        print(f"\n✓ Создана папка '{INPUT_DIR}'")
        print(f"  Поместите в нее ваши .xlsx файлы")
        sys.exit(0)
    
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        print(f"✓ Создана папка '{OUTPUT_DIR}'")

    # Получаем список файлов
    files = [f for f in os.listdir(INPUT_DIR) if f.endswith('.xlsx')]
    
    if not files:
        print(f"\n✗ В папке '{INPUT_DIR}' не найдено .xlsx файлов")
        sys.exit(0)
    
    print(f"\n✓ Найдено файлов: {len(files)}")
    print("-" * 60)
    
    # Счетчики для статистики
    successful = 0
    failed = 0
    
    # Обработка файлов
    for i, filename in enumerate(files, 1):
        filepath = os.path.join(INPUT_DIR, filename)
        print(f"\n[{i}/{len(files)}] Файл: {filename}")
        
        try:
            # Получаем информацию из заголовка
            header_info = find_header_info(filepath)
            
            # Генерируем новое имя
            new_filename = generate_new_filename(header_info)
            
            # Обрабатываем файл
            if process_excel_file(filepath, new_filename):
                successful += 1
            else:
                failed += 1
                
        except Exception as e:
            print(f"  ✗ Критическая ошибка: {e}")
            failed += 1
    
    # Итоговая статистика
    print("\n" + "=" * 60)
    print("РЕЗУЛЬТАТЫ ОБРАБОТКИ")
    print("=" * 60)
    print(f"✓ Успешно обработано: {successful}")
    if failed > 0:
        print(f"✗ С ошибками: {failed}")
    print(f"Результаты сохранены в папке: {OUTPUT_DIR}")
    print("=" * 60)
