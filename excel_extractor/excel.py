import xlrd
from xlrd import open_workbook

def is_valid_data(sheet, row, cols):
    """Проверяет валидность данных в указанных столбцах"""
    try:
        # Проверка целого числа в первом столбце
        int(sheet.cell_value(row, cols[0]))
        
        # Проверка числа с плавающей точкой во втором столбце
        float(sheet.cell_value(row, cols[1]))
        
        return True
    except (ValueError, TypeError):
        return False

def process_columns(sheet, all_data, left_cols, right_cols):
    """Обрабатывает все строки в указанных столбцах"""
    print(f"\n● Начало обработки столбцов: {left_cols} и {right_cols}")
    
    for row_idx in range(sheet.nrows):
        # Обработка левых столбцов (B=1, C=2)
        if is_valid_data(sheet, row_idx, left_cols):
            level = int(sheet.cell_value(row_idx, left_cols[0]))
            capacity = float(sheet.cell_value(row_idx, left_cols[1]))
            formatted = f"{capacity:.15f}".rstrip('0').rstrip('.')
            all_data.append((level, formatted))
            print(f"Обработана строка {row_idx+1} (левые столбцы): {level} ~ {formatted}")
        
        # Обработка правых столбцов (F=5, G=6)
        if is_valid_data(sheet, row_idx, right_cols):
            level = int(sheet.cell_value(row_idx, right_cols[0]))
            capacity = float(sheet.cell_value(row_idx, right_cols[1]))
            formatted = f"{capacity:.15f}".rstrip('0').rstrip('.')
            all_data.append((level, formatted))
            print(f"Обработана строка {row_idx+1} (правые столбцы): {level} ~ {formatted}")

def export_data(data, filename):
    """Экспорт данных с удалением дубликатов"""
    unique_data = {}
    for level, cap in data:
        unique_data[level] = cap
    
    sorted_data = sorted(unique_data.items(), key=lambda x: x[0])
    
    with open(filename, 'w', encoding='utf-8') as f:
        for level, cap in sorted_data:
            f.write(f"{level}~{cap}\n")
    
    print(f"\n✅ Успешно обработано записей: {len(sorted_data)}")
    print(f"Результат сохранён в: {filename}")

def main():
    input_file = "49.xls"
    output_file = "final_result.txt"
    
    try:
        wb = open_workbook(input_file)
        all_data = []
        
        # Конфигурация столбцов
        LEFT_COLS = (1, 2)   # B и C
        RIGHT_COLS = (5, 6)  # F и G
        
        for sheet in wb.sheets():
            print(f"\nОбработка листа: {sheet.name}")
            process_columns(sheet, all_data, LEFT_COLS, RIGHT_COLS)
        
        export_data(all_data, output_file)
    
    except Exception as e:
        print(f"❌ Критическая ошибка: {str(e)}")

if __name__ == "__main__":
    main()