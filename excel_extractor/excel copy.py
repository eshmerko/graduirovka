import xlrd
from xlrd import open_workbook

def find_tables(sheet):
    """Находит все таблицы на листе по структурным признакам"""
    tables = []
    current_row = 0
    cols_config = [
        {'left': 1, 'right': 5},  # B-C и F-G
        {'left': 1, 'right': 5}   # B-C и F-G для новых таблиц
    ]
    
    while current_row < sheet.nrows:
        # Поиск начала таблицы
        for config in cols_config:
            try:
                # Проверяем левую и правую части одновременно
                if (sheet.cell_type(current_row, config['left']) == xlrd.XL_CELL_NUMBER and
                    sheet.cell_type(current_row, config['left']+1) == xlrd.XL_CELL_NUMBER and
                    sheet.cell_type(current_row, config['right']) == xlrd.XL_CELL_NUMBER and
                    sheet.cell_type(current_row, config['right']+1) == xlrd.XL_CELL_NUMBER):
                    
                    start_row = current_row
                    end_row = start_row
                    
                    # Определяем границы таблицы
                    while end_row < sheet.nrows:
                        valid = True
                        # Проверяем все четыре столбца
                        for col in [config['left'], config['left']+1, config['right'], config['right']+1]:
                            if sheet.cell_type(end_row, col) not in (xlrd.XL_CELL_NUMBER, xlrd.XL_CELL_TEXT):
                                valid = False
                                break
                        if not valid:
                            break
                        end_row += 1
                    
                    if end_row > start_row + 1:  # Минимум 2 строки
                        tables.append({
                            'start_row': start_row,
                            'end_row': end_row,
                            'left_cols': (config['left'], config['left']+1),
                            'right_cols': (config['right'], config['right']+1)
                        })
                        current_row = end_row
                        break
            except IndexError:
                pass
        current_row += 1
    
    return tables

def process_sheet(sheet, all_data):
    """Обрабатывает один лист Excel"""
    print(f"\n● Обработка листа: {sheet.name}")
    tables = find_tables(sheet)
    
    for i, table in enumerate(tables, 1):
        print(f"\nТаблица {i}: строки {table['start_row']}-{table['end_row']-1}")
        print(f"Левые столбцы: {table['left_cols']}, Правые столбцы: {table['right_cols']}")
        
        # Обработка левой части
        for row in range(table['start_row'], table['end_row']):
            try:
                level = int(sheet.cell_value(row, table['left_cols'][0]))
                capacity = sheet.cell_value(row, table['left_cols'][1])
                all_data.append((level, f"{capacity:.15f}".rstrip('0').rstrip('.')))
                print(f"Обработана строка {row}: {level} ~ {capacity}")
            except:
                pass
        
        # Обработка правой части
        for row in range(table['start_row'], table['end_row']):
            try:
                level = int(sheet.cell_value(row, table['right_cols'][0]))
                capacity = sheet.cell_value(row, table['right_cols'][1])
                all_data.append((level, f"{capacity:.15f}".rstrip('0').rstrip('.')))
                print(f"Обработана строка {row}: {level} ~ {capacity}")
            except:
                pass

def export_data(data, filename):
    """Экспортирует данные в файл"""
    unique_data = {k: v for k, v in data}
    sorted_data = sorted(unique_data.items(), key=lambda x: x[0])
    
    with open(filename, 'w', encoding='utf-8') as f:
        for level, capacity in sorted_data:
            f.write(f"{level}~{capacity}\n")
    
    print(f"\n✅ Успешно экспортировано {len(sorted_data)} записей в {filename}")

def main():
    input_file = "49.xls"
    output_file = "result.txt"
    
    try:
        wb = open_workbook(input_file)
        all_data = []
        
        for sheet in wb.sheets():
            process_sheet(sheet, all_data)
        
        export_data(all_data, output_file)
    
    except Exception as e:
        print(f"❌ Ошибка: {str(e)}")

if __name__ == "__main__":
    main()