import pdfplumber
import pandas as pd
import pytesseract  # Для OCR
from PIL import Image  # Для работы с изображениями из PDF
import io  # Для работы с байтами
import re

# Укажите путь к исполняемому файлу Tesseract OCR (если он не в системном PATH)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extract_data_from_scanned_pdf(pdf_path, output_txt_path):
    """
    Извлекает данные из отсканированного PDF, содержащего таблицы со столбцами "Уровень наполнения" и "Вместимость",
    используя OCR для распознавания текста, сохраняет данные в формате "Уровень наполнения~Вместимость" в текстовый файл.

    Args:
        pdf_path (str): Путь к PDF файлу.
        output_txt_path (str): Путь к выходному текстовому файлу.
    """

    all_extracted_data = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages):
            print(f"Обработка страницы {page_num + 1}/{len(pdf.pages)}")

            # 1. Преобразуем страницу в изображение
            image = page.to_image()
            img = image.original  # Получаем PIL Image

            # 2. Применяем OCR для извлечения текста из изображения
            try:
                text = pytesseract.image_to_string(img, lang='rus') # 'rus' - русский язык, может потребоваться другой
            except Exception as e:
                print(f"Ошибка OCR на странице {page_num + 1}: {e}")
                continue  # Переходим к следующей странице в случае ошибки OCR

            # 3. Разбиваем текст на строки и пытаемся найти строки, содержащие данные
            lines = text.splitlines()
            level_col_index = None
            capacity_col_index = None
            header_found = False

            # Ищем заголовки и определяем индексы столбцов
            for i, line in enumerate(lines):
                if re.search(r'(?i)уровень\s+наполнения', line) and re.search(r'(?i)вместимость', line):
                    # Предполагаем, что заголовки разделены пробелами или другими символами
                    header_line = line.lower()
                    level_col_index = header_line.find("уровень наполнения")
                    capacity_col_index = header_line.find("вместимость")
                    header_found = True
                    break  # Заголовки найдены, можно начинать обработку данных

            if not header_found:
                print(f"Не удалось найти заголовки 'Уровень наполнения' и 'Вместимость' на странице {page_num + 1}. Пропускаем страницу.")
                continue

            # Извлекаем данные
            for line in lines[i + 1:]:  # Начинаем с первой строки после заголовков
                line = line.strip()
                if not line:
                    continue  # Пропускаем пустые строки

                # Извлекаем данные, основываясь на позициях заголовков
                try:
                    level = line[level_col_index:capacity_col_index].strip()
                    capacity_str = line[capacity_col_index:].strip()
                    capacity_str = capacity_str.replace(',', '.')  # Заменяем запятую на точку
                    capacity = float(capacity_str)

                    all_extracted_data.append((level, capacity))
                except (ValueError, IndexError) as e:
                    print(f"Ошибка при обработке строки '{line}' на странице {page_num + 1}: {e}")
                    continue


    # 4. Записываем извлеченные данные в текстовый файл
    with open(output_txt_path, 'w') as outfile:
        for level, capacity in all_extracted_data:
            outfile.write(f"{level}~{capacity:.3f}\n")

    print(f"Данные успешно извлечены и сохранены в файл: {output_txt_path}")


# Пример использования
pdf_file = 'input.pdf'  # Замените на путь к вашему PDF файлу
txt_file = 'output.txt'  # Замените на желаемый путь к выходному текстовому файлу

extract_data_from_scanned_pdf(pdf_file, txt_file)