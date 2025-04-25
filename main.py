import re
from docx import Document
from striprtf.striprtf import rtf_to_text
import sys

def process_cell(cell_text):
    cell_text = cell_text.strip()
    if not cell_text:
        return None
    parts = cell_text.split()
    numbers = []
    for part in parts:
        if re.match(r'^-?\d*\.?\d+$|^-?\d+\.?\d*$', part):
            numbers.append(part)
        else:
            return None
    if len(numbers) in (2, 3):
        return f"{numbers[0]}~{numbers[1]}"
    else:
        return None

def process_docx(file_path):
    doc = Document(file_path)
    data = []
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = cell.text.strip()
                result = process_cell(text)
                if result:
                    data.append(result)
    return data

def process_rtf(file_path):
    with open(file_path, 'r') as f:
        rtf_text = f.read()
    plain_text = rtf_to_text(rtf_text)
    data = []
    for line in plain_text.split('\n'):
        if '|' in line:
            cells = line.split('|')
            for cell in cells:
                result = process_cell(cell)
                if result:
                    data.append(result)
    return data

def process_file(file_path):
    if file_path.lower().endswith('.docx'):
        return process_docx(file_path)
    elif file_path.lower().endswith('.rtf'):
        return process_rtf(file_path)
    else:
        raise ValueError("Unsupported file format. Only .docx and .rtf are supported.")

def save_to_txt(data, output_path):
    # Сортируем данные перед записью в файл
    data.sort(key=lambda x: float(x.split('~')[0]))
    with open(output_path, 'w') as f:
        f.write('\n'.join(data))

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python script.py input_file [output_file]")
        sys.exit(1)
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else 'output.txt'

    try:
        data = process_file(input_file)
        save_to_txt(data, output_file)
        print(f"Data successfully saved to {output_file}")
    except Exception as e:
        print(f"An error occurred: {e}")