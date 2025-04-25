import sys
import re
import os
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QFileDialog,
    QLabel,
    QLineEdit,
    QPushButton,
    QTextEdit,
    QStatusBar,
    QMessageBox
)
from PySide6.QtCore import Qt
from docx import Document
from striprtf.striprtf import rtf_to_text


class FileConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Table Data Extractor")
        self.setGeometry(100, 100, 800, 600)
        self.init_ui()
        self.setup_connections()

    def init_ui(self):
        # Central Widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Input File Section
        input_layout = QHBoxLayout()
        self.input_label = QLabel("Input File:")
        self.input_entry = QLineEdit()
        self.input_entry.setMinimumWidth(400)
        self.browse_input_btn = QPushButton("Browse...")
        input_layout.addWidget(self.input_label)
        input_layout.addWidget(self.input_entry)
        input_layout.addWidget(self.browse_input_btn)

        # Output File Section
        output_layout = QHBoxLayout()
        self.output_label = QLabel("Output File:")
        self.output_entry = QLineEdit("output.txt")
        self.output_entry.setMinimumWidth(400)
        self.browse_output_btn = QPushButton("Browse...")
        output_layout.addWidget(self.output_label)
        output_layout.addWidget(self.output_entry)
        output_layout.addWidget(self.browse_output_btn)

        # Log Section
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("""
            font-family: Consolas; 
            font-size: 10pt; 
            background-color: #f0f0f0;
        """)

        # Convert Button
        self.convert_btn = QPushButton("Convert File")
        self.convert_btn.setStyleSheet("""
            QPushButton {
                padding: 8px;
                font-weight: bold;
                background-color: #4CAF50;
                color: white;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)

        # Status Bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)

        # Assemble Layout
        main_layout.addLayout(input_layout)
        main_layout.addLayout(output_layout)
        main_layout.addWidget(self.log_area)
        main_layout.addWidget(self.convert_btn)

    def setup_connections(self):
        self.browse_input_btn.clicked.connect(self.select_input_file)
        self.browse_output_btn.clicked.connect(self.select_output_file)
        self.convert_btn.clicked.connect(self.process_file)

    def select_input_file(self):
        file_filter = "Supported Files (*.docx *.rtf);;All Files (*)"
        filename, _ = QFileDialog.getOpenFileName(
            self, 
            "Select Input File", 
            os.path.expanduser("~"),
            file_filter
        )
        if filename:
            self.input_entry.setText(filename)

    def select_output_file(self):
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Save Output File",
            os.path.expanduser("~"),
            "Text Files (*.txt)"
        )
        if filename:
            self.output_entry.setText(filename)

    def log_message(self, message, status=False):
        self.log_area.append(message)
        if status:
            self.status_bar.showMessage(message)
        QApplication.processEvents()

    def process_cell(self, cell_text):
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

    def process_cell(self, cell_text):
        cell_text = cell_text.strip()
        if not cell_text:
            return None

        # Используем оригинальную логику из консольной версии
        parts = cell_text.split()
        numbers = []
        for part in parts:
            if re.match(r'^-?\d*\.?\d+$|^-?\d+\.?\d*$', part):
                numbers.append(part)
            else:
                return None  # Отбрасываем всю ячейку при нечисловых значениях
        
        # Сохраняем оригинальные условия
        if len(numbers) in (2, 3):
            return f"{numbers[0]}~{numbers[1]}"
        return None

    def process_rtf(self, file_path):
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                rtf_text = f.read()
            plain_text = rtf_to_text(rtf_text)
            data = []
            for line in plain_text.split('\n'):
                if '|' in line:
                    cells = [cell.strip() for cell in line.split('|') if cell.strip()]
                    for cell in cells:
                        result = self.process_cell(cell)
                        if result:
                            data.append(result)
                            self.log_message(f"[RTF] Found: {result}")
            return data
        except Exception as e:
            self.log_message(f"[ERROR] RTF processing: {str(e)}", status=True)
            return None

    def process_file(self):
        input_path = self.input_entry.text().strip()
        output_path = self.output_entry.text().strip()

        if not input_path:
            QMessageBox.critical(self, "Error", "Please select input file!")
            return

        if not output_path:
            output_path = "output.txt"
            self.output_entry.setText(output_path)

        self.log_area.clear()
        self.log_message("=== Starting processing ===", status=True)

        try:
            # Determine file type
            if input_path.lower().endswith('.docx'):
                data = self.process_docx(input_path)
            elif input_path.lower().endswith('.rtf'):
                data = self.process_rtf(input_path)
            else:
                raise ValueError("Unsupported file format")

            if not data:
                self.log_message("No valid data found in the file!", status=True)
                QMessageBox.warning(self, "Warning", "No valid data found in the file!")
                return
            # Добавляем сортировку как в консольной версии
            data.sort(key=lambda x: float(x.split('~')[0]))

            # Save results
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(data))

            success_msg = f"Successfully processed {len(data)} records!\nSaved to: {os.path.abspath(output_path)}"
            self.log_message(success_msg, status=True)
            QMessageBox.information(self, "Success", success_msg)

        except Exception as e:
            error_msg = f"Critical error: {str(e)}"
            self.log_message(error_msg, status=True)
            QMessageBox.critical(self, "Error", error_msg)
            raise


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = FileConverterApp()
    window.show()
    sys.exit(app.exec())