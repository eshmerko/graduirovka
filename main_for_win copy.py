import sys
import re
import os
import time
import tempfile
import pythoncom
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
from PySide6.QtGui import (
    QFont,
    QPixmap,
    QColor,
    QLinearGradient,
    QBrush,
    QIcon,
    QPainter,
    QPaintEvent
)
from PySide6.QtCore import (
    Qt,
    QPropertyAnimation,
    QEasingCurve,
    QSize
)
from striprtf.striprtf import rtf_to_text
import win32com.client as win32


class DeveloperWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.setup_animations()

    def setup_ui(self):
        self.setObjectName("DeveloperWidget")
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(20, 10, 20, 10)
        main_layout.setSpacing(15)

        # Иконка (замените путь на свой)
        # pixmap = QPixmap("dev.png").scaled(64, 64, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        # self.icon_label = QLabel()
        # self.icon_label.setPixmap(pixmap)
        # self.icon_label.setText("")  # Удалить эмодзи
        
        # Временная иконка с эмодзи
        self.icon_label = QLabel()
        # self.icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        # self.icon_label.setStyleSheet("""
        #     background-color: #4CAF50;
        #     border-radius: 32px;
        #     font-size: 24px;
        #     color: white;
        # """)
        # self.icon_label.setText("")

        # Текстовая информация
        text_layout = QVBoxLayout()
        text_layout.setSpacing(2)
        
        self.razrab_label = QLabel("Разработал:")
        self.dol_label = QLabel("Инженер-технолог")
        self.name_label = QLabel("Шмерко Евгений Леонидович")
        self.org_label = QLabel("ОАО «Пуховичинефтепродукт»")
        self.email_label = QLabel("Email: e.shmerko@beloil.by")
        self.phone_label = QLabel("Тел.: +375 44 7777710")
        self.year_label = QLabel("ver. 0.0.1")

        # Настройка шрифтов
        name_font = QFont("Segoe UI Semibold", 8)
        details_font = QFont("Segoe UI", 6)
        
        self.razrab_label.setFont(name_font)
        self.dol_label.setFont(details_font)
        self.name_label.setFont(name_font)
        self.org_label.setFont(details_font)
        self.email_label.setFont(details_font)
        self.phone_label.setFont(details_font)
        self.year_label.setFont(details_font)

        # Цвета
        primary_color = QColor("#2c3e50")
        self.razrab_label.setStyleSheet(f"color: {primary_color.name()};")
        self.dol_label.setStyleSheet(f"color: {primary_color.name()}; opacity: 0.9;")
        self.name_label.setStyleSheet(f"color: {primary_color.name()};")
        self.org_label.setStyleSheet(f"color: {primary_color.name()}; opacity: 0.9;")
        self.email_label.setStyleSheet(f"color: {primary_color.name()}; opacity: 0.8;")
        self.phone_label.setStyleSheet(f"color: {primary_color.name()}; opacity: 0.8;")
        self.year_label.setStyleSheet(f"color: {primary_color.name()}; opacity: 0.7;")

        # Выравнивание
        for label in [self.razrab_label, self.dol_label, self.name_label, self.org_label, 
                     self.email_label, self.phone_label, self.year_label]:
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        text_layout.addWidget(self.razrab_label)
        text_layout.addWidget(self.dol_label)
        text_layout.addWidget(self.name_label)
        text_layout.addWidget(self.org_label)
        text_layout.addWidget(self.email_label)
        text_layout.addWidget(self.phone_label)
        text_layout.addWidget(self.year_label)

        # Центрируем всю группу
        container = QWidget()
        container_layout = QHBoxLayout(container)
        container_layout.addStretch()
        container_layout.addWidget(self.icon_label)
        container_layout.addSpacing(15)
        container_layout.addLayout(text_layout)
        container_layout.addStretch()
        container_layout.setContentsMargins(0, 0, 0, 0)

        main_layout.addWidget(container)

        # Стилизация
        self.setStyleSheet("""
            QWidget#DeveloperWidget {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #f8f9fa, stop:1 #e9ecef);
                border-top: 1px solid #dee2e6;
                border-radius: 8px;
                margin: 5px;
            }
            QLabel {
                background: transparent;
            }
        """)

    def setup_animations(self):
        self.animation = QPropertyAnimation(self, b"windowOpacity")
        self.animation.setDuration(300)

    def enterEvent(self, event):
        self.animation.stop()
        self.animation.setStartValue(1.0)
        self.animation.setEndValue(0.95)
        self.animation.setEasingCurve(QEasingCurve.OutQuad)
        self.animation.start()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.animation.stop()
        self.animation.setStartValue(0.95)
        self.animation.setEndValue(1.0)
        self.animation.setEasingCurve(QEasingCurve.OutQuad)
        self.animation.start()
        super().leaveEvent(event)

    def setup_animations(self):
        self.animation = QPropertyAnimation(self, b"windowOpacity")
        self.animation.setDuration(300)

    def enterEvent(self, event):
        self.animation.stop()
        self.animation.setStartValue(1.0)
        self.animation.setEndValue(0.95)
        self.animation.setEasingCurve(QEasingCurve.OutQuad)
        self.animation.start()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.animation.stop()
        self.animation.setStartValue(0.95)
        self.animation.setEndValue(1.0)
        self.animation.setEasingCurve(QEasingCurve.OutQuad)
        self.animation.start()
        super().leaveEvent(event)


class FileConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Извлечение данных из таблиц")
        self.setGeometry(100, 100, 800, 600)
        # Инициализируем статусную панель
        self.statusBar().showMessage("Готово")
        self.setup_ui()
        self.setup_connections()

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(15)

        # Секция выбора файлов
        file_selection_layout = QVBoxLayout()
        file_selection_layout.addLayout(self.create_input_layout())
        file_selection_layout.addLayout(self.create_output_layout())

        # Лог-панель
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("""
            QTextEdit {
                font-family: 'Segoe UI';
                font-size: 11pt;
                background-color: #ffffff;
                border: 1px solid #dee2e6;
                border-radius: 4px;
                padding: 8px;
                min-height: 200px;
            }
        """)

        # Кнопка конвертации
        self.convert_btn = QPushButton("Конвертировать файл")
        self.convert_btn.setStyleSheet(self.get_button_style())

        # Блок разработчика
        developer_widget = DeveloperWidget()

        # Компоновка элементов
        main_layout.addLayout(file_selection_layout)
        main_layout.addWidget(self.log_area, 1)
        main_layout.addWidget(self.convert_btn)
        main_layout.addWidget(developer_widget)

    def create_input_layout(self):
        layout = QHBoxLayout()
        self.input_label = QLabel("Исходный файл:")
        self.input_entry = QLineEdit()
        self.input_entry.setPlaceholderText("Выберите файл для обработки...")
        self.input_entry.setMinimumWidth(400)
        self.browse_input_btn = QPushButton("Выбрать...")
        
        layout.addWidget(self.input_label)
        layout.addWidget(self.input_entry, 1)
        layout.addWidget(self.browse_input_btn)
        return layout

    def create_output_layout(self):
        layout = QHBoxLayout()
        self.output_label = QLabel("Результирующий файл:")
        self.output_entry = QLineEdit("результат.txt")
        self.output_entry.setMinimumWidth(400)
        self.browse_output_btn = QPushButton("Выбрать...")
        
        layout.addWidget(self.output_label)
        layout.addWidget(self.output_entry, 1)
        layout.addWidget(self.browse_output_btn)
        return layout

    def get_button_style(self):
        return """
            QPushButton {
                padding: 12px 24px;
                font-weight: 600;
                background-color: #4CAF50;
                color: white;
                border-radius: 6px;
                font-size: 12pt;
                margin: 10px 0;
                border: none;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
        """

    def setup_connections(self):
        self.browse_input_btn.clicked.connect(self.select_input_file)
        self.browse_output_btn.clicked.connect(self.select_output_file)
        self.convert_btn.clicked.connect(self.process_file)

    def select_input_file(self):
        file_filter = "Поддерживаемые файлы (*.docx *.doc *.rtf);;Все файлы (*)"
        filename, _ = QFileDialog.getOpenFileName(
            self, 
            "Выберите исходный файл", 
            os.path.expanduser("~"),
            file_filter
        )
        if filename:
            self.input_entry.setText(filename)

    def select_output_file(self):
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранение результата",
            os.path.expanduser("~"),
            "Текстовые файлы (*.txt)"
        )
        if filename:
            self.output_entry.setText(filename)

    def log_message(self, message, status=False):
        self.log_area.append(message)
        if status:
            # Используем встроенный метод statusBar()
            self.statusBar().showMessage(message, 5000)
        QApplication.processEvents()

    def convert_to_rtf(self, input_path):
        try:
            if not os.path.exists(input_path):
                self.log_message(f"[ОШИБКА] Файл не найден: {input_path}", status=True)
                return None

            pythoncom.CoInitialize()
            word = None
            doc = None
            temp_path = None

            try:
                word = win32.Dispatch("Word.Application")
                word.Visible = False
                word.DisplayAlerts = False

                for attempt in range(3):
                    try:
                        doc = word.Documents.Open(
                            FileName=input_path,
                            ConfirmConversions=False,
                            ReadOnly=True,
                            AddToRecentFiles=False,
                            PasswordDocument=""
                        )
                        break
                    except Exception as e:
                        if attempt == 2:
                            raise RuntimeError(f"Не удалось открыть файл после 3 попыток: {str(e)}")
                        time.sleep(1)

                if doc is None:
                    raise RuntimeError("Не удалось открыть документ в Word")

                temp_dir = tempfile.gettempdir()
                temp_path = os.path.join(temp_dir, f"temp_{os.path.basename(input_path)}.rtf")
                
                doc.SaveAs(temp_path, FileFormat=6)
                self.log_message(f"Успешно конвертировано в: {temp_path}")

                return temp_path

            except Exception as e:
                error_msg = f"[ОШИБКА] Конвертация в RTF: {str(e)}"
                if "The document is locked" in str(e):
                    error_msg += "\nФайл заблокирован для редактирования!"
                elif "password" in str(e).lower():
                    error_msg += "\nФайл защищен паролем!"
                self.log_message(error_msg, status=True)
                return None

            finally:
                try:
                    if doc:
                        doc.Close(SaveChanges=False)
                    if word:
                        word.Quit()
                    pythoncom.CoUninitialize()
                    if temp_path and not os.path.exists(temp_path):
                        os.remove(temp_path)
                except Exception as cleanup_error:
                    self.log_message(f"[ОШИБКА] Очистка ресурсов: {str(cleanup_error)}")

        except Exception as outer_error:
            self.log_message(f"[ОШИБКА] Внешняя ошибка конвертации: {str(outer_error)}", status=True)
            return None

    def process_rtf(self, rtf_path):
        try:
            with open(rtf_path, 'r', encoding='utf-8') as f:
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
                            self.log_message(f"[RTF] Найдено: {result}")
            return data
        except Exception as e:
            self.log_message(f"[ОШИБКА] Обработка RTF: {str(e)}", status=True)
            return None

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
        return None

    def process_file(self):
        input_path = self.input_entry.text().strip()
        output_path = self.output_entry.text().strip()

        if not input_path:
            QMessageBox.critical(self, "Ошибка", "Пожалуйста, выберите исходный файл!")
            return

        if not output_path:
            output_path = "результат.txt"
            self.output_entry.setText(output_path)

        self.log_area.clear()
        self.log_message("=== Начало обработки ===", status=True)

        try:
            if input_path.lower().endswith('.rtf'):
                rtf_path = input_path
            else:
                self.log_message("Конвертация в RTF...", status=True)
                rtf_path = self.convert_to_rtf(input_path)
                if not rtf_path:
                    raise ValueError("Не удалось конвертировать файл в RTF")

            data = self.process_rtf(rtf_path)
            
            if rtf_path != input_path:
                try:
                    os.remove(rtf_path)
                except Exception as e:
                    self.log_message(f"[ВНИМАНИЕ] Не удалось удалить временный файл: {str(e)}")

            if not data:
                self.log_message("Не найдено подходящих данных!", status=True)
                QMessageBox.warning(self, "Предупреждение", "В файле не найдено подходящих данных!")
                return
            
            data.sort(key=lambda x: float(x.split('~')[0]))

            with open(output_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(data))

            success_msg = f"Успешно обработано записей: {len(data)}!\nРезультат сохранен: {os.path.abspath(output_path)}"
            self.log_message(success_msg, status=True)
            QMessageBox.information(self, "Успех", success_msg)

        except Exception as e:
            error_msg = f"Критическая ошибка: {str(e)}"
            self.log_message(error_msg, status=True)
            QMessageBox.critical(self, "Ошибка", error_msg)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setWindowIcon(QIcon(":/icons/app_icon.png"))
    
    window = FileConverterApp()
    window.show()
    sys.exit(app.exec())