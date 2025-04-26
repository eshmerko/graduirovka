# Copyright (c) 2025 Шмерко Евгений Леонидович
# SPDX-License-Identifier: MIT

import sys
import os
import re
import time
import tempfile
import pythoncom
import requests
import json
from PySide6.QtCore import (
    Qt, QPropertyAnimation, QEasingCurve, QThread, Signal, QUrl, QSettings, QTimer
)
from PySide6.QtGui import (
    QFont, QPixmap, QColor, QLinearGradient, QBrush, QIcon, QPainter, QAction, QDesktopServices
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QFileDialog, QLabel, QLineEdit, QPushButton, QTextEdit, QStatusBar,
    QMessageBox, QDialog, QScrollArea, QScrollBar
)
from striprtf.striprtf import rtf_to_text
import win32com.client as win32

from config import AppConfig

class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"О программе (версия {AppConfig.VERSION})")
        self.setup_ui()
        self.setMinimumSize(400, 300)

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        developer_widget = DeveloperWidget()
        layout.addWidget(developer_widget)


# Загрузчик изображений
class ImageLoader(QThread):
    image_loaded = Signal(QPixmap)
    load_failed = Signal()

    def run(self):
        try:
            response = requests.get(
                "https://eshmerko.com/developer_photo.jpg",
                timeout=10,
                headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'application/json',
                'Cache-Control': 'no-cache'},
                verify=False
            )
            if response.status_code == 200:
                pixmap = QPixmap()
                pixmap.loadFromData(response.content)
                self.image_loaded.emit(pixmap)
            else:
                self.load_failed.emit()
        except Exception:
            self.load_failed.emit()

# Виджет разработчика
class DeveloperWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        self.setup_animations()
        self.init_photo_loading()

    def setup_ui(self):
        self.setObjectName("DeveloperWidget")
        self.setMinimumSize(500, 220)
        
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(30, 30, 30, 30)
        main_layout.setSpacing(30)

        # Блок с фотографией
        self.photo_label = QLabel("Фото\nнедоступно")
        self.photo_label.setFixedSize(120, 180)
        self.photo_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.photo_label.setStyleSheet("""
            QLabel {
                background-color: #f8f9fa;
                border: 2px dashed #ced4da;
                border-radius: 12px;
                color: #6c757d;
                font-size: 14px;
                font-weight: 500;
                padding: 10px;
            }
        """)

        # Информационный блок
        info_layout = QVBoxLayout()
        info_layout.setSpacing(10)

        # Имя и должность
        self.name_label = QLabel(AppConfig.DEVELOPER_NAME)
        self.name_label.setFont(QFont("Segoe UI Semibold", 16, QFont.Weight.Bold))
        self.name_label.setStyleSheet("color: #E4E4E4; margin-bottom: 5px;")

        self.position_label = QLabel("Инженер-технолог")
        self.position_label.setFont(QFont("Segoe UI", 14))
        self.position_label.setStyleSheet("color: #E4E4E4; margin-bottom: 8px;")

        # Контакты
        self.company_label = QLabel(AppConfig.COMPANY_NAME)
        self.company_label.setFont(QFont("Segoe UI", 12))
        self.company_label.setStyleSheet("color: #E4E4E4; margin-bottom: 15px;")

        contacts_layout = QVBoxLayout()
        contacts_layout.setSpacing(8)

        # Email
        email_widget = QWidget()
        email_layout = QHBoxLayout(email_widget)
        email_layout.setContentsMargins(0, 0, 0, 0)
        email_icon = QLabel("📧")
        email_icon.setFont(QFont("Segoe UI", 14))
        email_text = QLabel(AppConfig.DEVELOPER_EMAIL)
        email_text.setFont(QFont("Segoe UI", 12))
        email_text.setStyleSheet("color: #E4E4E4;")
        email_layout.addWidget(email_icon)
        email_layout.addWidget(email_text)
        email_layout.addStretch()

        # Телефон
        phone_widget = QWidget()
        phone_layout = QHBoxLayout(phone_widget)
        phone_layout.setContentsMargins(0, 0, 0, 0)
        phone_icon = QLabel("📱")
        phone_icon.setFont(QFont("Segoe UI", 14))
        phone_text = QLabel(AppConfig.DEVELOPER_PHONE)
        phone_text.setFont(QFont("Segoe UI", 12))
        phone_text.setStyleSheet("color: #E4E4E4;")
        phone_layout.addWidget(phone_icon)
        phone_layout.addWidget(phone_text)
        phone_layout.addStretch()

        contacts_layout.addWidget(email_widget)
        contacts_layout.addWidget(phone_widget)

        # Версия и лицензия
        version_license_layout = QVBoxLayout()
        
        self.version_label = QLabel(f"Версия: {AppConfig.VERSION}")
        self.version_label.setFont(QFont("Segoe UI", 10))
        self.version_label.setStyleSheet("color: #6c757d;")
        self.version_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        
        self.license_label = QLabel(AppConfig.license_header())
        self.license_label.setFont(QFont("Segoe UI", 9))
        self.license_label.setStyleSheet("color: #6c757d;")
        self.license_label.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.license_label.setWordWrap(True)
        
        version_license_layout.addWidget(self.version_label)
        version_license_layout.addWidget(self.license_label)

        # Сборка layout
        info_layout.addWidget(self.name_label)
        info_layout.addWidget(self.position_label)
        info_layout.addWidget(self.company_label)
        info_layout.addLayout(contacts_layout)
        info_layout.addStretch()
        info_layout.addLayout(version_license_layout)

        main_layout.addWidget(self.photo_label)
        main_layout.addLayout(info_layout)

        self.setStyleSheet("""
            QWidget#DeveloperWidget {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #ffffff, stop:1 #f8f9fa);
                border-radius: 18px;
                border: 1px solid #dee2e6;
                box-shadow: 0px 2px 8px rgba(0, 0, 0, 0.08);
            }
        """)

    def setup_animations(self):
        self.animation = QPropertyAnimation(self, b"windowOpacity")
        self.animation.setDuration(350)
        self.animation.setEasingCurve(QEasingCurve.OutQuad)

    def init_photo_loading(self):
        self.loader = ImageLoader()
        self.loader.image_loaded.connect(self.handle_image_loaded)
        self.loader.load_failed.connect(self.handle_image_load_failed)
        self.loader.start()

    def handle_image_loaded(self, pixmap):
        scaled_pixmap = pixmap.scaled(
            160, 160,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation
        )
        self.photo_label.setPixmap(scaled_pixmap)
        self.photo_label.setText("")
        self.photo_label.setStyleSheet("""
            QLabel {
                border: 2px solid #e9ecef;
                border-radius: 12px;
                background-color: #ffffff;
            }
        """)

    def handle_image_load_failed(self):
        self.photo_label.setText("Фото\nнедоступно")
        self.photo_label.setStyleSheet("""
            QLabel {
                background-color: #f8f9fa;
                border: 2px dashed #dee2e6;
                border-radius: 12px;
                color: #6c757d;
                font-size: 14px;
                font-weight: 500;
                padding: 10px;
            }
        """)

    def enterEvent(self, event):
        self.animation.stop()
        self.animation.setStartValue(1.0)
        self.animation.setEndValue(0.96)
        self.animation.start()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.animation.stop()
        self.animation.setStartValue(0.96)
        self.animation.setEndValue(1.0)
        self.animation.start()
        super().leaveEvent(event)

# Стартовый экран с инструкцией
class StartupScreen(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        
    def setup_ui(self):
        self.setWindowTitle("Добро пожаловать")
        self.setFixedSize(680, 500)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(30, 20, 30, 20)
        
        # Заголовок
        title = QLabel("Инструкция и условия использования")
        title.setFont(QFont("Segoe UI Semibold", 16))
        title.setStyleSheet("color: #2c3e50; margin-bottom: 15px;")
        
        # Текст с прокруткой
        scroll_area = QScrollArea()
        content = QLabel()
        content.setWordWrap(True)
        content.setTextFormat(Qt.TextFormat.RichText)
        content.setText(self.get_content_text())
        content.setStyleSheet("font-size: 12pt; color: #4a4a4a; padding: 10px;")
        
        scroll_area.setWidget(content)
        scroll_area.setWidgetResizable(True)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOn)
        
        # Кнопка принятия
        accept_btn = QPushButton("Принять и продолжить")
        accept_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 12px 24px;
                border-radius: 6px;
                font-size: 12pt;
                margin-top: 20px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        accept_btn.clicked.connect(self.accept)
        
        layout.addWidget(title)
        layout.addWidget(scroll_area)
        layout.addWidget(accept_btn, 0, Qt.AlignmentFlag.AlignCenter)
        
        self.setStyleSheet("""
            QDialog {
                background: #ffffff;
                border-radius: 12px;
            }
            QScrollArea {
                border: none;
            }
            QScrollBar:vertical {
                width: 12px;
                background: #f0f0f0;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background: #c0c0c0;
                min-height: 30px;
                border-radius: 6px;
            }
        """)
    
    def get_content_text(self):
        return f"""
        <h3>📋 Поддерживаемые форматы файлов:</h3>
        <ul>
            <li>Microsoft Word (.docx)</li>
            <li>Microsoft Word 97-2003 (.doc)</li>
            <li>Rich Text Format (.rtf)</li>
        </ul>
        
        <h3>🛠️ Инструкция по использованию:</h3>
        <ol>
            <li>Нажмите кнопку <b>'Выбрать...'</b> в разделе <i>'Исходный файл'</i></li>
            <li>Выберите документ для обработки</li>
            <li>Укажите имя результирующего файла (по умолчанию: результат.txt)</li>
            <li>Нажмите кнопку <b>'Конвертировать файл'</b></li>
            <li>Ожидайте завершения процесса в лог-панели</li>
        </ol>
        
        <h3>⚠️ Отказ от ответственности:</h3>
        <p>Программа {AppConfig.APP_NAME} версии {AppConfig.VERSION} предоставляется <b>'как есть'</b>, 
        без каких-либо гарантий. Разработчик ({AppConfig.DEVELOPER_NAME}) не несет ответственности за:</p>
        <ul>
            <li>Прямой или косвенный ущерб, вызванный использованием программы</li>
            <li>Потерю данных или их некорректную обработку</li>
            <li>Проблемы совместимости с конкретными версиями ПО</li>
            <li>Последствия использования нелицензионного программного обеспечения</li>
        </ul>
        <p>Используя данное программное обеспечение, вы соглашаетесь с этими условиями.</p>
        """

# Основной класс приложения
class FileConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_version = AppConfig.VERSION
        self.update_check_url = AppConfig.UPDATE_CHECK_URL
        self.base_download_url = AppConfig.BASE_DOWNLOAD_URL
        
        self.setWindowTitle(AppConfig.APP_NAME)
        self.setGeometry(100, 100, 800, 600)
        self.statusBar().showMessage("Готово")
        
        self.setup_ui()
        self.setup_connections()
        self.setup_menu()
        
        # Показ стартового экрана
        if not QSettings().value("agreement_accepted", False):
            self.show_startup_screen()

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(15)

        # Верхняя панель
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 15)
        
        right_container = QHBoxLayout()
        right_container.setSpacing(15)
        
        self.update_label = QLabel()
        self.update_label.setStyleSheet("""
            color: #28a745; 
            font-size: 11pt;
            qproperty-alignment: AlignRight;
        """)
        self.update_label.setOpenExternalLinks(False)
        self.update_label.linkActivated.connect(self.open_update_url)
        
        self.about_btn = QPushButton("Разработчик")
        self.about_btn.setStyleSheet("""
            QPushButton {
                padding: 5px 12px;
                font-weight: 500;
                background-color: #f8f9fa;
                color: #212529;
                border-radius: 4px;
                font-size: 11pt;
                border: 1px solid #dee2e6;
            }
            QPushButton:hover {
                background-color: #e9ecef;
                border-color: #ced4da;
            }
            QPushButton:pressed {
                background-color: #dee2e6;
            }
        """)
        self.about_btn.setFixedSize(120, 30)
        
        right_container.addWidget(self.update_label)
        right_container.addWidget(self.about_btn)
        
        header_layout.addStretch()
        header_layout.addLayout(right_container)

        # Выбор файлов
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
                background-color: #B8B8B8;
                border: 1px solid #dee2e6;
                border-radius: 4px;
                padding: 8px;
                min-height: 200px;
            }
        """)

        # Кнопка конвертации
        self.convert_btn = QPushButton("Конвертировать файл")
        self.convert_btn.setStyleSheet("""
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
        """)

        main_layout.addLayout(header_layout)
        main_layout.addLayout(file_selection_layout)
        main_layout.addWidget(self.log_area, 1)
        main_layout.addWidget(self.convert_btn)

        QTimer.singleShot(1000, self.check_for_updates)

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

    def setup_menu(self):
        menu_bar = self.menuBar()
        help_menu = menu_bar.addMenu("Справка")
        
        show_manual_action = QAction("Показать инструкцию", self)
        show_manual_action.triggered.connect(self.show_startup_screen)
        help_menu.addAction(show_manual_action)
        
        about_action = QAction("О программе", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)

    def setup_connections(self):
        self.browse_input_btn.clicked.connect(self.select_input_file)
        self.browse_output_btn.clicked.connect(self.select_output_file)
        self.convert_btn.clicked.connect(self.process_file)
        self.about_btn.clicked.connect(self.show_about_dialog)

    def show_startup_screen(self):
        startup_dialog = StartupScreen(self)
        if startup_dialog.exec() == QDialog.Accepted:
            QSettings().setValue("agreement_accepted", True)

    def show_about_dialog(self):
        dialog = AboutDialog(self)
        dialog.exec()

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
            self.statusBar().showMessage(message, 5000)
        QApplication.processEvents()

    def check_for_updates(self):
        try:
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept': 'application/json',
                'Cache-Control': 'no-cache'
            }

            response = requests.get(
                AppConfig.UPDATE_CHECK_URL,
                timeout=15,
                verify=False,
                headers=headers,
                allow_redirects=True
            )

            if response.status_code == 200:
                try:
                    data = json.loads(response.text)
                    latest_version = data.get('latest_version')
                    filename = data.get('filename')
                    
                    if (latest_version and filename and 
                        self.version_to_tuple(latest_version) > self.version_to_tuple(AppConfig.VERSION)):
                        self.show_update_notification(filename)
                    else:
                        self.update_label.clear()
                        
                except json.JSONDecodeError:
                    self.log_message("Ошибка: Некорректный JSON-формат в ответе сервера", status=True)
            
            elif response.status_code == 403:
                error_msg = "Доступ запрещен. Проверьте настройки сервера."
                self.log_message(error_msg, status=True)
            else:
                self.log_message(f"Ошибка сервера: {response.status_code}", status=True)
                
        except requests.exceptions.RequestException as e:
            self.log_message(f"Ошибка подключения: {str(e)}", status=True)
            
        except Exception as e:
            self.log_message(f"Неизвестная ошибка: {str(e)}", status=True)

    def version_to_tuple(self, version_str):
        return tuple(map(int, version_str.split('.')))

    def show_update_notification(self, filename):
        download_url = f"{AppConfig.BASE_DOWNLOAD_URL}{filename}"
        link_text = f'<a href="{download_url}" style="text-decoration:none; color:#28a745;">Доступно обновление: {filename}</a>'
        self.update_label.setText(link_text)

    def open_update_url(self, link):
        QDesktopServices.openUrl(QUrl(link))

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
    window = FileConverterApp()
    window.show()
    sys.exit(app.exec())