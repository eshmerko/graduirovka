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
from datetime import datetime
from PySide6.QtCore import (
    Qt, QPropertyAnimation, QEasingCurve, QThread, Signal, 
    QUrl, QSettings, QTimer, QDateTime
)
from PySide6.QtGui import (
    QFont, QPixmap, QColor, QLinearGradient, QBrush, 
    QIcon, QPainter, QAction, QDesktopServices
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QFileDialog, QLabel, QLineEdit, QPushButton, QTextEdit, QStatusBar,
    QMessageBox, QDialog, QScrollArea, QScrollBar
)
from striprtf.striprtf import rtf_to_text
import win32com.client as win32

from config import AppConfig

import xlrd
from xlrd import open_workbook

class StartupScreen(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()
        
    def setup_ui(self):
        self.setWindowTitle("Добро пожаловать")
        self.setFixedSize(600, 400)
        
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
        content.setStyleSheet("font-size: 12pt; color: #4a4a4a;")
        
        scroll_area.setWidget(content)
        scroll_area.setWidgetResizable(True)
        
        # Кнопка принятия
        accept_btn = QPushButton("Принять и продолжить")
        accept_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 12px 24px;
                border-radius: 6px;
                font-size: 12pt;
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
        """)
    def accept(self):
        QSettings().setValue("agreement_accepted", True)
        super().accept()

    def get_content_text(self):
        return """
        <h3>📋 Поддерживаемые форматы файлов:</h3>
        <ul>
            <li>Microsoft Word (.docx)</li>
            <li>Microsoft Word 97-2003 (.doc)</li>
            <li>Rich Text Format (.rtf)</li>
            <li>Microsoft Excel (.xls .xlsx)</li>
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
        <p>Данная программа предоставляется <b>'как есть'</b>, без каких-либо гарантий. 
        Разработчик не несет ответственности за:</p>
        <ul>
            <li>Прямой или косвенный ущерб, вызванный использованием программы</li>
            <li>Потерю данных или их некорректную обработку</li>
            <li>Проблемы совместимости с конкретными версиями ПО</li>
        </ul>
        <p>Используя данное программное обеспечение, вы соглашаетесь с этими условиями.</p>
        """

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
        self.current_version = "0.0.2"  # Замените на вашу версию
        self.update_check_url = "https://eshmerko.com/api/check-update/"
        self.update_info = None
        self.settings = QSettings("YourCompany", "YourApp")
        # Проверка соглашения при запуске
        self.check_agreement()
        # Добавим константы для настройки столбцов
        self.LEFT_COLS = (1, 2)   # B и C (0-based)
        self.RIGHT_COLS = (5, 6)  # F и G

    def check_agreement(self):
        """Проверка принятия пользовательского соглашения"""
        settings = QSettings()
        if not settings.value("agreement_accepted", False, type=bool):
            self.show_agreement_dialog()

    def process_excel_data(self, input_path, output_path):
        """Основной метод обработки Excel файлов"""
        try:
            self.log_message("Начало обработки Excel файла...", status=True)
            
            wb = open_workbook(input_path)
            all_data = []
            
            for sheet in wb.sheets():
                self.log_message(f"Обработка листа: {sheet.name}")
                self.process_excel_sheet(sheet, all_data)
            
            self.export_excel_data(all_data, output_path)
            return True
            
        except Exception as e:
            self.log_message(f"Ошибка обработки Excel: {str(e)}", status=True)
            return False

    def is_valid_excel_row(self, sheet, row, cols):
        """Проверка валидности данных в строке"""
        try:
            # Проверка целого числа в первом столбце
            int(sheet.cell_value(row, cols[0]))
            
            # Проверка числа с плавающей точкой во втором столбце
            float(sheet.cell_value(row, cols[1]))
            
            return True
        except (ValueError, TypeError):
            return False

    def process_excel_sheet(self, sheet, all_data):
        """Обработка данных из листа Excel"""
        for row_idx in range(sheet.nrows):
            # Обработка левых столбцов
            if self.is_valid_excel_row(sheet, row_idx, self.LEFT_COLS):
                level = int(sheet.cell_value(row_idx, self.LEFT_COLS[0]))
                capacity = float(sheet.cell_value(row_idx, self.LEFT_COLS[1]))
                formatted = f"{capacity:.15f}".rstrip('0').rstrip('.')
                all_data.append((level, formatted))
                self.log_message(f"Обработана строка {row_idx+1} (левые столбцы): {level} ~ {formatted}")

            # Обработка правых столбцов
            if self.is_valid_excel_row(sheet, row_idx, self.RIGHT_COLS):
                level = int(sheet.cell_value(row_idx, self.RIGHT_COLS[0]))
                capacity = float(sheet.cell_value(row_idx, self.RIGHT_COLS[1]))
                formatted = f"{capacity:.15f}".rstrip('0').rstrip('.')
                all_data.append((level, formatted))
                self.log_message(f"Обработана строка {row_idx+1} (правые столбцы): {level} ~ {formatted}")

    def export_excel_data(self, data, filename):
        """Экспорт данных с сортировкой и удалением дубликатов"""
        unique_data = {}
        for level, cap in data:
            unique_data[level] = cap
        
        sorted_data = sorted(unique_data.items(), key=lambda x: x[0])
        
        with open(filename, 'w', encoding='utf-8') as f:
            for level, cap in sorted_data:
                f.write(f"{level}~{cap}\n")
        
        success_msg = (f"Успешно обработано записей: {len(sorted_data)}!\n"
                      f"Результат сохранен: {os.path.abspath(filename)}")
        self.log_message(success_msg, status=True)
        QMessageBox.information(self, "Успех", success_msg)

            
    def show_agreement_dialog(self):
        """Показ диалога с соглашением"""
        dialog = StartupScreen(self)
        if dialog.exec() == QDialog.Accepted:
            QSettings().setValue("agreement_accepted", True)
        else:
            # Если пользователь не принял соглашение
            QMessageBox.warning(
                self,
                "Требуется согласие",
                "Для использования программы необходимо принять условия соглашения",
            )
            QTimer.singleShot(0, self.close)

        # Настройка главного окна
        self.setWindowTitle("File Converter Pro")
        self.setGeometry(100, 100, 800, 600)
        self.setMinimumSize(700, 500)
        
        # Инициализация интерфейса
        self.setup_ui()
        self.setup_connections()
        self.setup_menu()

        # Запуск проверки обновлений
        QTimer.singleShot(2000, self.check_for_updates)

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        # Верхняя панель
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 10)
        
        # Кнопка обновления
        self.update_btn = QPushButton()
        self.update_btn.setVisible(False)
        self.update_btn.setIcon(QIcon(":/icons/update.svg"))
        self.update_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 8px 16px;
                border-radius: 4px;
                font-size: 11pt;
                border: none;
                min-width: 180px;
            }
            QPushButton:hover { background-color: #45a049; }
            QPushButton:pressed { background-color: #3d8b40; }
        """)
        self.update_btn.clicked.connect(self.open_update_page)

        # Кнопка "О программе"
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
            QPushButton:hover { background-color: #e9ecef; }
            QPushButton:pressed { background-color: #dee2e6; }
        """)
        self.about_btn.setFixedSize(120, 30)

        # Правая часть шапки
        header_right = QHBoxLayout()
        header_right.setSpacing(15)
        header_right.addWidget(self.update_btn)
        header_right.addWidget(self.about_btn)
        
        header_layout.addStretch()
        header_layout.addLayout(header_right)

        # Область выбора файлов
        file_layout = QVBoxLayout()
        file_layout.addLayout(self.create_file_row("Исходный файл:", "Выбрать...", True))
        file_layout.addLayout(self.create_file_row("Результирующий файл:", "Сохранить как...", False))

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
            QPushButton:hover { background-color: #45a049; }
            QPushButton:pressed { background-color: #3d8b40; }
        """)

        # Сборка интерфейса
        main_layout.addLayout(header_layout)
        main_layout.addLayout(file_layout)
        main_layout.addWidget(self.log_area, 1)
        main_layout.addWidget(self.convert_btn)

    def create_file_row(self, label_text, button_text, is_input):
        layout = QHBoxLayout()
        label = QLabel(label_text)
        entry = QLineEdit()
        entry.setPlaceholderText("Укажите путь к файлу..." if is_input else "Результат.txt")
        entry.setMinimumWidth(400)
        
        browse_btn = QPushButton(button_text)
        browse_btn.setFixedSize(100, 30)
        
        if is_input:
            self.input_entry = entry
            self.browse_input_btn = browse_btn
        else:
            self.output_entry = entry
            self.browse_output_btn = browse_btn
        
        layout.addWidget(label)
        layout.addWidget(entry, 1)
        layout.addWidget(browse_btn)
        return layout

    def setup_menu(self):
        menu_bar = self.menuBar()
        help_menu = menu_bar.addMenu("Справка")
        
        manual_action = QAction("Показать инструкцию", self)
        manual_action.triggered.connect(self.show_manual)
        help_menu.addAction(manual_action)
        
        about_action = QAction("О программе", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)

    def setup_connections(self):
        self.browse_input_btn.clicked.connect(self.select_input_file)
        self.browse_output_btn.clicked.connect(self.select_output_file)
        self.convert_btn.clicked.connect(self.process_file)
        self.about_btn.clicked.connect(self.show_about_dialog)

    def check_for_updates(self):
        try:
            headers = {
                'User-Agent': f'FileConverterPro/{self.current_version}',
                'Accept': 'application/json'
            }
            
            # Исправленный URL запроса
            response = requests.get(
                f"https://eshmerko.com/api/check-update/contab/{self.current_version}/",
                headers=headers,
                timeout=10,
                verify=False
            )
                
            if response.status_code == 200:
                data = response.json()
                if data.get('update_available', False):
                    self.handle_update_available(data)
                else:
                    self.handle_no_updates()
            else:
                self.log_message(f"Ошибка проверки обновлений: {response.status_code}")
                
        except Exception as e:
            self.log_message(f"Ошибка при проверке обновлений: {str(e)}")
            if "Excel" in str(e):
                self.log_message("Проверьте установку Microsoft Excel", status=True)

    def handle_update_available(self, data):
        self.update_info = data
        self.update_btn.setText(f"Обновление до {data['latest_version']}")
        self.update_btn.setVisible(True)
        
        # Показать уведомление
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Information)
        msg.setWindowTitle("Доступно обновление")
        msg.setText(f"""
            <b>Доступна новая версия {data['latest_version']}!</b>
            <p>Дата выпуска: {data['release_date']}</p>
            <p>Изменения: {data['changelog']}</p>
            <p>Перейти на страницу загрузки?</p>
        """)
        
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg.button(QMessageBox.Yes).setText("Перейти")
        msg.button(QMessageBox.No).setText("Позже")
        
        if msg.exec() == QMessageBox.Yes:
            self.open_update_page()

    def open_update_page(self):
        if self.update_info and self.update_info.get('download_url'):
            QDesktopServices.openUrl(QUrl(self.update_info['download_url']))
            self.log_message("Открыта страница загрузки обновления")
        else:
            QMessageBox.warning(self, "Ошибка", "Ссылка на обновление недоступна")

    def handle_no_updates(self):
        self.update_btn.setVisible(False)
        self.update_info = None
        self.log_message("У вас актуальная версия программы")

    def show_manual(self):
        if not self.settings.value("agreement_accepted", False):
            startup_dialog = StartupScreen(self)
            if startup_dialog.exec() == QDialog.Accepted:
                self.settings.setValue("agreement_accepted", True)

    def select_input_file(self):
        file_filter = ("Документы (*.docx *.doc *.rtf);;"
                "Excel файлы (*.xls *.xlsx);;"
                "Все файлы (*.*)")
        filename, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл", "", file_filter)
        if filename:
            self.input_entry.setText(filename)

    def select_output_file(self):
        filename, _ = QFileDialog.getSaveFileName(
            self, "Сохранить результат", "", "Текстовые файлы (*.txt)")
        if filename:
            self.output_entry.setText(filename)

    def log_message(self, message, status=False):
        timestamp = QDateTime.currentDateTime().toString("hh:mm:ss")
        self.log_area.append(f"[{timestamp}] {message}")
        if status:
            self.statusBar().showMessage(message, 5000)
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
        return None
    
    def sanitize_filename(self, filename):
        """Очищает имя файла от запрещенных символов и нормализует пробелы."""
        forbidden_chars = r'[\\/*?:"<>|]'
        sanitized = re.sub(forbidden_chars, '_', filename)
        sanitized = sanitized.strip()
        sanitized = re.sub(r'[\s_]+', '_', sanitized)
        return sanitized

    def process_file(self):
        input_path = self.input_entry.text().strip()
        output_path = self.output_entry.text().strip()
        
        # Проверка наличия входного файла
        if not input_path:
            QMessageBox.critical(self, "Ошибка", "Пожалуйста, выберите исходный файл!")
            return
        
        # Получение расширения файла
        file_ext = os.path.splitext(input_path)[1].lower()
        
        # Обработка Excel файлов
        if file_ext in ('.xls', '.xlsx'):
            # Установка имени выходного файла по умолчанию
            if not output_path:
                output_path = "результат.txt"
                self.output_entry.setText(output_path)
            
            # Проверка и создание директории
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                try:
                    os.makedirs(output_dir)
                except OSError as e:
                    self.log_message(f"Ошибка создания директории: {str(e)}", status=True)
                    QMessageBox.critical(
                        self,
                        "Ошибка",
                        f"Невозможно создать директорию: {output_dir}"
                    )
                    return
            
            # Очистка лога и начало обработки
            self.log_area.clear()
            self.log_message("=== Начало обработки Excel файла ===", status=True)
            
            try:
                # Основная обработка Excel
                success = self.process_excel_data(input_path, output_path)
                if not success:
                    QMessageBox.critical(self, "Ошибка", "Ошибка обработки Excel файла!")
            
            except Exception as e:
                error_msg = f"Критическая ошибка: {str(e)}"
                self.log_message(error_msg, status=True)
                QMessageBox.critical(self, "Ошибка", error_msg)
        
        # Обработка Word/RTF файлов
        else:
            # Проверка расширения файла
            valid_extensions = ('.docx', '.doc', '.rtf')
            if not input_path.lower().endswith(valid_extensions):
                QMessageBox.critical(
                    self,
                    "Ошибка",
                    "Неподдерживаемый формат файла. Выберите файл с расширением .docx, .doc, .rtf, .xls или .xlsx."
                )
                return
            
            # Обработка имени выходного файла
            if not output_path:
                output_path = "результат.txt"
            else:
                dirname = os.path.dirname(output_path)
                filename = os.path.basename(output_path)
                filename_part, ext = os.path.splitext(filename)
                sanitized_filename = self.sanitize_filename(filename_part)
                if not ext:
                    ext = '.txt'
                ext = ext.lower()
                output_path = os.path.join(dirname, f"{sanitized_filename}{ext}")
            
            # Проверка и создание директории
            output_dir = os.path.dirname(output_path)
            if output_dir and not os.path.exists(output_dir):
                try:
                    os.makedirs(output_dir)
                except OSError as e:
                    self.log_message(f"Ошибка создания директории: {str(e)}", status=True)
                    QMessageBox.critical(
                        self,
                        "Ошибка",
                        f"Невозможно создать директорию: {output_dir}"
                    )
                    return
            
            self.output_entry.setText(output_path)
            self.log_area.clear()
            self.log_message("=== Начало обработки ===", status=True)
            
            try:
                # Конвертация в RTF (если нужно)
                if input_path.lower().endswith('.rtf'):
                    rtf_path = input_path
                else:
                    self.log_message("Конвертация в RTF...", status=True)
                    rtf_path = self.convert_to_rtf(input_path)
                    if not rtf_path:
                        raise ValueError("Не удалось конвертировать файл в RTF")
                
                # Обработка RTF
                data = self.process_rtf(rtf_path)
                
                # Удаление временного файла
                if rtf_path != input_path:
                    try:
                        os.remove(rtf_path)
                    except Exception as e:
                        self.log_message(f"[ВНИМАНИЕ] Не удалось удалить временный файл: {str(e)}")
                
                if not data:
                    self.log_message("Не найдено подходящих данных!", status=True)
                    QMessageBox.warning(self, "Предупреждение", "В файле не найдено подходящих данных!")
                    return
                
                # Сортировка и сохранение
                data.sort(key=lambda x: float(x.split('~')[0]))
                
                with open(output_path, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(data))
                
                success_msg = (f"Успешно обработано записей: {len(data)}!\n"
                            f"Результат сохранен: {os.path.abspath(output_path)}")
                self.log_message(success_msg, status=True)
                QMessageBox.information(self, "Успех", success_msg)
            
            except Exception as e:
                error_msg = f"Критическая ошибка: {str(e)}"
                self.log_message(error_msg, status=True)
                QMessageBox.critical(self, "Ошибка", error_msg)

    def convert_to_rtf(self, input_path):
        try:
            if not os.path.exists(input_path):
                self.log_message(f"[ОШИБКА] Файл не найден: {input_path}", status=True)
                return None

            pythoncom.CoInitialize()
            word = None
            doc = None
            temp_path = None
            file_created = False  # Флаг успешного создания файла

            try:
                # Инициализация Word
                try:
                    word = win32.Dispatch("Word.Application")
                    word.Visible = False
                    word.DisplayAlerts = False
                except Exception as e:
                    self.log_message("[ОШИБКА] Не удалось инициализировать Microsoft Word", status=True)
                    raise RuntimeError("Ошибка инициализации Word") from e

                # Попытки открытия документа
                for attempt in range(1, 4):
                    try:
                        doc = word.Documents.Open(
                            FileName=os.path.abspath(input_path),
                            ConfirmConversions=False,
                            ReadOnly=True,
                            AddToRecentFiles=False,
                            PasswordDocument=""
                        )
                        if doc:
                            break
                    except Exception as e:
                        if attempt == 3:
                            error_msg = f"Не удалось открыть документ после 3 попыток: {str(e)}"
                            if "The document is locked" in str(e):
                                error_msg += "\nФайл заблокирован для редактирования!"
                            raise RuntimeError(error_msg) from e
                        time.sleep(1.5)

                # Проверка успешности открытия
                if not doc:
                    raise RuntimeError("Документ не был открыт")

                # Создание временного файла
                temp_dir = tempfile.gettempdir()
                sanitized_name = self.sanitize_filename(os.path.basename(input_path))
                temp_path = os.path.join(temp_dir, f"temp_{sanitized_name}.rtf")

                # Сохранение документа
                try:
                    doc.SaveAs(temp_path, FileFormat=6)
                    file_created = True
                    self.log_message(f"Успешно создан временный файл: {temp_path}")
                except Exception as e:
                    raise RuntimeError(f"Ошибка сохранения RTF: {str(e)}") from e

                return temp_path

            except Exception as e:
                error_msg = f"[ОШИБКА] Конвертация: {str(e)}"
                self.log_message(error_msg, status=True)
                return None

            finally:
                try:
                    # Закрытие документа и Word
                    if doc:
                        doc.Close(SaveChanges=False)
                    if word:
                        word.Quit()
                    
                    # Удаление временного файла только при ошибке
                    if temp_path and os.path.exists(temp_path) and not file_created:
                        try:
                            os.remove(temp_path)
                        except Exception as remove_error:
                            self.log_message(f"[WARNING] Ошибка удаления файла: {str(remove_error)}")
                    
                    pythoncom.CoUninitialize()
                    
                except Exception as cleanup_error:
                    self.log_message(f"[ОШИБКА] Очистка ресурсов: {str(cleanup_error)}")
                    if "RPC_E_CALL_REJECTED" in str(cleanup_error):
                        self.log_message("[WARNING] Попробуйте перезапустить приложение", status=True)

        except Exception as outer_error:
            self.log_message(f"[ОШИБКА] Внешняя ошибка: {str(outer_error)}", status=True)
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


    def show_about_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("О программе")
        dialog.setFixedSize(400, 300)
        
        layout = QVBoxLayout(dialog)
        content = QLabel("""
            <h2>File Converter Pro</h2>
            <p>Версия: {}</p>
            <p>Разработчик: Шмерко Евгений Леонидович</p>
            <p>Веб-сайт: <a href="eshmerko.com">eshmerko.com</a></p>
            <p>Лицензия: MIT</p>
            <p>© 2025 Все права защищены</p>
        """.format(self.current_version))
        
        layout.addWidget(content)
        dialog.exec()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = FileConverterApp()
    window.show()
    sys.exit(app.exec())