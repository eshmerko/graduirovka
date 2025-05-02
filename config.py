# Copyright (c) 2025 Шмерко Евгений Леонидович
# SPDX-License-Identifier: MIT

class AppConfig:
    PROGRAM_SLUG = "contab"
    # Версия приложения
    VERSION = "0.0.2"
    
    # Лицензия
    LICENSE_INFO = {
        "name": "MIT License",
        "year": 2025,
        "copyright_holder": "Шмерко Евгений Леонидович",
        "spdx_identifier": "MIT",
        "notice": """Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software..."""
    }

    # Настройки обновлений
    UPDATE_CHECK_URL = "https://eshmerko.com/api/check-update/"
    BASE_DOWNLOAD_URL = "https://eshmerko.com/downloads/"

    
    # Информация о разработчике
    COMPANY_NAME = "ОАО «Пуховичинефтепродукт»"
    DEVELOPER_NAME = "Шмерко Евгений Леонидович"
    DEVELOPER_EMAIL = "e.shmerko@beloil.by"
    DEVELOPER_PHONE = "+375 44 7777710"
    
    # Текст для UI
    APP_NAME = "Извлечение данных из градуировочных таблиц для импорта в 1С"
    
    @classmethod
    def license_header(cls):
        return (
            f"Copyright (c) {cls.LICENSE_INFO['year']} {cls.LICENSE_INFO['copyright_holder']}\n"
            f"SPDX-License-Identifier: {cls.LICENSE_INFO['spdx_identifier']}"
        )