import ctypes
import os
import shutil
import subprocess
import sys
import time
import zipfile
from urllib.parse import urlencode
import winreg 

import requests
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QVBoxLayout,
    QLabel,
    QProgressBar,
    QWidget,
    QMessageBox,
    QSpacerItem,
    QSizePolicy
)

# Встроенные QSS стили
LIGHT_THEME_QSS = """
/* light_theme.qss */

/* Общие стили для всех виджетов */
QWidget {
    background-color: #ffffff; /* Чисто белый фон для современного вида */
    color: #333333; /* Темно-серый текст для лучшей читаемости */
    font-family: "Segoe UI", "Helvetica Neue", "Arial";
    font-size: 10pt;
}

/* Стили для QLabel */
QLabel {
    color: #333333;
    font-weight: normal;
}

/* Стили для QProgressBar */
QProgressBar {
    border: 1px solid #cccccc;
    border-radius: 5px;
    text-align: center;
    height: 20px;
}

QProgressBar::chunk {
    background-color: #0078d7; /* Акцентный синий цвет для заполненной части */
    border-radius: 5px;
}

/* Стили для QMainWindow */
QMainWindow {
    background-color: #ffffff;
}

/* Стили для QMessageBox */
QMessageBox {
    background-color: #ffffff;
    color: #333333;
}

QPushButton {
    background-color: #e7e7e7;
    border: 1px solid #b0b0b0;
    border-radius: 5px;
    padding: 6px 12px;
    color: #333333;
}

QPushButton:hover {
    background-color: #d4d4d4;
    border-color: #999999;
}

QPushButton:pressed {
    background-color: #c0c0c0;
    border-color: #7a7a7a;
}

/* Дополнительные стили для ссылок в QLabel */
QLabel a {
    color: #0078d7;
    text-decoration: underline;
}

QLabel a:hover {
    color: #005a9e;
}
"""

DARK_THEME_QSS = """
/* dark_theme.qss */

/* Общие стили для всех виджетов */
QWidget {
    background-color: #2d2d2d; /* Немного светлее для улучшенной читаемости */
    color: #e0e0e0; /* Светло-серый текст для контраста */
    font-family: "Segoe UI", "Helvetica Neue", "Arial";
    font-size: 10pt;
}

/* Стили для QLabel */
QLabel {
    color: #e0e0e0;
    font-weight: normal;
}

/* Стили для QProgressBar */
QProgressBar {
    border: 1px solid #555555;
    border-radius: 5px;
    text-align: center;
    height: 20px;
    background-color: #3a3a3a;
}

QProgressBar::chunk {
    background-color: #1e90ff; /* Акцентный синий цвет для заполненной части */
    border-radius: 5px;
}

/* Стили для QMainWindow */
QMainWindow {
    background-color: #2d2d2d;
}

/* Стили для QMessageBox */
QMessageBox {
    background-color: #2d2d2d;
    color: #e0e0e0;
}

QPushButton {
    background-color: #3a3a3a;
    border: 1px solid #555555;
    border-radius: 5px;
    padding: 6px 12px;
    color: #e0e0e0;
}

QPushButton:hover {
    background-color: #505050;
    border-color: #777777;
}

QPushButton:pressed {
    background-color: #606060;
    border-color: #888888;
}

/* Дополнительные стили для ссылок в QLabel */
QLabel a {
    color: #1e90ff;
    text-decoration: underline;
}

QLabel a:hover {
    color: #1c86ee;
}
"""

# Тексты для локализации
texts = {
    'ru': {
        'window_title': "Обновление программы [Версия Loader v1.2]",
        'header': "Обновление/Переустановка программы до новой версии",
        'status_initializing': "Инициализация...",
        'download_label': "Процесс обновления:",
        'extract_label': "Распаковка обновления:",
        'status_extracting': "Распаковка обновления...",
        'status_download_complete': "Загрузка завершена. Начинается распаковка...",
        'status_extract_complete': "Распаковка завершена.",
        'update_success': "Обновление прошло успешно",
        'update_error': "Ошибка при обновлении: {e}",
        'error_title': "Ошибка",
        'admin_error': "Не удалось запустить с правами администратора: {e}",
        'update_title': "Обновление",
        'error_info': "Если произошла ошибка при обновлении, то установите программу вручную",
        'github_download': "Загрузить с GitHub",
        'github_update': "Исходный код GitHub Update"
    },
    'en': {
        'window_title': "Program Update [Version Loader v1.2]",
        'header': "Updating/Reinstalling the program to the new version",
        'status_initializing': "Initializing...",
        'download_label': "Update process:",
        'extract_label': "Extracting update:",
        'status_extracting': "Extracting update...",
        'status_download_complete': "Download completed. Starting extraction...",
        'status_extract_complete': "Extraction completed.",
        'update_success': "Update was successful",
        'update_error': "Update error: {e}",
        'error_title': "Error",
        'admin_error': "Failed to run as administrator: {e}",
        'update_title': "Update",
        'error_info': "If an error occurs during the update, install the program manually",
        'github_download': "Download from GitHub",
        'github_update': "Source code GitHub update"
    }
}

def get_system_language():
    try:
        kernel32 = ctypes.windll.kernel32
        lang_id = kernel32.GetUserDefaultUILanguage()
        if lang_id == 0x0419:
            return 'ru'
        else:
            return 'en'
    except Exception:
        return 'en'

language = get_system_language()

def check_dpi_penguin_installed(extract_to):
    """Проверка наличия программы 'DPI Penguin.exe' в директории с Loader.exe"""
    return os.path.isfile(os.path.join(extract_to, "D.txt"))

class UpdateWorker(QThread):
    progress_download = pyqtSignal(int)
    progress_extract = pyqtSignal(int)
    update_finished = pyqtSignal(bool, str)

    def __init__(self, public_key, download_path, extract_to, main_exe, updater_exe, texts):
        super().__init__()
        self.public_key = public_key
        self.download_path = download_path
        self.extract_to = extract_to
        self.main_exe = main_exe
        self.updater_exe = updater_exe
        self.texts = texts

    def run(self):
        try:
            if not check_dpi_penguin_installed(self.extract_to):
                raise Exception(f"Программа 'DPI Penguin.exe' не найдена в папке {self.extract_to}")

            self.stop_service("WinDivert")
            self.terminate_process("winws.exe")
            self.terminate_process("goodbyedpi.exe")
            self.terminate_process("DPI Penguin.exe")
            self.delete_files()
            self.download_update()
            self.extract_zip()
            self.download_and_replace_config() 
            os.remove(self.download_path)
            self.restart_main_app()
            self.update_finished.emit(True, self.texts['update_success'])
        except Exception as e:
            error_message = self.texts['update_error'].format(e=e)
            self.update_finished.emit(False, error_message)

    def terminate_process(self, process_name):
        try:
            import psutil
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'].lower() == process_name.lower():
                    try:
                        proc.terminate()
                        proc.wait(timeout=5)
                    except psutil.NoSuchProcess:
                        pass
                    except psutil.TimeoutExpired:
                        proc.kill()
                    except Exception:
                        pass
        except ImportError:
            pass

    def stop_service(self, service_name):
        try:
            import win32serviceutil
            import win32service
            import winerror

            service_status = win32serviceutil.QueryServiceStatus(service_name)
            if service_status[1] == win32service.SERVICE_RUNNING:
                win32serviceutil.StopService(service_name)
                win32serviceutil.WaitForServiceStatus(service_name, win32service.SERVICE_STOPPED, timeout=30)
        except ImportError:
            pass
        except win32service.error as e:
            if e.winerror != winerror.ERROR_SERVICE_DOES_NOT_EXIST:
                raise e
        except Exception:
            pass

    def delete_files(self):
        base_path = self.extract_to
        updater_name = os.path.basename(self.updater_exe).lower()  # Имя текущего обновляющего файла
        exe_name = "DPI Penguin.exe".lower()  # Имя основной программы
        internal_folder = "_internal"  # Папка, которую нужно удалить

    # Проходим по файлам в директории
        for item in os.listdir(base_path):
            item_path = os.path.join(base_path, item)
            
            # Удаляем только exe-файл программы и папку _internal
            if item.lower() == updater_name or item.lower() == exe_name:
                try:
                    if os.path.isfile(item_path) or os.path.islink(item_path):
                        os.remove(item_path)  # Удаляем файл .exe или ссылку
                    elif os.path.isdir(item_path):
                        shutil.rmtree(item_path)  # Удаляем папку
                except Exception as e:
                    print(f"Ошибка при удалении файла {item}: {e}")
            elif item.lower() == internal_folder:
                # Удаляем папку _internal, если она существует
                try:
                    shutil.rmtree(item_path)  # Удаляем папку _internal
                except Exception as e:
                    print(f"Ошибка при удалении папки {internal_folder}: {e}")

    def download_update(self):
        base_url = 'https://cloud-api.yandex.net/v1/disk/public/resources/download?'
        final_url = base_url + urlencode({'public_key': self.public_key})
        response = requests.get(final_url)
        response.raise_for_status()
        download_url = response.json()['href']
        response = requests.get(download_url, stream=True)
        response.raise_for_status()
        total_length = response.headers.get('content-length')

        if total_length is None:
            with open(self.download_path, 'wb') as f:
                f.write(response.content)
            self.progress_download.emit(100)
        else:
            total_length = int(total_length)
            downloaded = 0
            chunk_size = 8192
            with open(self.download_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=chunk_size):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        percent = int(downloaded * 100 / total_length)
                        self.progress_download.emit(percent)
        time.sleep(1)

    def extract_zip(self):
        with zipfile.ZipFile(self.download_path, 'r') as zip_ref:
            members = zip_ref.namelist()
            total_files = len(members)
            for i, member in enumerate(members, 1):
                zip_ref.extract(member, self.extract_to)
                percent = int(i * 100 / total_files)
                self.progress_extract.emit(percent)
        time.sleep(1)

    def download_and_replace_config(self):
        """Загрузка и замена файла version_config.ini без отображения прогресса"""
        try:
            config_url = 'https://raw.githubusercontent.com/zhivem/DPI-Penguin/main/setting_version/version_config.ini'
            # Определение пути к version_config.ini относительно основного пути обновления
            config_relative_path = os.path.join("_internal", "setting_version", "version_config.ini")
            config_path = os.path.join(self.extract_to, config_relative_path)

            # Загрузка файла
            response = requests.get(config_url)
            response.raise_for_status()

            # Создание необходимых директорий, если они отсутствуют
            os.makedirs(os.path.dirname(config_path), exist_ok=True)

            # Запись файла
            with open(config_path, 'wb') as f:
                f.write(response.content)
        except Exception as e:
            raise Exception(f"Не удалось обновить конфигурационный файл: {e}")

    def restart_main_app(self):
        subprocess.Popen([self.main_exe])

class UpdateWindow(QMainWindow):
    def __init__(self, public_key, main_exe, updater_exe, texts, theme='light'):
        super().__init__()

        self.texts = texts
        if hasattr(sys, "_MEIPASS"):
            # Когда приложение запущено как .exe
            icon_path = os.path.join(sys._MEIPASS, "update_reset.ico")
        else:
            # Когда приложение запущено как скрипт
            icon_path = "update_reset.ico"
            
        self.setWindowTitle(self.texts['window_title'])
        self.setWindowIcon(QIcon(icon_path))
        self.setFixedSize(500, 260) 
        self.public_key = public_key
        self.main_exe = main_exe
        self.updater_exe = updater_exe

        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        self.download_path = os.path.join(base_path, 'update.zip')
        self.extract_to = base_path

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(15, 15, 15, 15)

        self.header_label = QLabel(self.texts['header'])
        self.header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        label_font = QFont("Arial", 11)
        self.header_label.setFont(label_font)
        self.header_label.setWordWrap(True)
        layout.addWidget(self.header_label)

        self.status_label = QLabel(self.texts['status_initializing'])
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        status_font = QFont("Arial", 10)
        self.status_label.setFont(status_font)
        layout.addWidget(self.status_label)

        self.download_label = QLabel(self.texts['download_label'])
        self.download_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        download_label_font = QFont("Arial", 10)
        self.download_label.setFont(download_label_font)
        layout.addWidget(self.download_label)

        self.progress_download = QProgressBar()
        self.progress_download.setRange(0, 100)
        self.progress_download.setValue(0)
        layout.addWidget(self.progress_download)

        self.extract_label = QLabel(self.texts['extract_label'])
        self.extract_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        extract_label_font = QFont("Arial", 10)
        self.extract_label.setFont(extract_label_font)
        layout.addWidget(self.extract_label)

        self.progress_extract = QProgressBar()
        self.progress_extract.setRange(0, 100)
        self.progress_extract.setValue(0)
        layout.addWidget(self.progress_extract)

        layout.addItem(QSpacerItem(20, 10, QSizePolicy.Policy.Minimum, QSizePolicy.Policy.Expanding))

        error_info_text = (
            f"{self.texts['error_info']} "
            f"<a href='https://github.com/zhivem/DPI-Penguin/releases' "
            f"style='color: #0078d7; text-decoration: underline;'>"
            f"{self.texts['github_download']}</a><br>"
            f"<a href='https://github.com/zhivem/Loader-for-DPI-Penguin' "
            f"style='color: #0078d7; text-decoration: underline;'>"
            f"{self.texts['github_update']}</a>"
        )
        self.error_info_label = QLabel(error_info_text)
        self.error_info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.error_info_label.setOpenExternalLinks(True)
        self.error_info_label.setWordWrap(True)
        self.error_info_label.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(self.error_info_label)

        central_widget.setLayout(layout)

        self.worker = UpdateWorker(
            public_key=self.public_key,
            download_path=self.download_path,
            extract_to=self.extract_to,
            main_exe=self.main_exe,
            updater_exe=self.updater_exe,
            texts=self.texts
        )
        self.worker.progress_download.connect(self.update_download_progress)
        self.worker.progress_extract.connect(self.update_extract_progress)
        self.worker.update_finished.connect(self.on_update_finished)
        self.worker.start()

    def update_download_progress(self, percent):
        self.status_label.clear()
        self.progress_download.setValue(percent)
        if percent >= 100:
            self.status_label.setText(self.texts['status_download_complete'])

    def update_extract_progress(self, percent):
        self.status_label.setText(self.texts['status_extracting'])
        self.progress_extract.setValue(percent)
        if percent >= 100:
            self.status_label.setText(self.texts['status_extract_complete'])

    def on_update_finished(self, success, message):
        if success:
            QMessageBox.information(self, self.texts['update_title'], message)
            self.close()
        else:
            QMessageBox.critical(self, self.texts['error_title'], message)
            self.close()

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin(texts):
    try:
        executable = sys.executable
        if getattr(sys, 'frozen', False):
            script = sys.executable
        else:
            script = os.path.abspath(__file__)
        params = f'"{script}"'
        ctypes.windll.shell32.ShellExecuteW(None, "runas", executable, params, None, 1)
    except Exception as e:
        app = QApplication.instance()
        if not app:
            app = QApplication(sys.argv)
        error_message = texts['admin_error'].format(e=e)
        QMessageBox.critical(None, texts['error_title'], error_message)
        sys.exit(1)

def get_system_theme():
    """Автоматическое определение системной темы на Windows 10+."""
    try:
        registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        key = winreg.OpenKey(registry, key_path)
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        winreg.CloseKey(key)
        return 'light' if value == 1 else 'dark'
    except Exception as e:
        print(f"Не удалось определить системную тему: {e}")
        return 'light'  # По умолчанию

def load_stylesheet(app, theme='light'):
    """Применяет QSS стиль к приложению."""
    try:
        if theme == 'dark':
            stylesheet = DARK_THEME_QSS
        else:
            stylesheet = LIGHT_THEME_QSS
        app.setStyleSheet(stylesheet)
    except Exception as e:
        print(f"Не удалось применить стиль: {e}")

def main():
    if not is_admin():
        run_as_admin(texts[language])
        sys.exit(0)

    public_key = 'https://disk.yandex.ru/d/ckFPOTqcG7XsTg'
    main_exe = 'DPI Penguin.exe'

    if getattr(sys, 'frozen', False):
        updater_exe = sys.executable
    else:
        updater_exe = os.path.abspath(__file__)

    app = QApplication(sys.argv)
    
    # Автоматическое определение темы
    theme = get_system_theme()  # 'light' или 'dark'
    
    load_stylesheet(app, theme)

    window = UpdateWindow(public_key, main_exe, updater_exe, texts[language], theme)
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
