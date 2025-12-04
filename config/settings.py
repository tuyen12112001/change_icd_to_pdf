
# config/settings.py
import os
import sys

def _get_base_dir():
    """
    Trả về thư mục gốc của project:
    - Khi chạy bình thường: thư mục chứa main.py
    - Khi chạy PyInstaller: thư mục tạm MEIPASS
    """
    if getattr(sys, 'frozen', False):  # PyInstaller
        return sys._MEIPASS
    # Khi chạy từ source, base là thư mục chứa main.py
    # settings.py nằm trong config/, nên lấy 2 cấp lên
    return os.path.dirname(os.path.dirname(__file__))

BASE_DIR = _get_base_dir()

# Thư mục assets cùng cấp với app/process/utils/config/icon
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
ICON_DIR   = os.path.join(BASE_DIR, "icon")


ICON_PATH = os.path.join(ICON_DIR, "icon.ico")  
IMAGE1_PATH = os.path.join(ASSETS_DIR, "1.png")
IMAGE2_PATH = os.path.join(ASSETS_DIR, "2.png")  



# Màu sắc chủ đạo
BG_COLOR = "#e6f2ff"
PANEL_BG = "#f9f9f9"
STATUS_INFO_COLOR = "blue"
STATUS_WARN_COLOR = "orange"
STATUS_SUCCESS_COLOR = "green"
STATUS_ERROR_COLOR = "red"

# UI text
APP_TITLE = "出図ツール ver 2.0"
HEADER_TEXT = "出図ツール"

# Progress bar default length
PROGRESS_LENGTH = 600
