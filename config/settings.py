
# config/settings.py
import os
import sys

def _get_base_dir():
    """
    プロジェクトのルートディレクトリを返します。
    - 通常実行時: main.py があるディレクトリ
    - PyInstaller 実行時: 一時ディレクトリ MEIPASS
    """
    if getattr(sys, 'frozen', False):  # PyInstaller
        return sys._MEIPASS
    # settings.py は config/ にあるので、2 階層上に移動します。
    return os.path.dirname(os.path.dirname(__file__))

BASE_DIR = _get_base_dir()
# アセットフォルダはapp/process/utils/config/iconと同じレベルにあります
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
ICON_DIR   = os.path.join(BASE_DIR, "icon")


ICON_PATH = os.path.join(ICON_DIR, "icon.ico")  
IMAGE3_PATH = os.path.join(ASSETS_DIR, "3.png")



# メインカラー
BG_COLOR = "#e6f2ff"
PANEL_BG = "#f9f9f9"
STATUS_INFO_COLOR = "blue"
STATUS_WARN_COLOR = "orange"
STATUS_SUCCESS_COLOR = "green"
STATUS_ERROR_COLOR = "red"

# UI text
APP_TITLE = "出図ツール ver 2.1"
HEADER_TEXT = "出図ツール"

# DocuWorks target folder (adjust as needed)
DOCUWORKS_TARGET_FOLDER = r"C:\Users\510-tuyen\OneDrive - 明和工業株式会社\ドキュメント\Fuji Xerox\DocuWorks\DWFolders\ユーザーフォルダ\PDF保存場所_1"


# Progress bar default length
PROGRESS_LENGTH = 600
