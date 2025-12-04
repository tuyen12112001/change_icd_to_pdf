import os
import subprocess
import time
import pygetwindow as gw

# ================================
# 🔍 Tìm shortcut trong Start Menu
# ================================
def find_shortcut(app_name):
    """
    Tìm shortcut của ứng dụng trong Start Menu.
    Trả về đường dẫn .lnk nếu tìm thấy, None nếu không.
    """
    start_menu = r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
    for root, dirs, files in os.walk(start_menu):
        for file in files:
            if app_name.lower() in file.lower() and file.endswith(".lnk"):
                return os.path.join(root, file)
    return None

# ================================
# ✅ Kiểm tra & mở DocuWorks
# ================================
def ensure_docuworks_running():
    """
    Kiểm tra DocuWorks có chạy không, nếu chưa thì mở từ shortcut.
    Sau đó thử active cửa sổ.
    """
    windows = [w for w in gw.getWindowsWithTitle('DocuWorks Desk') if w.title.startswith('DocuWorks')]
    if not windows:
        shortcut = find_shortcut("DocuWorks")
        if shortcut:
            print(f"🔄 DocuWorks chưa mở. Đang khởi động từ shortcut: {shortcut}")
            subprocess.Popen(['cmd', '/c', shortcut])
            time.sleep(5)
        else:
            print("❌ Không tìm thấy shortcut DocuWorks trong Start Menu.")
            return False

    windows = [w for w in gw.getWindowsWithTitle('DocuWorks Desk') if w.title.startswith('DocuWorks')]
    if windows:
        win = windows[0]
        win.restore()  # Khôi phục nếu bị thu nhỏ
        win.maximize()
        win.activate()
        time.sleep(0.8)
        print("✅ DocuWorks đã được active.")
        return True
    else:
        print("❌ Không thể active DocuWorks.")
        return False

# ================================
# ✅ Kiểm tra & mở ICAD
# ================================

def ensure_icad_running(icad_path):
    """
    Kiểm tra ICAD có đang chạy không, nếu chưa thì khởi động và active cửa sổ.
    """
    icad_windows = gw.getWindowsWithTitle('Micro Caelum')
    if not icad_windows:
        print("🔄 ICAD chưa mở, đang khởi động...")
        subprocess.Popen([icad_path])
        time.sleep(5)  # Chờ ứng dụng mở
        icad_windows = gw.getWindowsWithTitle('Micro Caelum')  # Kiểm tra lại sau khi mở

    if icad_windows:
        try:
            win = icad_windows[0]
            win.restore()  # Khôi phục nếu bị thu nhỏ
            win.maximize()
            win.activate()
            time.sleep(0.8)
            print("✅ ICAD đã được active.")
            return True
        except Exception as e:
            print(f"❌ Lỗi khi active ICAD: {e}")
            return False
    else:
        print("❌ Không tìm thấy cửa sổ ICAD sau khi khởi động.")
        return False