import os
import subprocess
import time
import pygetwindow as gw

try:
    import pyautogui
    import win32clipboard
except Exception:
    pyautogui = None
    win32clipboard = None

try:
    import win32con
except Exception:
    win32con = None

# ================================
# 🔍 Tìm shortcut trong Start Menu
# ================================
DOCUWORKS_EXE_CANDIDATES = [
    r"C:\Program Files (x86)\FUJIFILM\DocuWorks\bin\dwdesk.exe",
    r"C:\Program Files (x86)\Fuji Xerox\DocuWorks\bin\dwdesk.exe",
]


def find_dwdesk_exe():
    for path in DOCUWORKS_EXE_CANDIDATES:
        if os.path.isfile(path):
            return path
    return None


def open_docuworks_target_file_via_exe(target_file: str) -> bool:
    if not target_file or not os.path.isfile(target_file):
        print("❌ Target file does not exist.")
        return False

    exe = find_dwdesk_exe()
    if not exe:
        print("❌ Không tìm thấy dwdesk.exe trong đường dẫn mặc định.")
        return False

    try:
        cmd = f'"{exe}" /f "{target_file}"'
        subprocess.Popen(cmd, shell=True)
        print(f"✅ Opened DocuWorks file: {target_file}")
        return True
    except Exception:
        return False


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
def _get_path_from_clipboard():
    if not (win32clipboard and win32con):
        return None
    path = None
    try:
        win32clipboard.OpenClipboard()
        if win32clipboard.IsClipboardFormatAvailable(win32con.CF_HDROP):
            data = win32clipboard.GetClipboardData(win32con.CF_HDROP)
            if data:
                path = data[0]
        elif win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
            text = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
            if text:
                path = text.strip()
    except Exception:
        pass
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass
    return path


def _clear_clipboard():
    if not win32clipboard:
        return
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
    except Exception:
        pass
    finally:
        try:
            win32clipboard.CloseClipboard()
        except Exception:
            pass


def get_docuworks_open_folder_path_ctrl_t():
    if not pyautogui:
        return None
    try:
        active_win = gw.getActiveWindow()
        if not (active_win and active_win.title.startswith("DocuWorks")):
            return None
        _clear_clipboard()
        time.sleep(0.1)
        pyautogui.hotkey("ctrl", "t")
        time.sleep(0.3)
        return _get_path_from_clipboard()
    except Exception:
        return None


def resolve_user_folder_path(current_path: str, user_folder_name: str = "ユーザーフォルダ", target_subfolder: str = "PDF"):
    """
    Resolve path to ユーザーフォルダ/PDF
    Creates the subfolder if it doesn't exist.
    """
    if not current_path:
        return None

    norm_path = os.path.normpath(current_path)
    parts = norm_path.split(os.sep)

    # Find ユーザーフォルダ in the path
    if user_folder_name in parts:
        idx = parts.index(user_folder_name)
        base_path = os.sep.join(parts[:idx + 1])
        target_path = os.path.join(base_path, target_subfolder)
        # Create if doesn't exist
        if not os.path.isdir(target_path):
            try:
                os.makedirs(target_path, exist_ok=True)
                print(f"✅ Created folder: {target_path}")
            except Exception as e:
                print(f"⚠ Could not create folder: {e}")
        return target_path

    # Search upward for ユーザーフォルダ
    probe = norm_path
    while True:
        candidate = os.path.join(probe, user_folder_name)
        if os.path.isdir(candidate):
            target_path = os.path.join(candidate, target_subfolder)
            # Create if doesn't exist
            if not os.path.isdir(target_path):
                try:
                    os.makedirs(target_path, exist_ok=True)
                    print(f"✅ Created folder: {target_path}")
                except Exception as e:
                    print(f"⚠ Could not create folder: {e}")
            return target_path

        parent = os.path.dirname(probe)
        if parent == probe:
            break
        probe = parent

    return None


def ensure_pdf_folder_exists_from_current_path(current_path: str, user_folder_name: str = "ユーザーフォルダ", target_subfolder: str = "PDF"):
    """
    Từ path hiện tại (lấy từ Ctrl+T), tìm ユーザーフォルダ và đảm bảo thư mục PDF tồn tại.
    Trả về đường dẫn PDF nếu thành công, ngược lại trả về None.
    """
    target_path = resolve_user_folder_path(
        current_path,
        user_folder_name=user_folder_name,
        target_subfolder=target_subfolder,
    )
    if not target_path:
        print(f"⚠ Không tìm thấy folder '{user_folder_name}' trong đường dẫn hiện tại.")
        return None

    if os.path.isdir(target_path):
        print(f"✅ PDF folder already exists: {target_path}")
        return target_path

    print(f"📂 PDF folder does not exist. Creating: {target_path}")
    try:
        os.makedirs(target_path, exist_ok=True)
        print(f"✅ Successfully created PDF folder: {target_path}")
        return target_path
    except Exception as e:
        print(f"❌ Failed to create PDF folder: {e}")
        return None


def open_docuworks_folder_by_path(target_path: str, assume_running: bool = False, ensure_selected: bool = True) -> bool:
    if not target_path or not os.path.isdir(target_path):
        print("❌ Target folder does not exist.")
        return False

    if not assume_running:
        ok = ensure_docuworks_running()
        if not ok:
            return False

    if not pyautogui:
        print("❌ pyautogui is not available.")
        return False

    folder_name = os.path.basename(target_path.rstrip("\\/"))
    if not folder_name:
        print("❌ Invalid target folder name.")
        return False

    try:
        for _ in range(3):
            # Try to focus folder tree/list and jump by typing the name
            for _ in range(3):
                pyautogui.press("f6")
                time.sleep(0.2)

            pyautogui.typewrite(folder_name, interval=0.03)
            time.sleep(0.2)
            pyautogui.press("enter")
            time.sleep(0.4)

            if ensure_selected:
                pyautogui.press("space")
                time.sleep(0.1)

            current_path = get_docuworks_open_folder_path_ctrl_t()
            if current_path and os.path.normcase(os.path.normpath(current_path)) == os.path.normcase(os.path.normpath(target_path)):
                print("✅ Activated DocuWorks folder.")
                return True

        print("⚠ Could not activate the target folder after retries.")
        return False
    except Exception:
        return False


def open_docuworks_user_folder_from_current():
    current_path = get_docuworks_open_folder_path_ctrl_t()
    if not current_path:
        print("⚠ Không lấy được path hiện tại từ Ctrl+T.")
        return False

    user_folder_path = ensure_pdf_folder_exists_from_current_path(current_path)
    if not user_folder_path:
        print("⚠ Không tìm thấy folder 'PDF' trong đường dẫn hiện tại.")
        return False

    print(f"✅ Resolved PDF folder path: {user_folder_path}")
    return open_docuworks_folder_by_path(user_folder_path, assume_running=True, ensure_selected=True)


def close_docuworks():
    """Close all DocuWorks windows."""
    windows = [w for w in gw.getWindowsWithTitle('DocuWorks') if w.title.startswith('DocuWorks')]
    for win in windows:
        try:
            win.close()
        except Exception:
            pass
    # Wait for windows to close
    for _ in range(10):
        time.sleep(0.5)
        windows = [w for w in gw.getWindowsWithTitle('DocuWorks') if w.title.startswith('DocuWorks')]
        if not windows:
            print("✅ DocuWorks đã đóng.")
            return True
    print("⚠ DocuWorks có thể chưa đóng hoàn toàn.")
    return False


def open_docuworks_with_folder(folder_path: str) -> bool:
    """Open DocuWorks directly at the specified folder using /f flag."""
    if not folder_path or not os.path.isdir(folder_path):
        print("❌ Folder path does not exist.")
        return False

    exe = find_dwdesk_exe()
    if not exe:
        print("❌ Không tìm thấy dwdesk.exe.")
        return False

    try:
        cmd = f'"{exe}" /f "{folder_path}"'
        subprocess.Popen(cmd, shell=True)
        print(f"✅ Opened DocuWorks at: {folder_path}")
        return True
    except Exception as e:
        print(f"❌ Lỗi khi mở DocuWorks: {e}")
        return False


def activate_docuworks_and_open_user_folder(max_attempts: int = 3):
    """
    Activate DocuWorks and open user PDF folder.
    Also delete all PDF files from My Documents after folder selection.
    """
    # Step 1: Ensure DocuWorks is running so Ctrl+T works
    ok = ensure_docuworks_running(print_active_folder=False)
    if not ok:
        return False

    # Step 2: Read current path via Ctrl+T (retry a few times)
    current_path = None
    for _ in range(max_attempts):
        current_path = get_docuworks_open_folder_path_ctrl_t()
        if current_path:
            break
        time.sleep(0.4)

    if not current_path:
        print("⚠ Không lấy được path hiện tại từ Ctrl+T.")
        return False

    # Step 3: Từ path ban đầu, đảm bảo folder PDF tồn tại trước khi khởi động lại DocuWorks
    target_path = ensure_pdf_folder_exists_from_current_path(
        current_path,
        user_folder_name="ユーザーフォルダ",
        target_subfolder="PDF",
    )
    if not target_path:
        return False

    print(f"✅ Confirmed PDF folder path: {target_path}")

    # Step 4: Delete all PDF files from My Documents after folder selection
    print("🗑️ フォルダ選択後にMy Documents内のPDFファイルを削除します...")
    from utils.cleanup_pdf import delete_all_pdf_in_my_documents
    success, deleted_count, message = delete_all_pdf_in_my_documents()
    if success:
        print(f"✅ {message}")
    else:
        print(f"⚠️ PDF削除失敗: {message}")

    # Step 5: Close DocuWorks and re-open at the target folder
    closed = close_docuworks()
    if not closed:
        print("⚠ DocuWorks may not have closed cleanly; attempting to open anyway.")
    time.sleep(0.5)

    opened = open_docuworks_with_folder(target_path)
    if opened:
        print(f"✅ DocuWorks opened at: {target_path}")
    else:
        print("❌ Failed to open DocuWorks at the target folder.")
    return opened


def open_docuworks_at_user_folder_via_exe():
    current_path = get_docuworks_open_folder_path_ctrl_t()
    if not current_path:
        print("⚠ Không lấy được path hiện tại từ Ctrl+T.")
        return False

    user_folder_path = ensure_pdf_folder_exists_from_current_path(current_path)
    if not user_folder_path:
        print("⚠ Không tìm thấy folder 'PDF' trong đường dẫn hiện tại.")
        return False

    exe = find_dwdesk_exe()
    if not exe:
        print("❌ Không tìm thấy dwdesk.exe trong đường dẫn mặc định.")
        return False

    try:
        cmd = f'"{exe}" /f "{user_folder_path}"'
        subprocess.Popen(cmd, shell=True)
        print(f"✅ Opened DocuWorks at: {user_folder_path}")
        return True
    except Exception:
        return False


def ensure_docuworks_running(print_active_folder: bool = False, enforce_user_folder: bool = False):
    """
    Kiểm tra DocuWorks có chạy không, nếu chưa thì mở từ shortcut.
    Sau đó thử active cửa sổ.
    """
    windows = [w for w in gw.getWindowsWithTitle('DocuWorks') if w.title.startswith('DocuWorks')]
    if not windows:
        shortcut = find_shortcut("DocuWorks")
        if shortcut:
            print(f"🔄 DocuWorks chưa mở. Đang khởi động từ shortcut: {shortcut}")
            subprocess.Popen(['cmd', '/c', shortcut])
            # Wait for window to appear (up to 15s)
            for _ in range(15):
                time.sleep(1)
                windows = [w for w in gw.getWindowsWithTitle('DocuWorks') if w.title.startswith('DocuWorks')]
                if windows:
                    break
        else:
            print("❌ Không tìm thấy shortcut DocuWorks trong Start Menu.")
            return False

    windows = [w for w in gw.getWindowsWithTitle('DocuWorks') if w.title.startswith('DocuWorks')]
    if windows:
        win = windows[0]
        win.restore()  # Khôi phục nếu bị thu nhỏ
        win.maximize()
        win.activate()
        time.sleep(0.8)
        print("✅ DocuWorks đã được active.")

        if print_active_folder:
            open_path = get_docuworks_open_folder_path_ctrl_t()
            if open_path:
                print(f"📁 DocuWorks open folder path: {open_path}")
            else:
                print("⚠ Không lấy được path folder từ Ctrl+T.")

        if enforce_user_folder:
            open_docuworks_user_folder_from_current()
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