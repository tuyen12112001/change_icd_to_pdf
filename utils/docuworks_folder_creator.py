
# utils/docuworks_folder_creator.py
import time
import pyautogui
import pyperclip
import pygetwindow as gw
from utils.emergency_stop import emergency_manager
# An toàn & ổn định
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.2

def _open_new_folder_dialog():
    """Alt+F → N → F để mở hộp thoại tạo thư mục mới (một lần)."""
    pyautogui.keyDown('alt')
    pyautogui.press('f')
    pyautogui.keyUp('alt')
    time.sleep(0.25)
    pyautogui.press('n')
    time.sleep(0.25)
    pyautogui.press('f')
    time.sleep(0.4)

def _paste_and_confirm(name_text):
    """Paste tên vào ô nhập và Enter (Ctrl+A để làm sạch trước khi dán)."""
    pyperclip.copy(name_text)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.05)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(0.6)

def _dismiss_popup_with_enter():
    """Đóng popup (OK) bằng Enter."""
    pyautogui.press('enter')
    time.sleep(0.4)

def _is_popup_active(title_keyword: str = "docuworks", screen_ratio_threshold: float = 0.20) -> bool:
    """
    Heuristic phát hiện popup mà KHÔNG cần lấy danh sách cửa sổ:
      - Nếu title của active window KHÔNG chứa từ khóa (mặc định 'DocuWorks') ⇒ popup.
      - Hoặc diện tích active window nhỏ hơn tỷ lệ màn hình (mặc định 20%) ⇒ popup.
    """
    try:
        active = gw.getActiveWindow()
        if not active:
            return False

        title = (active.title or "").lower()
        if title_keyword and (title_keyword.lower() not in title) and (len(title) > 0):
            return True  # title khác DocuWorks ⇒ rất có thể là dialog/popup

        # Bổ sung: kiểm tra tỷ lệ kích thước so với màn hình để tăng độ chắc chắn
        screen_w, screen_h = pyautogui.size()
        screen_area = max(1, screen_w * screen_h)
        active_area = max(1, active.width * active.height)
        area_ratio = active_area / float(screen_area)

        if area_ratio < screen_ratio_threshold:
            return True

        return False
    except Exception:
        # Nếu không đọc được active window, an toàn hơn là coi như KHÔNG popup
        return False

def create_docuworks_folder_unique(base_name: str, ensure_docuworks_running, max_attempts: int = 20):
    """
    Tạo thư mục trong DocuWorks với tên `base_name`.
    Nếu trùng tên: đóng popup → ngay trong cùng dialog, nối thêm '1' vào tên hiện tại và thử lại
      (abc → abc1 → abc11 → ...), cho tới khi thành công hoặc vượt max_attempts.

    Yêu cầu:
      - Truyền vào ensure_docuworks_running (hàm của bạn; hàm này đã đảm bảo DocuWorks chạy/active).

    Returns:
      - Tên cuối cùng đã tạo (str) hoặc None nếu thất bại.
    """
    # 1) Đảm bảo DocuWorks sẵn sàng theo logic của bạn
    try:
        ok = ensure_docuworks_running()
        if not ok:
            print("❌ DocuWorks chưa sẵn sàng.")
            return None
    except Exception as e:
        print(f"❌ Lỗi ensure DocuWorks: {e}")
        return None

    # 2) Mở hộp thoại tạo thư mục (một lần)
    _open_new_folder_dialog()

    # 3) Thử tạo và xử lý trùng bằng cách nối thêm '1' trong cùng dialog
    attempt_name = base_name
    for _ in range(max_attempts):
        
        if emergency_manager.is_stop_requested():
            print("⚠ 非常停止が押されたため、DocuWorksフォルダ作成を中断します。")
            return None

        _paste_and_confirm(attempt_name)

        # Dùng heuristic title + tỷ lệ diện tích màn hình để phát hiện popup
        if _is_popup_active(title_keyword="docuworks", screen_ratio_threshold=0.20):
            _dismiss_popup_with_enter()
            time.sleep(0.2)
            attempt_name = attempt_name + "1"
            continue

        print(f"✅ DocuWorksで新しいフォルダを作成しました: {attempt_name}")
        return attempt_name

    print("❌ 既存フォルダ名との重複により、最大試行回数を超えました。作成に失敗しました。")
