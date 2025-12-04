
import ctypes
from ctypes import wintypes

# Khai báo các hàm/kiểu dữ liệu từ user32.dll
user32 = ctypes.WinDLL('user32', use_last_error=True)

# Định nghĩa prototype cho EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

user32.EnumWindows.argtypes = (EnumWindowsProc, wintypes.LPARAM)
user32.EnumWindows.restype = wintypes.BOOL

user32.IsWindowVisible.argtypes = (wintypes.HWND,)
user32.IsWindowVisible.restype = wintypes.BOOL

user32.GetWindowTextW.argtypes = (wintypes.HWND, wintypes.LPWSTR, ctypes.c_int)
user32.GetWindowTextW.restype = ctypes.c_int

user32.GetWindowTextLengthW.argtypes = (wintypes.HWND,)
user32.GetWindowTextLengthW.restype = ctypes.c_int

def get_open_window_titles(include_invisible=False, unique=True):
    """
    Trả về danh sách tiêu đề (title) của các cửa sổ ứng dụng đang mở trên Windows.

    Parameters
    ----------
    include_invisible : bool
        True  -> bao gồm cả cửa sổ không hiển thị (IsWindowVisible == False),
        False -> chỉ lấy cửa sổ hiển thị.
    unique : bool
        True  -> loại bỏ trùng lặp tên cửa sổ,
        False -> giữ nguyên toàn bộ.

    Returns
    -------
    list[str]
        Danh sách tiêu đề cửa sổ (chuỗi Unicode).
    """
    titles = []

    def _enum_cb(hwnd, lparam):
        # Lọc theo hiển thị (nếu cần)
        if not include_invisible and not user32.IsWindowVisible(hwnd):
            return True  # tiếp tục liệt kê

        # Lấy độ dài tiêu đề
        length = user32.GetWindowTextLengthW(hwnd)
        if length == 0:
            return True

        # Cấp bộ đệm và lấy chuỗi tiêu đề
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buffer, length + 1)
        title = buffer.value.strip()

        if title:
            titles.append(title)

        return True  # tiếp tục liệt kê các cửa sổ khác

    # Gọi EnumWindows với callback
    cb_func = EnumWindowsProc(_enum_cb)
    if not user32.EnumWindows(cb_func, 0):
        raise ctypes.WinError(ctypes.get_last_error())

    if unique:
        # Giữ thứ tự xuất hiện, loại bỏ trùng lặp
        seen = set()
        ordered_unique = []
        for t in titles:
            if t not in seen:
                seen.add(t)
                ordered_unique.append(t)
        return ordered_unique
    else:
        return titles


if __name__ == "__main__":
    # Ví dụ sử dụng
    for t in get_open_window_titles():
        print(t)
