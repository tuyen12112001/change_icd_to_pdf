
# emergency_stop.py
import threading

class EmergencyStopManager:
    def __init__(self):
        self._stop_event = threading.Event()

    def trigger_stop(self):
        """Kích hoạt dừng khẩn cấp."""
        self._stop_event.set()
        print("⚠ 非常停止 được kích hoạt!")

    def reset(self):
        """Reset trạng thái dừng (nếu cần chạy lại)."""
        self._stop_event.clear()

    def is_stop_requested(self) -> bool:
        """Kiểm tra xem có yêu cầu dừng không."""
        return self._stop_event.is_set()


# Tạo một instance toàn cục để dùng chung
emergency_manager = EmergencyStopManager()
