
# utils/ui_helpers.py
import tkinter as tk

def blink_widget(widget, times=3, color="#ffcccc", interval=200):
    original_color = widget.cget("bg")
    def toggle(count):
        if count > 0:
            new_color = color if count % 2 == 0 else original_color
            widget.config(bg=new_color)
            widget.after(interval, toggle, count - 1)
        else:
            widget.config(bg=original_color)
    toggle(times * 2)

def update_error_box(app, message, status="error", exclusive_pairs=(("warning", "info"),)):
    """
    Cập nhật error_box theo 'status' và đảm bảo mỗi loại chỉ giữ 1 thông điệp.
    Nếu status thuộc cặp loại trừ, xóa tag của đối nghịch (vd: warning xoá info).
    """
    def _update():
        app.error_box.config(state=tk.NORMAL)

        # Xóa đoạn mang cùng tag 'status'
        ranges = app.error_box.tag_ranges(status)
        for i in range(0, len(ranges), 2):
            app.error_box.delete(ranges[i], ranges[i+1])

        # Xóa tag đối nghịch nếu nằm trong cặp loại trừ
        opposite = None
        for a, b in exclusive_pairs:
            if status == a:
                opposite = b
                break
            if status == b:
                opposite = a
                break

        if opposite:
            opp_ranges = app.error_box.tag_ranges(opposite)
            for i in range(0, len(opp_ranges), 2):
                app.error_box.delete(opp_ranges[i], opp_ranges[i+1])

        # Icon theo trạng thái
        icons = {"error": "❌", "success": "✅", "info": "ℹ️", "warning": "⚠️"}
        icon = icons.get(status, "•")
        text = f"{icon} {message}\n"

        # Chèn và gán tag
        app.error_box.insert("end", text, status)

        # Hiệu ứng nhấp nháy nhẹ theo trạng thái
        if status == "error":
            blink_widget(app.error_box, color="#ffcccc")
        elif status == "success":
            blink_widget(app.error_box, color="#ccffcc")

        # Khóa lại & cuộn xuống
        app.error_box.config(state=tk.DISABLED)
        app.error_box.see("end")

    app.after(0, _update)

def log_error(app, msg):   update_error_box(app, msg, status="error")
def log_success(app, msg): update_error_box(app, msg, status="success")
def log_info(app, msg):    update_error_box(app, msg, status="info")
def log_warning(app, msg): update_error_box(app, msg, status="warning")

def animate_loading(app, base_text="処理中", dots=3, interval=500):
    if not hasattr(app, "loading_count"):
        app.loading_count = 0
    app.loading_count = (app.loading_count + 1) % (dots + 1)
    text = base_text + "." * app.loading_count
    app.status_label.config(text=text)
    if getattr(app, "is_running", False):
        app.loading_job = app.after(interval, lambda: animate_loading(app, base_text, dots, interval))

def stop_loading(app):
    if hasattr(app, "loading_job"):
        app.after_cancel(app.loading_job)
    app.is_running = False

def update_status(app, text, progress, color="blue"):
    app.status_label.config(text=text, fg=color)
    app.progress["value"] = progress
