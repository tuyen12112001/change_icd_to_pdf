
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

def clear_error_box(app):
    """Xóa toàn bộ nội dung trong error_box"""
    def _clear():
        app.error_box.config(state=tk.NORMAL)
        app.error_box.delete("1.0", "end")
        app.error_box.config(state=tk.DISABLED)
    
    app.after(0, _clear)

def update_error_box(app, message, status="error", exclusive_pairs=(("warning", "info"),)):
    """
    Cập nhật error_box - chỉ thêm thông báo mới, không xóa cũ
    """
    def _update():
        app.error_box.config(state=tk.NORMAL)

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

def update_file_comparison_message(app, message, status="error"):
    """
    Quản lý thông báo so sánh file XDW và ICD
    Chỉ giữ 1 thông báo loại này - xóa cái cũ nếu có thêm cái mới
    Nếu success, cũng xóa các button Yes/No
    
    Args:
        app: ShutsuzuuApp instance
        message: Nội dung thông báo
        status: "warning" (ファイル数が一致しません), hoặc "info" (処理が完了)
    """
    def _update():
        app.error_box.config(state=tk.NORMAL)

        # Xóa tất cả thông báo so sánh file cũ + button Yes/No
        for tag in ["error", "success", "warning", "info"]:
            ranges = app.error_box.tag_ranges(tag)
            for i in range(0, len(ranges), 2):
                content = app.error_box.get(ranges[i], ranges[i+1])
                # Chỉ xóa nếu là thông báo so sánh file (chứa keyword) hoặc hỏi xóa file
                if ("xdwファイル数" in content or "処理が完了しました" in content or 
                    "コピーされたXDWファイルを削除しますか" in content):
                    app.error_box.delete(ranges[i], ranges[i+1])
        
        # Nếu success, xóa luôn các button Yes/No (child windows trong error_box)
        if status == "info":
            # Tìm và xóa tất cả windows (button) trong error_box
            for widget_name in app.error_box.window_names():
                app.error_box.delete(widget_name)

        # Icon theo trạng thái
        icons = {"error": "❌", "success": "✅", "warning": "⚠️", "info": "ℹ️"}
        icon = icons.get(status, "•")
        text = f"{icon} {message}\n"

        # Chèn thông báo mới
        app.error_box.insert("end", text, status)

        # Hiệu ứng nhấp nháy
        if status == "warning":
            blink_widget(app.error_box, color="#ffcccc")
        elif status == "info":
            blink_widget(app.error_box, color="#ccffff")

        # Khóa lại & cuộn xuống
        app.error_box.config(state=tk.DISABLED)
        app.error_box.see("end")

    app.after(0, _update)

def log_error(app, msg):   
    update_error_box(app, msg, status="error")

def log_success(app, msg): 
    update_error_box(app, msg, status="success")

def log_info(app, msg):    
    update_error_box(app, msg, status="info")

def log_warning(app, msg): 
    update_error_box(app, msg, status="warning")

def add_delete_xdw_buttons(app, on_yes_callback, on_no_callback):
    """
    Thêm 2 button "Yes" "No" vào error box để user chọn xóa file XDW hay không
    
    Args:
        app: ShutsuzuuApp instance
        on_yes_callback: Function gọi khi user ấn Yes
        on_no_callback: Function gọi khi user ấn No
    """
    def _add_buttons():
        app.error_box.config(state=tk.NORMAL)
        
        # Thêm dòng hỏi trước button
        app.error_box.insert("end", "\n📋 コピーされたXDWファイルを削除しますか?\n", "info")
        
        # Tạo frame chứa 2 button
        button_frame = tk.Frame(app.error_box, bg="white")
        
        yes_btn = tk.Button(
            button_frame,
            text="削除する",
            command=on_yes_callback,
            bg="#90EE90",
            fg="black",
            width=15,
            padx=5,
            pady=5,
            cursor="hand2"
        )
        yes_btn.pack(side=tk.LEFT, padx=5)
        
        no_btn = tk.Button(
            button_frame,
            text="削除しない",
            command=on_no_callback,
            bg="#FFB6C6",
            fg="black",
            width=15,
            padx=5,
            pady=5,
            cursor="hand2"
        )
        no_btn.pack(side=tk.LEFT, padx=5)
        
        # Chèn dòng trống rồi embed button frame
        app.error_box.insert("end", "")
        app.error_box.window_create("end", window=button_frame)
        app.error_box.insert("end", "\n")
        
        app.error_box.config(state=tk.DISABLED)
        app.error_box.see("end")
    
    app.after(0, _add_buttons)

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
