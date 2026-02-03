
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
    """error_box 内のすべてのコンテンツを削除します"""
    def _clear():
        app.error_box.config(state=tk.NORMAL)
        app.error_box.delete("1.0", "end")
        app.error_box.config(state=tk.DISABLED)
    
    app.after(0, _clear)

def update_error_box(app, message, status="error", exclusive_pairs=(("warning", "info"),)):
    """
    error_box を更新 - 新しいメッセージのみを追加し、古いメッセージは削除しないでください
    """
    def _update():
        app.error_box.config(state=tk.NORMAL)

        # アイコンの種類
        icons = {"error": "❌", "success": "✅", "info": "ℹ️", "warning": "⚠️"}
        icon = icons.get(status, "•")
        text = f"{icon} {message}\n"

        # タグを挿入して割り当てる
        app.error_box.insert("end", text, status)

        # ステータスに応じたライトの点滅効果
        if status == "error":
            blink_widget(app.error_box, color="#ffcccc")
        elif status == "success":
            blink_widget(app.error_box, color="#ccffcc")

        # ロックして下にスクロール
        app.error_box.config(state=tk.DISABLED)
        app.error_box.see("end")

    app.after(0, _update)

def update_file_comparison_message(app, message, status="error"):
    """
    XDWファイルとICDファイルの比較通知を管理します
    これらの通知のうち1つだけを残し、新しい通知が追加された場合は古い通知を削除します
    成功した場合は、「はい/いいえ」ボタンも削除します
    
    Args:
        app: ShutsuzuuApp instance
        message: 通知内容
        status: "warning" (ファイル数が一致しません), hoặc "info" (処理が完了)
    """
    def _update():
        app.error_box.config(state=tk.NORMAL)

        # 古いファイル比較通知をすべてクリア + はい/いいえボタン
        for tag in ["error", "success", "warning", "info"]:
            ranges = app.error_box.tag_ranges(tag)
            for i in range(0, len(ranges), 2):
                content = app.error_box.get(ranges[i], ranges[i+1])
                if ("xdwファイル数" in content or "処理が完了しました" in content or 
                    "コピーされたXDWファイルを削除しますか" in content):
                    app.error_box.delete(ranges[i], ranges[i+1])
        
        # 情報ステータスのメッセージが表示されたら、はい/いいえボタン（error_box の子ウィンドウ）を削除します。
        if status == "info":
            # error_box 内のすべてのウィンドウ（ボタン）を検索して削除します
            for widget_name in app.error_box.window_names():
                app.error_box.delete(widget_name)

        # アイコンの種類
        icons = {"error": "❌", "success": "✅", "warning": "⚠️", "info": "ℹ️"}
        icon = icons.get(status, "•")
        text = f"{icon} {message}\n"

        # 新しいメッセージを挿入
        app.error_box.insert("end", text, status)

        # 点滅効果
        if status == "warning":
            blink_widget(app.error_box, color="#ffcccc")
        elif status == "info":
            blink_widget(app.error_box, color="#ccffff")

        # ロックして下にスクロール
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
    ユーザーが XDW ファイルを削除するかどうかを選択できるように、エラー ボックスに「はい」「いいえ」の 2 つのボタンを追加します。
    
    Args:
        app: ShutsuzuuApp instance
        on_yes_callback: ユーザーが「削除する」を押した際に呼び出される関数
        on_no_callback: ユーザーが「削除しない」を押した際に呼び出される関数
    """
    def _add_buttons():
        app.error_box.config(state=tk.NORMAL)
        
        # ボタンの前にプロンプ​​トを追加する
        app.error_box.insert("end", "\n📋 コピーされたXDWファイルを削除しますか?\n", "info")
        
        # 2つのボタンを含むフレームを作成する
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
        
        # 空白行を挿入してボタンフレームを埋め込む
        app.error_box.insert("end", "")
        app.error_box.window_create("end", window=button_frame)
        app.error_box.insert("end", "\n")
        
        app.error_box.config(state=tk.DISABLED)
        app.error_box.see("end")
    
    app.after(0, _add_buttons)

def add_delete_pdf_buttons(app, on_yes_callback, on_no_callback):
    """
    ユーザーが PDF ファイルを削除するかどうかを選択できるように、エラー ボックスに「はい」「いいえ」の 2 つのボタンを追加します。
    
    Args:
        app: ShutsuzuuApp instance
        on_yes_callback: ユーザーが「削除する」を押した際に呼び出される関数
        on_no_callback: ユーザーが「削除しない」を押した際に呼び出される関数
    """
    def _add_buttons():
        app.error_box.config(state=tk.NORMAL)
        
        # ボタンの前にプロンプ​​トを追加する
        app.error_box.insert("end", "\n📋 コピーされたPDFファイルを削除しますか?\n", "info")
        
        # 2つのボタンを含むフレームを作成する
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
        
        # 空白行を挿入してボタンフレームを埋め込む
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

