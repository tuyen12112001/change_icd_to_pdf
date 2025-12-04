
# app/main_app.py
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES

from config.settings import (
    ICON_PATH, BG_COLOR, PANEL_BG,
    APP_TITLE, HEADER_TEXT, PROGRESS_LENGTH,
)
from utils.UI_helpers import (
    blink_widget, update_error_box,
    animate_loading, stop_loading, update_status,
)
from process.process_manager import ProcessManager


class ShutsuzuuApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("800x550")
        self.configure(bg=BG_COLOR)

        # Icon
        try:
            self.iconbitmap(ICON_PATH)
        except Exception:
            pass  # Không chặn chạy nếu icon lỗi

        # ウィンドウを最前面に
        self.attributes('-topmost', True)
        self.lift()
        self.focus_force()

        # --- Khởi tạo trạng thái trước ---
        self.info = None
        self.excel_full_path = ""
        self.is_running = False

        # --- KHỞI TẠO PROCESS MANAGER TRƯỚC KHI TẠO UI ---
        from process.process_manager import ProcessManager
        self.process_manager = ProcessManager(self)

        # Tạo UI
        self._build_ui()

        # 保存情報 & trạng thái
        self.info = None
        self.excel_full_path = ""
        self.is_running = False

    def _build_ui(self):
        # ヘッダー
        header = tk.Label(self, text=HEADER_TEXT, font=("Arial", 24, "bold"), fg="#004080", bg=BG_COLOR)
        header.pack(pady=15)

        # Excelファイル入力
        excel_frame = tk.Frame(self, bg=PANEL_BG, bd=2, relief="groove")
        excel_frame.pack(pady=10, padx=20, fill="x")
        tk.Label(excel_frame, text="Excelファイルをドラッグ＆ドロップしてください", font=("Arial", 12, "bold"), bg=PANEL_BG).pack(pady=5)
        self.excel_entry = tk.Entry(excel_frame, width=80, font=("Arial", 14, "bold"), bg="white",
                                    highlightthickness=2, highlightbackground="#004080", highlightcolor="#004080")
        self.excel_entry.pack(pady=5, ipady=4, padx=10)
        self.excel_entry.drop_target_register(DND_FILES)
        self.excel_entry.dnd_bind('<<Drop>>', self.on_drop_excel)

        # ステータス + プログレスバー
        status_frame = tk.Frame(self, bg=PANEL_BG, bd=2, relief="groove")
        status_frame.pack(pady=10, padx=20, fill="x")
        self.status_label = tk.Label(status_frame, text="準備中...", font=("Arial", 14, "bold"), fg="blue", bg=PANEL_BG)
        self.status_label.pack(pady=10)
        self.progress = ttk.Progressbar(status_frame, length=PROGRESS_LENGTH, mode="determinate")
        self.progress.pack(pady=5)
        self.progress["value"] = 0

        # Error Box
        error_frame = tk.Frame(self, bg=PANEL_BG, bd=2, relief="groove")
        error_frame.pack(pady=10, padx=20, fill="both", expand=True)
        tk.Label(error_frame, text="エラーメッセージ", font=("Arial", 12, "bold"), bg=PANEL_BG).pack(anchor="center", pady=5)
        self.error_box = tk.Text(error_frame, height=5, width=90, font=("Arial", 12), bg="white", fg="red")
        self.error_box.pack(padx=10, pady=5, fill="both", expand=True)
        self.error_box.config(state=tk.DISABLED)

        # Tag màu cho từng loại trạng thái
        self.error_box.tag_config("error", foreground="red")
        self.error_box.tag_config("success", foreground="green")
        self.error_box.tag_config("info", foreground="#0066cc")       # xanh dương
        self.error_box.tag_config("warning", foreground="#cc6600")

        # ボタン
        button_frame = tk.Frame(self, bg=BG_COLOR)
        button_frame.pack(pady=20)

        self.start_btn = tk.Button(button_frame, text="開始", command=self.process_manager.start_process,
                                   bg="#32cd32", fg="white", activebackground="#228b22",
                                   width=12, font=("Arial", 12, "bold"))
        self.start_btn.pack(side=tk.LEFT, padx=15)

        self.print_done_btn = tk.Button(button_frame, text="印刷完了", state=tk.DISABLED,
                                        command=self.process_manager.after_print,
                                        bg="#ffa500", fg="white", activebackground="#ff8c00",
                                        width=12, font=("Arial", 12, "bold"))
        self.print_done_btn.pack(side=tk.LEFT, padx=15)

        self.stop_btn = tk.Button(button_frame, text="非常停止", command=self.process_manager.emergency_stop,
                                  bg="#ff4500", fg="white", activebackground="#cc3700",
                                  width=12, font=("Arial", 12, "bold"))
        self.stop_btn.pack(side=tk.LEFT, padx=15)

        self.quit_btn = tk.Button(button_frame, text="終了", command=self.quit,
                                  bg="#dc143c", fg="white", activebackground="#a40000",
                                  width=12, font=("Arial", 12, "bold"))
        self.quit_btn.pack(side=tk.RIGHT, padx=15)

    # --- Sự kiện UI ---
    def on_drop_excel(self, event):
        file_path = event.data.strip("{}")
        if file_path.lower().endswith((".xlsx", ".xls")):
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, os.path.basename(file_path))
            self.excel_full_path = file_path
            blink_widget(self.excel_entry)
            self.print_done_btn.config(state=tk.DISABLED)
            self.status_label.config(text="Excelファイルを確認しました。開始ボタンを押してください。", fg="blue")
        else:
            messagebox.showerror("エラー", "Excelファイルを選択してください。")

    def browse_icd_folder(self):
        folder = filedialog.askdirectory(title="ICADフォルダを選択してください")
        if folder:
            self.icd_entry.delete(0, tk.END)
            self.icd_entry.insert(0, folder)
