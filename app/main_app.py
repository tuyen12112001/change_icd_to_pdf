
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
    blink_widget
)
from utils.check_ICAD_and_Docuworks import (
    activate_docuworks_and_open_user_folder,
)
from process.process_manager import ProcessManager

class ShutsuzuuApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        
        # ウィンドウサイズと位置を設定（右側に配置）
        window_width = 800
        window_height = 600
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x_position = screen_width - window_width - 50  # 右端から50pxの余白
        y_position = 300  # 上から300pxの位置
        
        self.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
        self.configure(bg=BG_COLOR)

        # Icon
        try:
            self.iconbitmap(ICON_PATH)
        except Exception:
            pass  # アイコンが壊れていても実行をブロックしない

        # ウィンドウを最前面に
        self.attributes('-topmost', True)
        self.lift()
        self.focus_force()

        # --- 以前の状態を初期化する ---
        self.info = None
        self.excel_full_path = ""
        self.folder_full_path = ""
        self.input_mode = "excel"  # "excel" のみ使用
        self.mode_var = tk.StringVar(value="excel")
        self.is_running = False

        # --- UI を作成する前にプロセス マネージャーを初期化する ---
        self.process_manager = ProcessManager(self)

        # UIを作成する
        self._build_ui()

    def _build_ui(self):
        # ヘッダー
        header = tk.Label(self, text=HEADER_TEXT, font=("Arial", 24, "bold"), fg="#004080", bg=BG_COLOR)
        header.pack(pady=15)

        # モード選択
        mode_frame = tk.Frame(self, bg=BG_COLOR)
        mode_frame.pack(pady=(0, 8))
        tk.Radiobutton(
            mode_frame,
            text="Excel",
            variable=self.mode_var,
            value="excel",
            command=self.on_mode_change,
            bg=BG_COLOR,
            font=("Arial", 11, "bold"),
        ).pack(side=tk.LEFT, padx=10)
        tk.Radiobutton(
            mode_frame,
            text="DocuWorks",
            variable=self.mode_var,
            value="docuworks",
            command=self.on_mode_change,
            bg=BG_COLOR,
            font=("Arial", 11, "bold"),
        ).pack(side=tk.LEFT, padx=10)

        # 入力フレーム用のコンテナ（モードで表示切替）
        self.input_container = tk.Frame(self, bg=BG_COLOR)
        self.input_container.pack(pady=10, padx=20, fill="x")

        # ExcelモードUI
        self.excel_frame = tk.Frame(self.input_container, bg=PANEL_BG, bd=2, relief="groove")
        tk.Label(self.excel_frame, text="Excelファイル選択", font=("Arial", 12, "bold"), bg=PANEL_BG).pack(pady=(6, 2))
        tk.Label(self.excel_frame, text="Excelファイルをドラッグ＆ドロップしてください", font=("Arial", 12), bg=PANEL_BG).pack(pady=(0, 4))

        self.excel_entry = tk.Entry(self.excel_frame, width=80, font=("Arial", 14, "bold"), bg="white",
                                    highlightthickness=2, highlightbackground="#004080", highlightcolor="#004080")
        self.excel_entry.pack(padx=10, pady=(0, 8), ipady=4)

        self.excel_frame.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
        self.excel_frame.dnd_bind('<<Drop>>', self.on_drop_excel)  # type: ignore[attr-defined]
        self.excel_entry.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
        self.excel_entry.dnd_bind('<<Drop>>', self.on_drop_excel)  # type: ignore[attr-defined]

        # DocuWorksモードUI
        self.test_frame = tk.Frame(self.input_container, bg=PANEL_BG, bd=2, relief="groove")
        tk.Label(self.test_frame, text="DocuWorks Desk", font=("Arial", 12, "bold"), bg=PANEL_BG).pack(pady=(12, 8))
        self.docuworks_btn = tk.Button(
            self.test_frame,
            text="Active DocuWorks",
            command=self.on_docuworks_test_click,
            bg="#1e90ff",
            fg="white",
            activebackground="#1c7ed6",
            width=24,
            font=("Arial", 12, "bold"),
        )
        self.docuworks_btn.pack(pady=(2, 14))

        # ステータス + プログレスバー
        status_frame = tk.Frame(self, bg=PANEL_BG, bd=2, relief="groove")
        status_frame.pack(pady=10, padx=20, fill="x")
        self.status_label = tk.Label(status_frame, text="準備中...", font=("Arial", 14, "bold"), fg="blue", bg=PANEL_BG)
        self.status_label.pack(pady=10)
        self.progress = ttk.Progressbar(status_frame, length=PROGRESS_LENGTH, mode="determinate")
        self.progress.pack(pady=5)
        self.progress["value"] = 0

        # エラーメッセージボックス
        error_frame = tk.Frame(self, bg=PANEL_BG, bd=2, relief="groove")
        error_frame.pack(pady=10, padx=20, fill="both", expand=True)
        tk.Label(error_frame, text="エラーメッセージ", font=("Arial", 12, "bold"), bg=PANEL_BG).pack(anchor="center", pady=5)
        self.error_box = tk.Text(error_frame, height=5, width=90, font=("Arial", 12), bg="white", fg="red")
        self.error_box.pack(padx=10, pady=5, fill="both", expand=True)
        self.error_box.config(state=tk.DISABLED)

        # 各ステータスタイプのカラータグ
        self.error_box.tag_config("error", foreground="red")
        self.error_box.tag_config("success", foreground="green")
        self.error_box.tag_config("info", foreground="#0066cc")       
        self.error_box.tag_config("warning", foreground="#cc6600")

        # ボタン
        self.button_frame = tk.Frame(self, bg=BG_COLOR)
        self.button_frame.pack(pady=20)

        self.start_btn = tk.Button(self.button_frame, text="開始", command=self.process_manager.start_process,
                                   bg="#32cd32", fg="white", activebackground="#228b22",
                                   width=12, font=("Arial", 12, "bold"))
        self.start_btn.pack(side=tk.LEFT, padx=15)

        self.print_done_btn = tk.Button(self.button_frame, text="印刷完了", state=tk.DISABLED,
                                        command=self.process_manager.after_print,
                                        bg="#ffa500", fg="white", activebackground="#ff8c00",
                                        width=12, font=("Arial", 12, "bold"))
        self.print_done_btn.pack(side=tk.LEFT, padx=15)

        # 交換完了ボタンは非表示（将来互換のためインスタンスのみ保持）
        self.exchange_done_btn = tk.Button(self.button_frame, text="交換完了", state=tk.DISABLED,
                                           command=self.on_exchange_btn_click,
                                           bg="#9370db", fg="white", activebackground="#6a5acd",
                                           width=12, font=("Arial", 12, "bold"))
        
        # Flag để track trạng thái button
        self.exchange_btn_mode = "first"  # "first" hoặc "retry"

        self.stop_btn = tk.Button(self.button_frame, text="非常停止", command=self.process_manager.emergency_stop,
                                  bg="#ff4500", fg="white", activebackground="#cc3700",
                                  width=12, font=("Arial", 12, "bold"))
        self.stop_btn.pack(side=tk.LEFT, padx=15)

        self.quit_btn = tk.Button(self.button_frame, text="終了", command=self.quit,
                                  bg="#dc143c", fg="white", activebackground="#a40000",
                                  width=12, font=("Arial", 12, "bold"))
        self.quit_btn.pack(side=tk.RIGHT, padx=15)

        # 初期モード反映
        self.on_mode_change()

    def on_mode_change(self):
        selected_mode = self.mode_var.get()
        self.input_mode = "excel" if selected_mode == "excel" else "docuworks"

        self.excel_frame.pack_forget()
        self.test_frame.pack_forget()

        if selected_mode == "excel":
            self.excel_frame.pack(fill="x")
            self.button_frame.pack(pady=20)
            self.status_label.config(text="Excelファイルを選択してください。", fg="blue")
        else:
            self.test_frame.pack(fill="x")
            self.button_frame.pack_forget()
            self.status_label.config(text="DocuWorksを起動してください。", fg="blue")

    # --- UI イベント ---
    def on_drop_excel(self, event):
        file_path = ""
        try:
            dropped = self.tk.splitlist(event.data)
            if dropped:
                file_path = dropped[0]
        except Exception:
            file_path = event.data.strip("{}")

        file_path = file_path.strip().strip('"')
        if file_path.lower().endswith((".xlsx", ".xls")):
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, os.path.basename(file_path))
            self.excel_full_path = file_path
            blink_widget(self.excel_entry)
            self.print_done_btn.config(state=tk.DISABLED)
            self.status_label.config(text="Excelファイルを確認しました。開始ボタンを押してください。", fg="blue")
        else:
            messagebox.showerror("エラー", "Excelファイルを選択してください。")

    def on_docuworks_test_click(self):
        activated = activate_docuworks_and_open_user_folder()
        if activated:
            self.status_label.config(text="DocuWorksでPDFフォルダを開きました。", fg="blue")
        else:
            messagebox.showerror("エラー", "DocuWorksでPDFフォルダを開けませんでした。")

    def on_exchange_btn_click(self):
        """Button動的処理：交換完了 or 再張り切り"""
        if self.exchange_btn_mode == "first":
            # 最初のクリック：交換完了処理
            self.process_manager.after_exchange()
        else:
            # 再度のクリック：再張り切り処理
            self.process_manager.retry_exchange()

