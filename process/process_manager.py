
# process/process_manager.py
import threading
import os
from tkinter import messagebox

from utils.UI_helpers import (
    animate_loading, stop_loading, update_status,
    log_error, clear_error_box,
    update_file_comparison_message, add_delete_xdw_buttons, add_delete_pdf_buttons,
)
from config.settings import STATUS_ERROR_COLOR, STATUS_WARN_COLOR

from .create import step1_create_and_copy
from .printing import step2_print_icd
from .xdw_collection import step3_collect_xdw
from .pdf_collection import step3_collect_pdf, step4_exchange_pdf, compare_icd_pdf, retry_exchange_pdf
from .rename_pdf import remove_suffix_3d_from_pdf
from .clear import step4_cleanup
from .clear_pdf import step4_cleanup_pdf
from utils.cleanup_xdw import cleanup_xdw_on_user_request, show_no_delete_xdw_message
from utils.cleanup_pdf import cleanup_pdf_on_user_request, show_no_delete_pdf_message
from utils.excel_collect import add_ls_lk_excel_set_to_output
from utils.excel_remove import excel_remove
from utils.emergency_stop import emergency_manager, cleanup_on_stop
from utils.file_compare import compare_icd_xdw


class ProcessManager:
    def __init__(self, app):
        self.app = app

    # イベント: 非常に停止
    def emergency_stop(self):
        emergency_manager.trigger_stop()
        self.app.status_label.config(text="処理を強制停止しました。", fg=STATUS_ERROR_COLOR)
        self.app.progress["value"] = 0
        log_error(self.app, "ユーザーによって非常停止が実行されました。")
        self.app.print_done_btn.config(state="disabled")
        
        # 作成されたフォルダをクリーンアップ
        cleanup_on_stop(self.app, self.app.info)
        
        messagebox.showwarning("警告", "処理を強制停止しました。")

    # イベント: 開始
    def start_process(self):
        emergency_manager.reset()
        self.app.print_done_btn.config(state="disabled")
        clear_error_box(self.app)  # 古いエラーメッセージをすべて削除する

        # モード判定（ExcelまたはFolder）
        input_mode = self.app.mode_var.get() if hasattr(self.app, "mode_var") else "excel"
        
        if input_mode == "excel":
            excel_path = os.path.normpath(self.app.excel_full_path.strip()) if getattr(self.app, "excel_full_path", "") else ""
            if not excel_path or not os.path.isfile(excel_path):
                messagebox.showerror("エラー", "有効なExcelファイルを選択してください。")
                return
            folder_path = None
        else:  # folder mode
            folder_path = os.path.normpath(self.app.folder_full_path.strip()) if getattr(self.app, "folder_full_path", "") else ""
            if not folder_path or not os.path.isdir(folder_path):
                messagebox.showerror("エラー", "有効なICDフォルダを選択してください。")
                return
            excel_path = None

        # Vô hiệu hóa nút 開始 để tránh nhấn nhiều lần
        self.app.start_btn.config(state="disabled")

        # Bật hiệu ứng loading
        self.app.is_running = True
        animate_loading(self.app, "ステップ1: ファイルコピー中")
        self.app.progress["value"] = 0

        threading.Thread(target=self._run_steps, args=(excel_path, folder_path), daemon=True).start()

    
    def _open_folder_safe(self, path):
        """
        Mở thư mục kết quả một cách an toàn, tương thích Windows, macOS, Linux.
        """
        import sys
        import subprocess
        try:
            if sys.platform.startswith("win"):
                os.startfile(path)
        except Exception:
            pass

    def _run_steps(self, excel_path, folder_path=None):
        try:
            # Step 1: Copy ICD files
            update_status(self.app, "ステップ1: ファイルコピー中...", 25)
            self.app.info = step1_create_and_copy(excel_path=excel_path, icd_folder_path=folder_path)

            # Kiểm tra lỗi ngay sau khi Step 1
            if not self.app.info or "error" in self.app.info:
                stop_loading(self.app)
                log_error(self.app, self.app.info.get("error", "不明なエラーが発生しました。"))
                return

            # Dừng loading và hiển thị trạng thái hoàn tất Step 1
            stop_loading(self.app)
            update_status(self.app, f"コピー完了: {self.app.info['copied_count']} ファイル", 50)

            # Hiển thị thông báo ICD
            if self.app.info.get("not_found"):
                msg_icd = f"{len(self.app.info['not_found'])} 件のICDファイルが見つかりません:\n" + "\n".join(self.app.info['not_found'][:10])
                status = "error"
            else:
                msg_icd = "すべてのICDファイルが見つかりました！"
                status = "success"

            # Delay 1.5 giây trước khi chuyển sang Step 2
            self.app.after(1500, lambda: self._continue_steps(excel_path, msg_icd, status))

        except Exception as e:
            log_error(self.app, str(e))
        finally:
            # Cho phép nhấn lại 開始 nếu cần
            self.app.start_btn.config(state="normal")

    def _continue_steps(self, excel_path, msg_icd, status):
        try:
            # Chỉ xử lý Excel khi là mode Excel
            msg_cleanup = ""
            msg_hold_add = ""
            
            if excel_path:  # Excel mode
                # Excel xử lý
                related = add_ls_lk_excel_set_to_output(
                    excel_path=excel_path,
                    output_dir=self.app.info["output_folder"],
                    include_selected=True,
                    recursive=False
                )

                folder_id = self.app.info["excel_name_clean"]
                base_dir = self.app.info["output_folder"]
                kept, removed, skipped = excel_remove(folder_id, base_dir)

                # Hiển thị thông báo tổng hợp
                msg_cleanup = f"Excel整理完了: 保持={kept}, 削除={removed}, スキップ={skipped}"

                # 保留と追加に関するお知らせを追加
                if self.app.info.get("added_due_to_addition") and self.app.info.get("skipped_due_to_hold"):
                    msg_hold_add += f"\nℹ️追加が指定されたため、コピーした部品番号: {', '.join(self.app.info['added_due_to_addition'][:10])}"
                    if len(self.app.info['added_due_to_addition']) > 10:
                        msg_hold_add += " ... (残り省略)"
                else:
                    if self.app.info.get("skipped_due_to_hold"):
                        msg_hold_add += f"\nℹ️保留のためコピーしなかった部品番号: {', '.join(self.app.info['skipped_due_to_hold'][:10])}"
                        if len(self.app.info['skipped_due_to_hold']) > 10:
                            msg_hold_add += " ... (残り省略)"

                    if self.app.info.get("added_due_to_addition"):
                        msg_hold_add += f"\nℹ️追加が指定されたため、コピーした部品番号: {', '.join(self.app.info['added_due_to_addition'][:10])}"
                        if len(self.app.info['added_due_to_addition']) > 10:
                            msg_hold_add += " ... (残り省略)"
            else:  # Folder mode
                msg_cleanup = "フォルダモード: すべてのICDファイルをコピーしました。"

            combine_msg = msg_icd + "\n" + msg_cleanup + msg_hold_add
            combine_msg += f"\nこれから {self.app.info['copied_count']} 部品図を印刷します"

            # Show vào error_box theo trạng thái
            from utils.UI_helpers import update_error_box
            update_error_box(self.app, combine_msg, status=status)

            update_status(self.app, "ステップ2: 印刷中...", 75)

            # Step 2: 印刷処理 — chạy thread phụ để không block UI
            threading.Thread(target=self._print_icd, daemon=True).start()

        except Exception as e:
            log_error(self.app, str(e))

    def _print_icd(self):
        try:
            folder_name = step2_print_icd(self.app.info["output_folder"], self.app.info["excel_name_clean"])

            if emergency_manager.is_stop_requested():
                print("⚠ 非常停止により処理を中断しました。")
                return

            if not folder_name:
                raise Exception("印刷に失敗しました。DocuWorksが開いていない可能性があります。")

            self.app.info["docuworks_folder"] = folder_name
            self.app.after(0, lambda: self.app.print_done_btn.config(state="normal"))
            self.app.after(0, lambda: update_status(self.app, "印刷完了を確認してください。", 75, color=STATUS_WARN_COLOR))

        except Exception as e:
            log_error(self.app, str(e))

    # Sự kiện: 印刷完了
    def after_print(self):
        # Kiểm tra nếu đã nhấn 非常停止
        if emergency_manager.is_stop_requested():
            print("⚠ 非常停止が押されたため、印刷完了処理を中断します。")
            self.app.status_label.config(text="処理は非常停止で中断されました。", fg=STATUS_ERROR_COLOR)
            return

        # PDF mode か Excel mode かで処理を分ける
        if self.app.input_mode == "folder":
            # PDFモード
            self._after_print_pdf_mode()
        else:
            # Excelモード
            self._after_print_excel_mode()

    def _after_print_excel_mode(self):
        """Excelモード: PDFファイルを収集・処理"""
        # Step 3: PDF変換（Ctrl+A, Alt+T, K, 0, 3を実行）
        update_status(self.app, "ステップ3: PDFコンバージョン中...", 85)
        
        success = step3_collect_pdf(self.app.info["output_folder"])
        
        if success:
            update_status(self.app, "PDFコンバージョン完了。交換完了ボタンを押してください。", 90, color=STATUS_WARN_COLOR)
            self.app.exchange_done_btn.config(state="normal")
        else:
            log_error(self.app, "PDFコンバージョンに失敗しました。")

    def _after_print_pdf_mode(self):
        """PDFモード: Ctrl+A, Alt+T, K, 0, 3 を実行して PDF化"""
        update_status(self.app, "ステップ3: PDFコンバージョン中...", 85)
        
        success = step3_collect_pdf(self.app.info["output_folder"])
        
        if success:
            update_status(self.app, "PDFコンバージョン完了。交換完了ボタンを押してください。", 90, color=STATUS_WARN_COLOR)
            self.app.exchange_done_btn.config(state="normal")
            # messagebox.showinfo("情報", "PDFコンバージョンコマンドを実行しました。\n交換完了ボタンを押してください。")
        else:
            log_error(self.app, "PDFコンバージョンに失敗しました。")

    # Sự kiện: 交換完了（PDFモード）
    def after_exchange(self):
        """PDF処理: PDF検索・コピー・クリーンアップ（Folder/Excelモード共通）"""
        if emergency_manager.is_stop_requested():
            print("⚠ 非常停止が押されたため、交換完了処理を中断します。")
            self.app.status_label.config(text="処理は非常停止で中断されました。", fg=STATUS_ERROR_COLOR)
            return

        update_status(self.app, "ステップ4: PDF検索・移動中...", 95)
        
        success = step4_exchange_pdf(self.app.info["output_folder"])
        
        if success:
            # PDF ファイル以外を削除（クリーンアップ）- ICD ファイルを削除
            update_status(self.app, "ステップ5: クリーンアップ中...", 98)
            step4_cleanup_pdf(self.app.info["output_folder"])
            
            # ICD と PDF を比較（-3D削除前に比較する必要がある）
            update_status(self.app, "ステップ6: ファイル比較中...", 99)
            missing, extra = compare_icd_pdf(
                self.app.info["output_folder"],
                self.app.info.get("icd_list", [])
            )
            
            # PDFの名前から "-3D" を削除
            update_status(self.app, "ステップ7: PDF名前変更中 (-3D を削除)...", 99)
            rename_log = remove_suffix_3d_from_pdf(self.app.info["output_folder"])
            if rename_log:
                print(f"✅ {len(rename_log)} 個のPDFファイルを名前変更しました")
            
            # 結果を確認
            if missing or extra:
                warning_msg = (
                    f"処理が完了しましたが、ファイル数が一致しません。\n"
                    f"ICDファイル数: {len(self.app.info.get('icd_list', []))} 件\n"
                    f"PDFファイル数: {len([f for f in os.listdir(self.app.info['output_folder']) if f.lower().endswith('.pdf')])} 件\n"
                )
                if missing:
                    warning_msg += f"\n不足ファイル({len(missing)}):\n" + "\n".join(missing[:10])
                    if len(missing) > 10:
                        warning_msg += "\n... (残り省略)"
                if extra:
                    warning_msg += f"\n余分ファイル({len(extra)}):\n" + "\n".join(extra[:10])
                    if len(extra) > 10:
                        warning_msg += "\n... (残り省略)"

                messagebox.showwarning("注意", warning_msg)
                update_file_comparison_message(self.app, warning_msg, status="warning")
                # ファイル数が一致しない場合、再張り切りボタンを表示（button text を変更）
                self.app.exchange_btn_mode = "retry"
                self.app.after(0, lambda: self.app.exchange_done_btn.config(text="再張り切り", state="normal"))
                
                # PDF削除ボタンをerror_boxに表示
                add_delete_pdf_buttons(
                    self.app,
                    on_yes_callback=lambda: cleanup_pdf_on_user_request(self.app, self.app.info["output_folder"]),
                    on_no_callback=lambda: show_no_delete_pdf_message(self.app)
                )
            else:
                success_msg = (
                    f"処理が完了しました。\n"
                    f"ICDファイル数: {len(self.app.info.get('icd_list', []))} 件\n"
                    f"PDFファイル数: {len([f for f in os.listdir(self.app.info['output_folder']) if f.lower().endswith('.pdf')])} 件\n"
                    f"ファイル数が一致しました！"
                )
                messagebox.showinfo("情報", success_msg)
                update_file_comparison_message(self.app, success_msg, status="success")
            
            update_status(self.app, "完了！すべての処理が終了しました。", 100, color="green")
            self.app.exchange_done_btn.config(state="disabled")
            self._open_folder_safe(self.app.info["output_folder"])
        else:
            log_error(self.app, "PDF検索に失敗しました。")

    # イベント: 再張り切り
    def retry_exchange(self):
        """再張り切り: DocuWorksからのみ貼り付けを実行"""
        if emergency_manager.is_stop_requested():
            print("⚠ 非常停止が押されたため、再張り切り処理を中断します。")
            self.app.status_label.config(text="処理は非常停止で中断されました。", fg=STATUS_ERROR_COLOR)
            return

        if not self.app.info or not self.app.info.get("output_folder"):
            messagebox.showerror("エラー", "出力フォルダが見つかりません。")
            return

        update_status(self.app, "再張り切り処理中...", 95)
        
        success = retry_exchange_pdf(self.app.info["output_folder"])
        
        if success:
            # PDF ファイル以外を削除（クリーンアップ）- ICD ファイルを削除
            update_status(self.app, "クリーンアップ中...", 98)
            step4_cleanup_pdf(self.app.info["output_folder"])
            
            # ICD と PDF を比較
            update_status(self.app, "ファイル比較中...", 99)
            missing, extra = compare_icd_pdf(
                self.app.info["output_folder"],
                self.app.info.get("icd_list", [])
            )
            
            # PDFの名前から "-3D" を削除
            update_status(self.app, "PDF名前変更中 (-3D を削除)...", 99)
            rename_log = remove_suffix_3d_from_pdf(self.app.info["output_folder"])
            if rename_log:
                print(f"✅ {len(rename_log)} 個のPDFファイルを名前変更しました")
            
            # 結果を確認
            if missing or extra:
                warning_msg = (
                    f"処理が完了しましたが、ファイル数が一致しません。\n"
                    f"ICDファイル数: {len(self.app.info.get('icd_list', []))} 件\n"
                    f"PDFファイル数: {len([f for f in os.listdir(self.app.info['output_folder']) if f.lower().endswith('.pdf')])} 件\n"
                )
                if missing:
                    warning_msg += f"\n不足ファイル({len(missing)}):\n" + "\n".join(missing[:10])
                    if len(missing) > 10:
                        warning_msg += "\n... (残り省略)"
                if extra:
                    warning_msg += f"\n余分ファイル({len(extra)}):\n" + "\n".join(extra[:10])
                    if len(extra) > 10:
                        warning_msg += "\n... (残り省略)"

                messagebox.showwarning("注意", warning_msg)
                update_file_comparison_message(self.app, warning_msg, status="warning")
                # 再度ファイル数が一致しない場合、ボタンは再張り切り状態のまま保つ（何度でも試せる）
                self.app.exchange_btn_mode = "retry"
                self.app.after(0, lambda: self.app.exchange_done_btn.config(text="再張り切り", state="normal"))
            else:
                success_msg = (
                    f"処理が完了しました。\n"
                    f"ICDファイル数: {len(self.app.info.get('icd_list', []))} 件\n"
                    f"PDFファイル数: {len([f for f in os.listdir(self.app.info['output_folder']) if f.lower().endswith('.pdf')])} 件\n"
                    f"ファイル数が一致しました！"
                )
                messagebox.showinfo("情報", success_msg)
                update_file_comparison_message(self.app, success_msg, status="success")
            
            update_status(self.app, "完了！再張り切り処理が終了しました。", 100, color="green")
            # 成功した場合のみ、ボタンを交換完了に戻す
            self.app.exchange_btn_mode = "first"
            self.app.after(0, lambda: self.app.exchange_done_btn.config(text="交換完了", state="disabled"))
            self._open_folder_safe(self.app.info["output_folder"])
        else:
            log_error(self.app, "再張り切り処理に失敗しました。")

