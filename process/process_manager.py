
# process/process_manager.py
import threading
import os
from tkinter import messagebox

from utils.UI_helpers import (
    animate_loading, stop_loading, update_status,
    log_error, clear_error_box,
    update_file_comparison_message, add_delete_xdw_buttons,
)
from config.settings import STATUS_ERROR_COLOR, STATUS_WARN_COLOR

from .create import step1_create_and_copy
from .printing import step2_print_icd
from .xdw_collection import step3_collect_xdw
from .clear import step4_cleanup
from utils.cleanup_xdw import cleanup_xdw_on_user_request, show_no_delete_xdw_message
from utils.excel_collect import add_ls_lk_excel_set_to_output
from utils.excel_remove import excel_remove
from utils.emergency_stop import emergency_manager, cleanup_on_stop


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

        excel_path = os.path.normpath(self.app.excel_full_path.strip()) if getattr(self.app, "excel_full_path", "") else ""
        if not excel_path or not os.path.isfile(excel_path):
            messagebox.showerror("エラー", "有効なExcelファイルを選択してください。")
            return

        # Vô hiệu hóa nút 開始 để tránh nhấn nhiều lần
        self.app.start_btn.config(state="disabled")

        # Bật hiệu ứng loading
        self.app.is_running = True
        animate_loading(self.app, "ステップ1: ファイルコピー中")
        self.app.progress["value"] = 0

        threading.Thread(target=self._run_steps, args=(excel_path,), daemon=True).start()

    def _run_steps(self, excel_path):
        try:
            # Step 1: Copy ICD files
            update_status(self.app, "ステップ1: ファイルコピー中...", 25)
            self.app.info = step1_create_and_copy(excel_path)

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
            combine_msg = msg_icd + "\n" + msg_cleanup
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

        update_status(self.app, "ステップ3: .xdwファイルを取得中...", 85)
        moved_count = step3_collect_xdw(self.app.info["output_folder"], self.app.info["docuworks_folder"])

        if moved_count > 0:
            update_status(self.app, f".xdwファイル {moved_count} 件を移動しました。", 95)
            update_status(self.app, "ステップ4: クリーンアップ中...", 98)
            try:
                step4_cleanup(self.app.info["output_folder"])
                update_status(self.app, "完了！すべての処理が終了しました。", 100, color="green")

                # Mở thư mục kết quả (Windows)
                try:
                    os.startfile(self.app.info["output_folder"])
                except Exception:
                    pass

                # Kiểm tra số lượng
                if moved_count != self.app.info['copied_count']:
                    warning_msg = (
                        f"処理が完了しましたが、ファイル数が一致しません。\n"
                        f"ICDファイル数: {self.app.info['copied_count']} 件\n"
                        f".xdwファイル数: {moved_count} 件\n"
                        "DocuWorksのファイルを確認してください。"
                    )
                    messagebox.showwarning("注意", warning_msg)
                    update_file_comparison_message(self.app, warning_msg, status="warning")
                    
                    # Thêm 2 button Yes/No để user chọn xóa file XDW
                    def on_yes():
                        cleanup_xdw_on_user_request(self.app, self.app.info["output_folder"])
                    
                    def on_no():
                        show_no_delete_xdw_message(self.app)
                    
                    add_delete_xdw_buttons(self.app, on_yes, on_no)
                else:
                    success_msg = (
                        f"処理が完了しました。\n移動した.xdwファイル数: {moved_count} 件\n"
                        "印刷処理が正常に完了しました。"
                    )
                    messagebox.showinfo("情報", success_msg)
                    update_file_comparison_message(self.app, success_msg, status="info")

            except Exception as e:
                log_error(self.app, f"クリーンアップに失敗しました: {str(e)}")
        else:
            log_error(self.app, ".xdwファイルの取得に失敗しました。")

