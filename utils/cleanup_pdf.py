# utils/cleanup_pdf.py
import os
import time
import pyautogui
from process.clear import force_delete
from utils.refresh_explore import refresh_explorer


def delete_all_pdf_files(output_folder):
    """
    削除 output_folder 内のすべての .pdf ファイル (強制削除)
    
    Args:
        output_folder: PDFファイルが格納されているフォルダパス
    
    Returns:
        (success: bool, deleted_count: int, error_msg: str)
    """
    try:
        if not os.path.exists(output_folder):
            return False, 0, f"フォルダが存在しません: {output_folder}"
        
        pdf_files = [f for f in os.listdir(output_folder) if f.lower().endswith(".pdf")]
        deleted_count = 0
        errors = []
        
        for pdf_file in pdf_files:
            file_path = os.path.join(output_folder, pdf_file)
            print(f"[削除開始] {pdf_file}")
            if force_delete(file_path):
                deleted_count += 1
                print(f"[削除完了] {pdf_file}")
            else:
                errors.append(f"{pdf_file} (強制削除失敗)")
        
        # ✅ Refresh Explorer after delete
        try:
            refresh_explorer(output_folder)
            print("🔄 Explorer refreshed after deleting PDF.")
        except Exception as e:
            print(f"⚠ Error refreshing Explorer: {e}")
            pass

        if errors:
            error_msg = f"{deleted_count} 件削除しましたが、{len(errors)} 件削除できません:\n" + "\n".join(errors)
            print(f"[警告] {error_msg}")
            return False, deleted_count, error_msg
        
        success_msg = f"{deleted_count} 個のPDFファイルを強制削除しました。"
        print(f"[成功] {success_msg}")
        return True, deleted_count, success_msg
        
    except Exception as e:
        error_msg = f"PDFファイルの強制削除に失敗しました: {str(e)}"
        print(f"[エラー] {error_msg}")
        return False, 0, error_msg


def cleanup_pdf_on_user_request(app, output_folder):
    """
    ユーザーがPDFファイル削除を要求したときの処理
    削除 + メッセージ表示を一括処理
    
    Args:
        app: ShutsuzuuApp instance
        output_folder: PDFファイルが格納されているフォルダパス
    """
    from utils.UI_helpers import log_success, log_warning, log_error
    
    try:
        success, deleted_count, message = delete_all_pdf_files(output_folder)
        
        if success:
            # 成功メッセージ
            result_msg = (
                f" {deleted_count} 個のPDFファイルを削除しました。\n"
                "DocuWorksのファイルを確認してからお試しください。"
            )
            log_success(app, result_msg)
        else:
            # 失敗メッセージ
            result_msg = (
                f" PDFファイル削除に失敗しました:\n{message}\n"
                "ファイルを手動で削除してください。"
            )
            log_warning(app, result_msg)
                
        
        # ✅ Refresh Explorer after delete
        try:
            refresh_explorer(output_folder)
            print("🔄 Explorer refreshed after deleting PDF.")
        except Exception as e:
            print(f"⚠ Error refreshing Explorer: {e}")
            pass
    
    except Exception as e:
        log_error(app, f"PDFファイル削除処理に失敗: {str(e)}")


def show_no_delete_pdf_message(app):
    """
    ユーザーが「削除しない」を選択した場合のメッセージ
    
    Args:
        app: ShutsuzuuApp instance
    """
    from utils.UI_helpers import log_warning
    
    no_delete_msg = (
        "⚠️ PDFファイルを削除しません。\n"
        "DocuWorksのファイルを確認して、手動で削除するか再度お試しください。"
    )
    log_warning(app, no_delete_msg)
