# utils/cleanup_pdf.py
import os
import time
import pyautogui
from pathlib import Path
from process.clear import force_delete
from utils.refresh_explore import refresh_explorer


def get_docuworks_pdf_folder():
    r"""
    Windows API SHGetFolderPathW を使用してマイドキュメント フォルダパスを取得
    CSIDL_PERSONAL (5) = My Documents (ユーザーが設定した実際のフォルダ)
    """
    try:
        import ctypes
        from ctypes import wintypes

        buf = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        csidl_personal = 5  # CSIDL_PERSONAL = My Documents
        res = ctypes.windll.shell32.SHGetFolderPathW(None, csidl_personal, None, 0, buf)
        
        if res == 0 and os.path.isdir(buf.value):
            return buf.value
        
        # Fallback
        fallback = os.path.join(os.path.expanduser("~"), "Documents")
        if os.path.isdir(fallback):
            return fallback
        
        print(f"[警告] マイドキュメント フォルダが見つかりません")
        return None
    except Exception as e:
        print(f"[エラー] マイドキュメント フォルダ取得失敗: {str(e)}")
        return None


def delete_all_pdf_in_my_documents():
    """
    マイドキュメント フォルダ内のすべての .pdf ファイルを削除
    - 最大60秒間、My Documentsをリアルタイム監視
    - ファイルサイズが安定（1秒間変化なし）してから削除
    - ファイルが全部なくなるまでループを続ける
    
    Returns:
        (success: bool, deleted_count: int, message: str)
    """
    try:
        pdf_folder = get_docuworks_pdf_folder()
        if not pdf_folder or not os.path.isdir(pdf_folder):
            error_msg = "マイドキュメント フォルダが見つかりません。"
            print(f"[エラー] {error_msg}")
            return False, 0, error_msg
        
        print(f"\n📂 マイドキュメント内のPDFファイル削除開始: {pdf_folder}")
        print(f"⏱️  最大60秒間、ファイルをリアルタイム監視します...")
        
        deleted_count = 0
        failed_files = []
        start_time = time.time()
        timeout_sec = 60
        
        while time.time() - start_time < timeout_sec:
            # My Documents内のPDFファイルをリスト化
            pdf_files = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]
            
            if not pdf_files:
                print(f"\n✅ マイドキュメント内にPDFファイルがなくなりました（削除完了）")
                break
            
            print(f"\n📋 見つかったPDFファイル: {len(pdf_files)} 件")
            
            for pdf_file in pdf_files:
                file_path = os.path.join(pdf_folder, pdf_file)
                
                # ファイルが存在するか確認
                if not os.path.isfile(file_path):
                    continue
                
                try:
                    # ファイルサイズの安定性をチェック（1秒待機後に確認）
                    size_1 = os.path.getsize(file_path)
                    print(f"   [確認中] {pdf_file} (サイズ: {size_1} bytes)")
                    time.sleep(1)
                    
                    # ファイルがまだ存在するか確認
                    if not os.path.isfile(file_path):
                        print(f"   ℹ️ ファイルが既に削除されました: {pdf_file}")
                        continue
                    
                    size_2 = os.path.getsize(file_path)
                    
                    # サイズが変わらなければ削除しても安全
                    if size_1 == size_2:
                        print(f"   ✅ サイズ安定確認。削除中: {pdf_file}")
                        os.remove(file_path)
                        deleted_count += 1
                        print(f"   ✅ 削除成功: {pdf_file}")
                    else:
                        print(f"   ⏳ ファイル作成中（サイズ変化: {size_1} → {size_2}）。再チェック待機...")
                
                except PermissionError:
                    # ファイルロック中 → 強制削除を試す
                    print(f"   ⚠️ ファイルがロック中: {pdf_file}。強制削除を試みます...")
                    if force_delete(file_path):
                        deleted_count += 1
                        print(f"   ✅ 強制削除成功: {pdf_file}")
                    else:
                        failed_files.append(pdf_file)
                        print(f"   ❌ 強制削除失敗: {pdf_file}")
                
                except FileNotFoundError:
                    print(f"   ℹ️ ファイルが見つかりません: {pdf_file}")
                
                except Exception as e:
                    failed_files.append(pdf_file)
                    print(f"   ❌ 削除失敗: {pdf_file} ({str(e)})")
            
            # 次の監視周期まで少し待機
            time.sleep(1)
        
        # タイムアウト後の最終確認
        remaining_pdfs = [f for f in os.listdir(pdf_folder) if f.lower().endswith(".pdf")]
        
        # Refresh Explorer
        try:
            refresh_explorer(pdf_folder)
            print("🔄 Explorer更新完了")
        except Exception:
            pass
        
        # 結果メッセージ
        if remaining_pdfs:
            msg = f"マイドキュメントから {deleted_count} 個のPDFを削除しました。\n（{len(remaining_pdfs)} 個削除失敗・残存）"
            print(f"[警告] {msg}")
            return False, deleted_count, msg
        elif failed_files:
            msg = f"マイドキュメントから {deleted_count} 個のPDFを削除しました。\n（{len(failed_files)} 個削除失敗）"
            print(f"[警告] {msg}")
            return False, deleted_count, msg
        else:
            msg = f"マイドキュメントから {deleted_count} 個のPDFファイルをすべて削除しました。"
            print(f"[成功] {msg}")
            return True, deleted_count, msg
    
    except Exception as e:
        error_msg = f"マイドキュメント内のPDF削除処理エラー: {str(e)}"
        print(f"[エラー] {error_msg}")
        import traceback
        traceback.print_exc()
        return False, 0, error_msg


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
    
    # Enable 印刷完了ボタン以便ユーザーが再度試せるように
    app.print_done_btn.config(state="normal")


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
    
    # Enable 印刷完了ボタン以便ユーザーが再度試せるように
    app.print_done_btn.config(state="normal")