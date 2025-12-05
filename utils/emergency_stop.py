

# emergency_stop.py
import threading
import os
import subprocess


class EmergencyStopManager:
    def __init__(self):
        self._stop_event = threading.Event()

    def trigger_stop(self):
        """非常停止を作動させます。"""
        self._stop_event.set()
        print("⚠ 非常停止 được kích hoạt!")

    def reset(self):
        """一時停止状態をリセットします（再起動が必要な場合）。"""
        self._stop_event.clear()

    def is_stop_requested(self) -> bool:
        """停止要求があるかどうかを確認します。"""
        return self._stop_event.is_set()


def cleanup_on_stop(app, info_dict):
    """
    非常停止時の後処理: Step 1で作成されたフォルダを削除
    
    Args:
        app: ShutsuzuuApp インスタンス
        info_dict: step1_create_and_copy の戻り値（output_folder を含む）
    """
    if not info_dict or "output_folder" not in info_dict:
        return
    
    folder_to_delete = info_dict["output_folder"]
    _try_delete_folder(app, folder_to_delete)


def _try_delete_folder(app, folder_path):
    """フォルダ削除を試行（force delete 使用）"""
    
    try:
        if not os.path.exists(folder_path):
            print(f"[削除対象] フォルダが存在しません: {folder_path}")
            return
        
        print(f"[削除開始] フォルダ: {folder_path}")
        
        # Force delete を試行（Windows cmd経由）
        _force_delete_with_cmd(app, folder_path)
        
    except Exception as e:
        print(f"[削除エラー] {str(e)}")


def _force_delete_with_cmd(app, folder_path):
    """Windows cmd で強制削除を試行"""
    from utils.UI_helpers import log_success, log_error
    
    try:
        print(f"[cmd実行] rmdir /s /q を実行中...")
        
        # cmd.exe で rmdir コマンドを実行
        cmd = f'cmd.exe /c "rmdir /s /q "{folder_path}""'
        result = subprocess.run(
            cmd,
            shell=True,
            capture_output=True,
            text=True,
            timeout=10
        )
        
        # 削除確認
        if not os.path.exists(folder_path):
            log_success(app, f"フォルダを削除しました: {folder_path}")
            print(f"[削除完了] ✓ {folder_path}")
        else:
            error_msg = f"フォルダを手動で削除してください: {folder_path}"
            log_error(app, error_msg)
            print(f"[削除失敗] {error_msg}")
                
    except subprocess.TimeoutExpired:
        error_msg = "cmd での削除がタイムアウトしました。フォルダを手動で削除してください。"
        log_error(app, error_msg)
        print(f"[タイムアウト] {error_msg}")
    except Exception as e:
        error_msg = f"フォルダを手動で削除してください: {str(e)}"
        log_error(app, error_msg)
        print(f"[エラー] {error_msg}")


# 共有用のグローバルインスタンスを作成する
emergency_manager = EmergencyStopManager()
