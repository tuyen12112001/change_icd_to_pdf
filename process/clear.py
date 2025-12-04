
import os
import shutil
import pyautogui
import time
import subprocess

def force_delete(file_path):
    """
    Ép xóa file bằng lệnh hệ thống (Windows).
    """
    try:
        subprocess.run(["cmd", "/c", "del", "/f", "/q", file_path], shell=True)
        print(f"✅ 強制削除: {file_path}")
        return True
    except Exception as e:
        print(f"⚠ 強制削除失敗: {file_path} | エラー: {e}")
        return False


def step4_cleanup(folder_path):
    try:
        if not os.path.exists(folder_path):
            print(f"❌ フォルダが存在しません: {folder_path}")
            return False

        # ★ 残す拡張子（小文字で統一）
        keep_exts = {'.xdw', '.xls', '.xlsx', '.xlsm'}
        deleted_files = []

        # 1. Xóa file không thuộc keep_exts
        for root, dirs, files in os.walk(folder_path):
            if "xdw file" in root:
                continue

            for file in files:
                ext = os.path.splitext(file)[1].lower()
                if ext not in keep_exts:
                    file_path = os.path.join(root, file)
                    if force_delete(file_path):
                        deleted_files.append(file)

        # 2. Xóa thư mục rỗng (trừ "xdw file")
        for root, dirs, files in os.walk(folder_path):
            for d in dirs:
                if d != "xdw file":
                    dir_path = os.path.join(root, d)
                    if not os.listdir(dir_path):
                        try:
                            shutil.rmtree(dir_path)
                        except Exception as e:
                            print(f"⚠ 空フォルダ削除失敗: {dir_path} | エラー: {e}")

        # 3. Refresh Explorer
        try:
            pyautogui.hotkey("f5")
            time.sleep(1)
        except Exception:
            pass

        # 4. Kết quả
        if deleted_files:
            print("✅ 削除したファイル一覧:")
            for f in deleted_files:
                print(f"   - {f}")
        else:
            print("ℹ 削除対象ファイルはありません。")

        return True

    except Exception as e:
        print(f"❌ クリーンアップ時のエラー: {e}")
        return False

