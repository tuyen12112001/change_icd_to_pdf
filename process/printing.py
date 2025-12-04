import os
import pyautogui
import pyperclip
import time
import sys
import pygetwindow as gw
from pynput.keyboard import Key, Controller
import mss
import cv2
import numpy as np
from utils.check_ICAD_and_Docuworks import ensure_docuworks_running, ensure_icad_running
from utils.emergency_stop import emergency_manager
from config.settings import IMAGE1_PATH, IMAGE2_PATH

keyboard = Controller()


# ===========================================================
# ✅ DocuWorksで新しいフォルダを作成
# ===========================================================
def create_docuworks_folder(excel_name_clean):
    """
    DocuWorksで新しいフォルダを作成し、名前を返す。
    """
    
    if not ensure_docuworks_running():
        raise Exception("DocuWorksが開いていません。")


    # フォルダ名はExcel名（LS-は既に除去済み）
    folder_name = excel_name_clean

    # Alt+F → N → Fで新規フォルダ作成
    pyautogui.keyDown('alt')
    pyautogui.press('f')
    pyautogui.keyUp('alt')
    time.sleep(0.3)
    pyautogui.press('n')
    time.sleep(0.3)
    pyautogui.press('f')
    time.sleep(0.5)

    # 名前入力
    pyautogui.typewrite(folder_name, interval=0.05)
    pyautogui.press('enter')
    time.sleep(1)

    if emergency_manager.is_stop_requested():
        print("⚠ フォルダ入力後に非常停止が押されました。処理を中断します。")
        return None


    print(f"✅ DocuWorksで新しいフォルダを作成しました: {folder_name}")
    return folder_name

# ===========================================================
# 🔍 画像認識で座標を取得（複数モニター対応）
# ===========================================================
def locate_center_mss(template_path, threshold=0.80):
    """
    画面全体をキャプチャし、テンプレート画像の位置を検索。
    見つかった場合は中心座標を返す。見つからない場合はNone。
    """
    try:
        with mss.mss() as sct:
            monitor = sct.monitors[0]  # 全画面キャプチャ
            screenshot = np.array(sct.grab(monitor))
            screenshot = cv2.cvtColor(screenshot, cv2.COLOR_BGRA2BGR)

        template = cv2.imread(template_path, cv2.IMREAD_COLOR)
        if template is None:
            print(f"⚠ 画像を読み込めません: {template_path}")

        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(result)

        if max_val < threshold:
            return None

        th, tw = template.shape[:2]
        center_x = max_loc[0] + tw // 2 + monitor['left']
        center_y = max_loc[1] + th // 2 + monitor['top']
        return (center_x, center_y)
    except Exception as e:
        print(f"⚠ locate_center_mssでエラー: {e}")
        return None

# ===========================================================
# 🔎 画像をクリック（見つかるまでリトライ）
# ===========================================================
def click_one_of_images(image_paths, max_attempts=10, confidence=0.80, wait_time=1):
    """
    Tìm và click một trong các ảnh trong danh sách image_paths.
    Nếu tìm thấy ảnh nào thì click và return True, nếu không thì return False.
    """
    for attempt in range(max_attempts):
        for image_path in image_paths:
            try:
                loc = locate_center_mss(image_path, threshold=confidence)
                if loc:
                    pyautogui.click(loc)
                    print(f"🖱 {os.path.basename(image_path)} をクリックしました (試行 {attempt+1})")
                    return True
            except Exception as e:
                print(f"⚠ 画像検索中にエラー: {e}")
        print(f"⏳ 見つかりません (試行 {attempt+1}/{max_attempts})...")
        time.sleep(wait_time)
    print("❌ どの画像も見つかりませんでした")
    return False


# ===========================================================
# ✅ ステップ2: 印刷処理
# ===========================================================
def step2_print_icd(output_dir, excel_name_clean):
    """
    ステップ2:
    - DocuWorksで新しいフォルダを作成。
    - ICADで印刷処理を実行。
    - 作成したフォルダ名を返す（ステップ3で削除するため）。
    """
    try:
        # 1. DocuWorksでフォルダ作成
        folder_name = create_docuworks_folder(excel_name_clean)
        if not folder_name:
            raise Exception("DocuWorksフォルダの作成に失敗しました。")

        # # ユーザーに印刷先フォルダを通知
        # messagebox.showinfo("情報", f"DocuWorksでフォルダ '{folder_name}' に印刷してください。")

        # 2. ICADを起動または確認
        icad_path = r"C:\MC2\bin\icad.exe"
        ensure_icad_running(icad_path)
        time.sleep(1)
    
        if emergency_manager.is_stop_requested():
            print("⚠ プリンタ選択後に非常停止が押されました。処理を中断します。")
            return None

        # 3. ICADで印刷メニューを開く (Alt+F → D)
        pyautogui.keyDown('alt')
        pyautogui.press('f')
        time.sleep(0.2)
        pyautogui.press('d')
        pyautogui.keyUp('alt')
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(1.0)

        if emergency_manager.is_stop_requested():
            print("⚠ プリンタ選択後に非常停止が押されました。処理を中断します。")
            return None
        
        # 4. 出力フォルダを貼り付け
        pyperclip.copy(output_dir)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(2.0)

        if emergency_manager.is_stop_requested():
            print("⚠ プリンタ選択後に非常停止が押されました。処理を中断します。")
            return None
        
        # 5. 印刷オプションを開く
        pyautogui.hotkey("alt", "o")
        time.sleep(0.5)

        if emergency_manager.is_stop_requested():
            print("⚠ プリンタ選択後に非常停止が押されました。処理を中断します。")
            return None
        
        # Alt + Down để mở danh sách
        pyautogui.keyDown('alt')
        pyautogui.press('o')
        time.sleep(0.5)
        keyboard.press(Key.alt)
        keyboard.press(Key.down)
        keyboard.release(Key.alt)
        time.sleep(0.5)

        click_one_of_images([IMAGE1_PATH, IMAGE2_PATH], max_attempts=10, confidence=0.80, wait_time=1)
        
        if emergency_manager.is_stop_requested():
            print("⚠ プリンタ選択後に非常停止が押されました。処理を中断します。")
            return None
        
        pyautogui.press("enter")
        time.sleep(1)

        # 7. 全ファイル選択 (Shift+End) → Enterで印刷開始
        keyboard.press(Key.home)
        keyboard.release(Key.home)
        time.sleep(0.2)
        keyboard.press(Key.shift)
        time.sleep(0.2)
        keyboard.press(Key.end)
        time.sleep(0.1)
        keyboard.release(Key.end)
        time.sleep(0.1)
        keyboard.release(Key.shift)
        time.sleep(0.2)
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)

        print("✅ 印刷処理が完了しました")
        if emergency_manager.is_stop_requested():
            print("⚠ プリンタ選択後に非常停止が押されました。処理を中断します。")
            return None
        return folder_name  # フォルダ名を返す
    
        
    except Exception as e:
        print(f"❌ 印刷中にエラーが発生しました: {e}")
        return None