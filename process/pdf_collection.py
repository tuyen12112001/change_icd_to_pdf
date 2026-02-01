import os
import time
import pyautogui
import pyperclip
from utils.emergency_stop import emergency_manager
from utils.check_ICAD_and_Docuworks import ensure_docuworks_running


def compare_icd_pdf(output_folder, icd_list):
    """
    ICDファイル（Step 1でコピーした）とPDFファイル（Step 4で貼り付けた）を比較
    
    戻り値:
        missing: ICDあるがPDFにない
        extra: PDFあるがICDにない
    """
    # ICD ファイル名リスト（拡張子なし）
    icd_files = [os.path.splitext(os.path.basename(f))[0] for f in icd_list]

    # PDF ファイル名リスト（拡張子なし、出力フォルダから検索）
    pdf_files = [os.path.splitext(f)[0] for f in os.listdir(output_folder) 
                 if f.lower().endswith(".pdf")]

    # 比較
    missing = [icd for icd in icd_files if icd not in pdf_files]
    extra = [pdf for pdf in pdf_files if pdf not in icd_files]

    return missing, extra


def step3_collect_pdf(output_dir):
    """
    ステップ3 (PDFモード):
    - DocuWorks ウィンドウをアクティブにする
    - Ctrl+A: 全てを選択
    - Alt+T: メニューを開く
    - K: オプションを選択
    - 0: オプションを選択
    - 3: オプションを選択
    """
    try:
        print(f"🔍 DocuWorks ウィンドウをアクティブにしています...")
        
        # DocuWorks を起動または確認
        if not ensure_docuworks_running():
            print("❌ DocuWorks をアクティブ化できません。")
            return False
        
        time.sleep(1)
        print(f"📝 PDF コンバージョンコマンドを実行中...")
        
        # Ctrl+A を実行
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.5)
        
        # Alt を実行
        pyautogui.press("alt")
        time.sleep(0.2)
        
        # T を実行
        pyautogui.press("t")
        time.sleep(0.5)
        
        # K を実行
        pyautogui.press("k")
        time.sleep(0.3)
        
        # 0 を実行
        pyautogui.press("0")
        time.sleep(0.3)
        
        # 3 を実行
        pyautogui.press("3")
        time.sleep(1.0)
        
        print("✅ PDFコンバージョンコマンドを実行しました")
        return True

    except Exception as e:
        print(f"❌ PDF操作中にエラーが発生しました: {e}")
        return False


def step4_exchange_pdf(output_dir):
    """
    ステップ4 (PDFモード - 交換完了):
    - DocuWorks ウィンドウをアクティブにする
    - Ctrl+F: 検索ダイアログを開く
    - Enter: 確認
    - *.pdf を入力
    - Enter: 検索実行
    - Ctrl+A: 全てを選択
    - Ctrl+C: コピー
    - Ctrl+V: 出力フォルダに貼り付け
    """
    try:
        print(f"🔍 DocuWorks ウィンドウをアクティブにしています...")
        
        # DocuWorks を起動または確認
        if not ensure_docuworks_running():
            print("❌ DocuWorks をアクティブ化できません。")
            return False
        
        time.sleep(1)
        print(f"🔍 PDFファイルを検索中: *.pdf")
        
        # Ctrl+F を実行
        pyautogui.hotkey("ctrl", "f")
        time.sleep(0.5)
        
        # Enter を実行
        pyautogui.press("enter")
        time.sleep(0.3)
        
        # *.pdf を入力
        pyperclip.copy("*.pdf")
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.3)
        
        # Enter を実行（検索実行）
        pyautogui.press("enter")
        time.sleep(3)
        
        print("✅ PDFファイル検索完了。ファイルを選択中...")
        
        # Ctrl+A を実行（全てを選択）
        pyautogui.hotkey("ctrl", "a")
        time.sleep(0.5)
        
        # Ctrl+C を実行（コピー）
        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.5)
        
        print("✅ PDFファイルをコピーしました。元のフォルダに戻る中...")
        
        # Alt + Left 7回を実行（フォルダを7階層上に戻る）
        pyautogui.keyDown("alt")
        for i in range(7):
            pyautogui.press("left")
            time.sleep(0.2)
        pyautogui.keyUp("alt")
        time.sleep(1)
        
        print("✅ 元のフォルダに戻りました。出力フォルダを開く中...")
        
        # 出力フォルダを開く
        os.startfile(output_dir)
        time.sleep(2)
        
        # Ctrl+V を実行（貼り付け）
        pyautogui.hotkey("ctrl", "v")
        time.sleep(1.5)
        
        print("✅ PDFファイルを出力フォルダに貼り付けました")
        return True

    except Exception as e:
        print(f"❌ PDF検索・コピー中にエラーが発生しました: {e}")
        return False
