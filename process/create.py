from tkinter import filedialog, messagebox
import os
import shutil
import pandas as pd
import re
from utils.searchTools import search_number, search_gradually

def step1_create_and_copy(excel_path=None):
    """
    ステップ1 (Excelモード):
    - Excelファイルから列Kを読み込みICDファイルをコピー
    - ユーザーに出力フォルダを選択させる。
    - ICDフォルダから一致する.icdファイルをコピー。
    - config.txtを保存。
    戻り値:
        成功時: {output_folder, excel_name_clean, copied_count}
        エラー時: {"error": "エラーメッセージ"}
    """
    try:
        if not excel_path:
            return {"error": "Excelファイルを選択してください。"}
        return _step1_excel_mode(excel_path)

    except Exception as e:
        return {"error": f"エラーが発生しました: {str(e)}"}


def _step1_excel_mode(excel_path):
    """
    Excelモード処理
    """
    try:
        # Excelファイル名処理
        excel_name = os.path.splitext(os.path.basename(excel_path))[0]
        if not excel_name.startswith("LS-"):
            return {"error": "このファイルは製作部品表ではありません。ファイル名は 'LS-' で始まる必要があります。"}
        
        # ユーザーにフォルダ選択ダイアログを表示
        output_folder = filedialog.askdirectory(title="出力フォルダを選択してください")
        if not output_folder:
            return {"error": "出力フォルダが選択されませんでした。"}
        
        # フォルダ選択後、My Documents内のPDFファイルをすべて削除
        from utils.cleanup_pdf import delete_all_pdf_in_my_documents
        success, deleted_count, message = delete_all_pdf_in_my_documents()
        if success:
            print(f"✅ Step 1 PDF削除完了: {message}")
        else:
            print(f"⚠️ Step 1 PDF削除失敗: {message}")

        # Excel名からクリーンな名前を生成
        excel_name_clean = excel_name.split("-", 1)[1]
        excel_name_clean = re.sub(r"\(.*?\)", "", excel_name_clean).strip()

        # サブフォルダ作成
        target_folder = os.path.join(output_folder, excel_name_clean)
        target_folder = os.path.normpath(target_folder)
        
        if os.path.exists(target_folder):
            return {"error": f"フォルダ '{excel_name_clean}' は既に存在します。別のExcelファイルを選択してください。"}
        os.makedirs(target_folder)

        # Excel列K読み込み
        df = pd.read_excel(excel_path, usecols="K,AB,AD", skiprows=7, header=None)
        df.columns = ["K", "AB", "AD"]
        df["AD"] = df["AD"].astype(str).str.strip()

        # フィルタリング前に保持と追加を確認する
        skipped_due_to_hold = df[df["AD"] == "保留"]["K"].dropna().astype(str).str.strip().tolist()
        added_due_to_addition = df[df["AD"] == "追加"]["K"].dropna().astype(str).str.strip().tolist()

        # 「追加」の有無を確認する
        has_addition = not df[df["AD"] == "追加"].empty
        
        if has_addition:
            filtered_df = df[(df["AD"] == "追加") & (df["K"].notna()) & (df["K"] != 0)]
        else:
            filtered_df = df[
                (df["K"].notna()) & (df["K"] != 0) &
                (df["AD"] != "保留") &
                (df["AB"].notna()) & (df["AB"] != "") & (df["AB"].astype(str) != "0")
            ]

        if filtered_df.empty:
            messagebox.showwarning("データなし","Excelファイルに有効な部品番号がありません。")
            return None
        
        # ICDファイルコピー
        copied_files = []
        not_found = []
        for _, row in filtered_df.iterrows():
            part_number = str(row["K"]).strip()
            status_ad = row["AD"]
            found_file = search_number(part_number)  # 専用機
            if not found_file:
                found_file = search_gradually('Y:\\標準機\\', part_number)  # 標準機
            if found_file:
                shutil.copy(found_file, target_folder)
                copied_files.append(found_file)
            else:
                not_found.append(part_number)

        # config.txt保存
        with open(os.path.join(target_folder, "config.txt"), "w", encoding="utf-8") as f:
            for file in copied_files:
                f.write(file + "\n")

        print(f"Copied {len(copied_files)} files to {target_folder}")
        print(f"\n{'='*60}")
        print(f"📁 フォルダパス: {target_folder}")
        print(f"{'='*60}\n")
        
        if skipped_due_to_hold:
            print("❌ 保留のためコピーしなかった部品番号:")
            for part in skipped_due_to_hold:
                print(f" - {part}")

        if added_due_to_addition:
            print("✅ 追加が指定されたため、コピーした部品番号:")
            for part in added_due_to_addition:
                print(f" + {part}")

        return {
            "output_folder": target_folder,
            "excel_name_clean": excel_name_clean,
            "copied_count": len(copied_files),
            "not_found": not_found,
            "icd_list": copied_files,
            "skipped_due_to_hold": skipped_due_to_hold,
            "added_due_to_addition": added_due_to_addition
        }

    except Exception as e:
        return {"error": f"Excelモード処理エラー: {str(e)}"}
