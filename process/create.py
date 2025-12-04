from tkinter import filedialog, messagebox
import os
import shutil
import pandas as pd
import re
from utils.searchTools import search_number, search_gradually

def step1_create_and_copy(excel_path):
    """
    ステップ1:
    - ユーザーに出力フォルダを選択させる。
    - Excelファイルから列Kを読み込み。
    - ICDフォルダから一致する.icdファイルをコピー。
    - config.txtを保存。
    戻り値:
        成功時: {output_folder, excel_name_clean, copied_count}
        エラー時: {"error": "エラーメッセージ"}
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
        df = pd.read_excel(excel_path, usecols="K,AB", skiprows=7, header=None)
        df.columns = ["K", "AB"]

        
        # AB列の値を数値に変換する
        df["AB"] = pd.to_numeric(df["AB"], errors="coerce").fillna(0)


        # フィルタリング: Kが有効で、ABが0でない行のみ
        filtered_df = df[(df["K"].notna()) & (df["K"] != 0) & (df["AB"] != 0)]


        values = [str(val).strip() for val in filtered_df["K"].tolist()]
        
        if not values:
            messagebox.showwarning ("データなし","Excelファイルに有効な部品番号がありません。")
            return None
        # ICDファイルコピー
        copied_files = []
        not_found = []
        for val in values:
            found_file = search_number(val)  # 専用機
            if not found_file:
                found_file = search_gradually('Y:\\標準機\\', val)  # 標準機
            if found_file:
                shutil.copy(found_file, target_folder)
                copied_files.append(found_file)
            else:
                not_found.append(val)


        # config.txt保存
        with open(os.path.join(target_folder, "config.txt"), "w", encoding="utf-8") as f:
            for file in copied_files:
                f.write(file + "\n")


        print(f"Copied {len(copied_files)} files to {target_folder}")
        return {
            "output_folder": target_folder,
            "excel_name_clean": excel_name_clean,
            "copied_count": len(copied_files),
            "not_found": not_found
        }

    except Exception as e:
        return {"error": f"エラーが発生しました: {str(e)}"}