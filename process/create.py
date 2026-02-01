from tkinter import filedialog, messagebox
import os
import shutil
import pandas as pd
import re
from utils.searchTools import search_number, search_gradually

def step1_create_and_copy(excel_path=None, icd_folder_path=None):
    """
    ステップ1:
    - Excelモード: Excelファイルから列Kを読み込みICDファイルをコピー
    - フォルダモード: 指定フォルダ内のすべての.icdファイルを直接コピー
    - ユーザーに出力フォルダを選択させる。
    - ICDフォルダから一致する.icdファイルをコピー。
    - config.txtを保存。
    戻り値:
        成功時: {output_folder, excel_name_clean, copied_count}
        エラー時: {"error": "エラーメッセージ"}
    """
    try:
        # モード判定
        if excel_path and not icd_folder_path:
            return _step1_excel_mode(excel_path)
        elif icd_folder_path and not excel_path:
            return _step1_folder_mode(icd_folder_path)
        else:
            return {"error": "ExcelファイルまたはICDフォルダのいずれかを選択してください。"}

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


def _step1_folder_mode(icd_folder_path):
    """
    フォルダモード処理: 指定フォルダ内の "drawing" サブフォルダを検索し、
    そこから.icdファイルを取得（ASで始まるファイルは除外、サブフォルダは除外）
    """
    try:
        # フォルダの存在確認
        if not os.path.isdir(icd_folder_path):
            return {"error": f"ICDフォルダが見つかりません: {icd_folder_path}"}

        # "drawing" フォルダを探す
        drawing_folder = os.path.join(icd_folder_path, "drawing")
        if not os.path.isdir(drawing_folder):
            return {"error": f"'drawing' フォルダが見つかりません: {drawing_folder}"}

        # ユーザーに出力フォルダを選択させる
        output_folder = filedialog.askdirectory(title="出力フォルダを選択してください")
        if not output_folder:
            return {"error": "出力フォルダが選択されませんでした。"}

        # フォルダ名からサブフォルダ名を生成
        parent_folder_name = os.path.basename(icd_folder_path.rstrip("\\").rstrip("/"))
        target_folder = os.path.join(output_folder, parent_folder_name)
        target_folder = os.path.normpath(target_folder)

        if os.path.exists(target_folder):
            return {"error": f"フォルダ '{parent_folder_name}' は既に存在します。別のフォルダを選択してください。"}
        
        os.makedirs(target_folder)

        # drawing フォルダ内のすべての.icdファイルを取得（ASで始まるものは除外、サブフォルダは除外）
        copied_files = []
        skipped_as_files = []  # AS で始まるファイルを記録
        
        icd_files = [f for f in os.listdir(drawing_folder) 
                     if os.path.isfile(os.path.join(drawing_folder, f)) and f.lower().endswith('.icd')]

        if not icd_files:
            return {"error": f"drawing フォルダに.icdファイルが見つかりません: {drawing_folder}"}

        for icd_file in icd_files:
            # AS で始まるファイルをスキップ
            if icd_file.upper().startswith('AS'):
                skipped_as_files.append(icd_file)
                print(f"⏭️ スキップ (AS): {icd_file}")
                continue

            src_path = os.path.join(drawing_folder, icd_file)
            dst_path = os.path.join(target_folder, icd_file)
            try:
                shutil.copy(src_path, dst_path)
                copied_files.append(src_path)
                print(f"✅ コピー: {icd_file}")
            except Exception as e:
                print(f"❌ コピー失敗: {icd_file} - {str(e)}")

        if not copied_files:
            return {"error": f"コピーするICDファイルが見つかりません（ASで始まるファイルを除外後）"}

        # config.txt保存
        with open(os.path.join(target_folder, "config.txt"), "w", encoding="utf-8") as f:
            for file in copied_files:
                f.write(file + "\n")

        print(f"Copied {len(copied_files)} files to {target_folder}")
        if skipped_as_files:
            print(f"⏭️ Skipped {len(skipped_as_files)} AS files:")
            for f in skipped_as_files:
                print(f"  - {f}")

        return {
            "output_folder": target_folder,
            "excel_name_clean": parent_folder_name,
            "copied_count": len(copied_files),
            "not_found": [],
            "icd_list": copied_files,
            "skipped_due_to_hold": skipped_as_files,  # AS files をこっちに入れる
            "added_due_to_addition": [],
            "skipped_as_files": skipped_as_files  # 追加で記録
        }

    except Exception as e:
        return {"error": f"フォルダモード処理エラー: {str(e)}"}