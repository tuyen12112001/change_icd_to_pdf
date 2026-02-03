import os
import re
from pathlib import Path


def remove_suffix_3d_from_pdf(output_folder):
    """
    PDFファイル名から "-3D" というサフィックスを削除する
    例: "drawing-3D.pdf" → "drawing.pdf"
    
    Args:
        output_folder (str): 処理対象フォルダのパス
        
    Returns:
        dict: {old_filename: new_filename} の形式で名前変更ログを返す
    """
    pattern = re.compile(r"^(?P<base>.+?)-3D(?P<ext>\.[^.]+)$", re.IGNORECASE)
    rename_log = {}
    
    try:
        for fname in os.listdir(output_folder):
            if fname.lower().endswith(".pdf"):
                m = pattern.match(fname)
                if m:
                    base = m.group("base")
                    ext = m.group("ext")
                    candidate_name = base + ext
                    
                    src_path = os.path.join(output_folder, fname)
                    dst_path = os.path.join(output_folder, candidate_name)
                    
                    # 重複チェック
                    if os.path.exists(dst_path) and src_path != dst_path:
                        print(f"⚠ スキップ: '{candidate_name}' は既に存在します")
                        continue
                    
                    # ファイル名変更
                    try:
                        os.rename(src_path, dst_path)
                        rename_log[fname] = candidate_name
                        print(f"✅ 名前変更: {fname} → {candidate_name}")
                    except Exception as e:
                        print(f"❌ 名前変更失敗: {fname} ({str(e)})")
    except Exception as e:
        print(f"❌ PDF フォルダの読み込みに失敗しました: {str(e)}")
    
    return rename_log
