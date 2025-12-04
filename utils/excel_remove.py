
import os
import re
import unicodedata
from pathlib import Path
from typing import Tuple

VALID_EXT = (".xlsx", ".xls", ".xlsm")

def normalize_text(s: str) -> str:
    return unicodedata.normalize("NFC", s).strip().casefold()

def extract_id(name_no_ext: str) -> str:
    s = unicodedata.normalize("NFC", name_no_ext)
    s = re.sub(r"^(LS|LK)[-_ ]?", "", s, flags=re.IGNORECASE)  # LS/LK接頭辞を除去
    s = re.sub(r"\s*\(.*\)$", "", s)                           # 括弧書き "(1)" 等を除去
    return s

def excel_remove(folder_id: str, base_dir: str) -> Tuple[int, int, int]:

    # パスを絶対化
    base_dir = str(Path(base_dir).resolve())
    if not os.path.isdir(base_dir):
        raise FileNotFoundError(f"エラー: 指定したパスが存在しません → {base_dir}")

    fid = normalize_text(folder_id)
    kept = removed = skipped = 0

    for name in os.listdir(base_dir):
        ext = os.path.splitext(name)[1].lower()
        if ext not in VALID_EXT or name.startswith("~$"):
            skipped += 1
            continue

        name_no_ext = os.path.splitext(name)[0]
        if "(先行手配)" in unicodedata.normalize("NFC", name_no_ext):
            file_path = os.path.join(base_dir, name)
            try:
                os.remove(file_path)
                print(f"削除しました(先行手配): {file_path}")
                removed += 1
            except Exception as e:
                print(f"削除失敗(先行手配): {file_path} → {type(e).__name__}: {e}")
                skipped += 1
            continue



        id_extracted = extract_id(name_no_ext)
        match = normalize_text(id_extracted) == fid

        file_path = os.path.join(base_dir, name)
        if not match:
            try:
                os.remove(file_path)
                print(f"削除しました: {file_path}")
                removed += 1
            except Exception as e:
                print(f"削除失敗: {file_path} → {type(e).__name__}: {e}")
        else:
            print(f"保持しました: {file_path}")
            kept += 1

    print(f"\n結果: 保持={kept}, 削除={removed}, スキップ(非対象)={skipped}")
    return kept, removed, skipped
