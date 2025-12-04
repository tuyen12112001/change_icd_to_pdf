
# excel_collect.py
import os
import re
import shutil
import unicodedata

def normalize_text(s: str) -> str:
    return unicodedata.normalize('NFKC', s or '').strip()

def next_nonconflict_path(dst_path: str) -> str:
    if not os.path.exists(dst_path):
        return dst_path
    base, ext = os.path.splitext(os.path.basename(dst_path))
    dir_ = os.path.dirname(dst_path)
    i = 1
    while True:
        candidate = os.path.join(dir_, f"{base} ({i}){ext}")
        if not os.path.exists(candidate):
            return candidate
        i += 1

def extract_id_from_name(file_name: str) -> str | None:
    name = normalize_text(os.path.splitext(os.path.basename(file_name))[0])
    m = re.match(r'^(?:LS|LK)[-_ ]?([0-9A-Za-z]+)', name, flags=re.IGNORECASE)
    return m.group(1) if m else None

def add_ls_lk_excel_set_to_output(excel_path: str, output_dir: str,
                                  include_selected: bool = True, recursive: bool = False) -> list[str]:
    copied = []
    if not excel_path or not os.path.isfile(excel_path):
        print("❌ 指定Excelが不正です。")
        return copied
    if not output_dir or not os.path.isdir(output_dir):
        print(f"❌ 出力フォルダが存在しません: {output_dir}")
        return copied

    base_dir = os.path.dirname(excel_path)
    id_part = extract_id_from_name(os.path.basename(excel_path))
    if not id_part:
        print("⚠ 指定Excel名から ID（LS-/LK-の直後）が取得できませんでした。例：LS-1234_... の形式にしてください。")
        return copied

    import re
    ls_pat = re.compile(rf'^LS[-_ ]?{re.escape(id_part)}', flags=re.IGNORECASE)
    lk_pat = re.compile(rf'^LK[-_ ]?{re.escape(id_part)}', flags=re.IGNORECASE)

    walker = (os.walk(base_dir) if recursive else [(base_dir, [], os.listdir(base_dir))])
    targets = []
    for root, _, files in walker:
        for name in files:
            ext = os.path.splitext(name)[1].lower()
            if ext not in (".xlsx", ".xls", ".xlsm"):
                continue
            name_norm = normalize_text(os.path.splitext(name)[0])
            if ls_pat.match(name_norm) or lk_pat.match(name_norm):
                targets.append(os.path.join(root, name))
        if not recursive:
            break

    if not include_selected:
        targets = [p for p in targets if os.path.abspath(p) != os.path.abspath(excel_path)]

    if not targets:
        print("ℹ 一致する Excel が見つかりませんでした。")
        return copied

    for src in targets:
        try:
            fname = os.path.basename(src)
            dst = os.path.join(output_dir, fname)
            dst = next_nonconflict_path(dst)
            shutil.copy2(src, dst)
            copied.append(dst)
            print(f"✅ コピー: {src} -> {dst}")
        except Exception as e:
            print(f"⚠ コピー失敗: {src} | エラー: {e}")

    print(f"📦 合計 {len(copied)} 件をコピーしました。")
    return copied
