
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
import shutil
import pandas as pd
import re
import glob
import sys
import win32com.client
import subprocess

# ===================== Logic tìm kiếm =====================
def search_gradually(base_path, workpiece):
    max_depth = 4
    found_file = None
    black_keywords = ['.pdf', '-OLD', '.xls', '-old','-E']
    
    # 括弧を除去（[ABC123D] → ABC123D）
    if workpiece.startswith('[') and workpiece.endswith(']'):
        workpiece = workpiece[1:-1]

    for depth in range(1, max_depth + 1):
        pattern = base_path + ('*\\' * depth) + f'*{workpiece}*.icd'
        for f in glob.iglob(pattern, recursive=True):
            if any(keyword in os.path.basename(f) for keyword in black_keywords):
                continue
            filename = os.path.splitext(os.path.basename(f))[0]
            # 完全一致のファイル名を優先
            if filename == workpiece:
                return f
            # workpieceで始まるファイル名のみ取得（接頭辞なし）
            if filename.startswith(workpiece):
                # (M)、-OLDなど不要な接尾辞を除外
                suffix = filename[len(workpiece):]
                # 接尾辞が-3D、-2Dなどのみ許可
                if re.fullmatch(r'-\w+', suffix):
                    if not found_file:
                        found_file = f
        if found_file:
            break
    return found_file

def search_number(workpiece: str, mode: str = "0"):
    if mode == "0":
        # 専用機モード（従来の検索方法）
        try:
            split_workpiece = workpiece.split('-')
            zuban = split_workpiece[0]
            seiban = split_workpiece[1]
            try:
                hinban = split_workpiece[2]
            except IndexError:
                hinban = None

            seiban_alp = seiban[:3]
            seiban_num = seiban[3:7]
            num_front = seiban_num[:2]
            num_back = seiban_num[2:]

            # TSZ\12*
            path_1 = f'Y:\\専用機\\{seiban_alp}\\{num_front}*'
            file_1 = next(glob.iglob(path_1), None)

            if not file_1:
                raise Exception('ファイルが見つかりません (path_1)')

            # 1234*
            path_2 = f'{file_1}\\{seiban_num}*'
            file_2 = next(glob.iglob(path_2), None)

            if not file_2:
                raise Exception('ファイルが見つかりません (path_2)')

            # .lnk 処理
            path_3 = f'{file_2}\\*.lnk'
            file_3 = next(glob.iglob(path_3), None)

            if file_3:
                wshell = win32com.client.Dispatch("WScript.Shell")
                shortcut = wshell.CreateShortcut(file_3)
                file_2 = shortcut.TargetPath

            # 実ファイル検索
            path_4 = f'{file_2}\\**\\*{workpiece}*'
            file_4 = None
            for file in glob.iglob(path_4, recursive=True):
                if '-OLD' in file:
                    continue
                if not file.lower().endswith('.icd'):
                    continue

                filename = os.path.splitext(os.path.basename(file))[0]
                parts = filename.split('-')

                # aaaa-aaa1234 以上は一致しているか
                if len(parts) < 2:
                    continue

                if parts[0] != zuban or parts[1] != seiban:
                    continue

                # hinban 判定
                if hinban is None:
                    # 完全に2要素のみ許可、または 3要素で3番目が "3D" の場合も許可
                    # 例: A-B または A-B-3D
                    if len(parts) == 2:
                        file_4 = file
                        break
                    elif len(parts) == 3 and parts[2].upper() == "3D":
                        file_4 = file
                        break
                else:
                    # hinban が指定されているなら3要素目以降OK
                    if len(parts) >= 3:
                        file_4 = file
                        break

            if not file_4:
                raise Exception('検索図番が見つかりません\n例：\n・ショートカットがない\n・現行の明和品番でない。')

            return file_4
        except Exception as err:
            return None
    else:
        # 標準機モード（search_gradually使用）
        try:
            default_path = 'Y:\\標準機\\'
            file_1 = search_gradually(default_path, workpiece)

            if not file_1:
                raise Exception('検索図番が見つかりません。標準機モードは完全一致検索です。')

            return file_1
        except Exception as err:
            return None

