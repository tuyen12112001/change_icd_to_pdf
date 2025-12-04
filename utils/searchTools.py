
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
    black_keywords = ['.pdf', '-OLD', '.xls']

    for depth in range(1, max_depth + 1):
        pattern = base_path + ('*\\' * depth) + f'*{workpiece}*'
        for f in glob.iglob(pattern, recursive=True):
            if not any(keyword in os.path.basename(f) for keyword in black_keywords):
                found_file = f
                break
        if found_file:
            break
    return found_file

def search_number(workpiece: str):
    try:
        split_workpiece = workpiece.split('-')
        if len(split_workpiece) < 3:
            return None

        seiban = split_workpiece[1]
        seiban_alp = seiban[:3]
        seiban_num = seiban[3:7]
        num_front = seiban_num[:2]

        path_1 = f'Y:\\専用機\\{seiban_alp}\\{num_front}*'
        file_1 = next(glob.iglob(path_1), None)
        if not file_1:
            return None

        path_2 = f'{file_1}\\{seiban_num}*'
        file_2 = next(glob.iglob(path_2), None)
        if not file_2:
            return None

        path_3 = f'{file_2}\\*.lnk'
        file_3 = next(glob.iglob(path_3), None)
        if file_3:
            wshell = win32com.client.Dispatch("WScript.Shell")
            shortcut = wshell.CreateShortcut(file_3)
            file_2 = shortcut.TargetPath

        path_4 = f'{file_2}\\**\\*{workpiece}*'
        for file in glob.iglob(path_4, recursive=True):
            if '-OLD' in file:
                continue
            if '.icd' in file:
                return file
        return None
    except Exception:
        return None

