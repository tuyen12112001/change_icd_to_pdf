# process/printing.py
import os
import shutil
import pyautogui
import pyperclip
import time
from pynput.keyboard import Key, Controller
import mss
import cv2
import numpy as np
from utils.check_ICAD_and_Docuworks import ensure_docuworks_running, ensure_icad_running
from utils.emergency_stop import emergency_manager
from config.settings import IMAGE3_PATH
from utils.docuworks_folder_creator import create_docuworks_folder_unique
from utils.cleanup_pdf import get_docuworks_pdf_folder

keyboard = Controller()

# ===========================================================
# Image Recognition helpers
# ===========================================================

def locate_center_mss(template_path, threshold=0.80):
    """
    Scan full screen with MSS + OpenCV template matching.
    Returns center (x, y) of the best match, or None if not found.
    """
    try:
        with mss.mss() as sct:
            monitor = sct.monitors[0]
            screenshot = np.array(sct.grab(monitor))
            screenshot = cv2.cvtColor(screenshot, cv2.COLOR_BGRA2BGR)

        template = cv2.imread(template_path, cv2.IMREAD_COLOR)
        if template is None:
            print(f"⚠ Cannot read image: {template_path}")
            return None

        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(result)

        if max_val < threshold:
            return None

        th, tw = template.shape[:2]
        center_x = max_loc[0] + tw // 2 + monitor['left']
        center_y = max_loc[1] + th // 2 + monitor['top']
        return (center_x, center_y)
    except Exception as e:
        print(f"⚠ locate_center_mss error: {e}")
        return None


def click_one_of_images(image_paths, max_attempts=10, confidence=0.80, wait_time=1):
    """Click the first matching image found, retry up to max_attempts times."""
    for attempt in range(max_attempts):
        for image_path in image_paths:
            try:
                loc = locate_center_mss(image_path, threshold=confidence)
                if loc:
                    pyautogui.click(loc)
                    print(f"🖱 Clicked: {os.path.basename(image_path)} (attempt {attempt + 1})")
                    return True
            except Exception as e:
                print(f"⚠ Image search error: {e}")
        print(f"⏳ Not found (attempt {attempt + 1}/{max_attempts})...")
        time.sleep(wait_time)
    print("❌ No image matched")
    return False


# ===========================================================
# File helpers
# ===========================================================

def _is_file_ready(file_path, stable_checks=2, wait_interval=0.3):
    """Wait until the file size is stable and readable."""
    last_size = -1
    stable_count = 0

    for _ in range(20):
        if not os.path.isfile(file_path):
            return False

        try:
            current_size = os.path.getsize(file_path)
            if current_size > 0 and current_size == last_size:
                stable_count += 1
            else:
                stable_count = 0
            last_size = current_size

            with open(file_path, "rb"):
                pass

            if stable_count >= stable_checks:
                return True
        except Exception:
            stable_count = 0

        time.sleep(wait_interval)

    return False


def _get_nonconflict_path(destination_dir, filename):
    """Return a path in destination_dir that does not collide with existing files."""
    base, ext = os.path.splitext(filename)
    candidate = os.path.join(destination_dir, filename)
    index = 1
    while os.path.exists(candidate):
        candidate = os.path.join(destination_dir, f"{base}_{index}{ext}")
        index += 1
    return candidate


# ===========================================================
# ICD name matching helpers
# ===========================================================

def _normalize_icd_names(icd_list):
    """Return a set of lowercase base-names (no extension) from icd_list."""
    result = set()
    for p in icd_list or []:
        base = os.path.splitext(os.path.basename(p))[0].strip().lower()
        result.add(base)
    return result


def _watch_and_move_incremental(icd_list, destination_dir, my_docs_dir,
                                 poll_interval=1.0, timeout_sec=300):
    """
    Incremental watcher — quét My Documents liên tục.

    Logic:
      - Bắt đầu với prev_count = 0
      - Mỗi lần quét, đếm số file PDF khớp tên ICD (current_count)
      - Nếu current_count = 1: KHÔNG di chuyển gì (chờ file thứ 2 để chắc file 1 ổn định)
      - Nếu current_count >= 2 và current_count > prev_count:
          -> Di chuyển tất cả file khớp mà chưa di chuyển
          -> Kiểm tra _is_file_ready() trước khi di chuyển
      - Lặp lại cho đến khi không còn file mới hoặc hết timeout
    """
    if not destination_dir or not os.path.isdir(destination_dir):
        print(f"⚠ Invalid destination: {destination_dir}")
        return []

    if not my_docs_dir or not os.path.isdir(my_docs_dir):
        print(f"⚠ Invalid source (My Documents): {my_docs_dir}")
        return []

    icd_names = _normalize_icd_names(icd_list)
    moved_paths = []
    moved_bases = set()
    start = time.time()
    prev_count = 0

    def _current_available():
        """Tra ve (matched_dict, all_pdf_names_in_source)"""
        result = {}
        all_pdfs = []
        try:
            for name in os.listdir(my_docs_dir):
                if not name.lower().endswith(".pdf"):
                    continue
                all_pdfs.append(name)
                base = os.path.splitext(name)[0].strip().lower()
                if base in icd_names and base not in moved_bases:
                    result[base] = os.path.join(my_docs_dir, name)
        except Exception as e:
            print(f"⚠ Scan error: {e}")
        return result, all_pdfs

    print(f"👁  Incremental watcher: Start monitoring")
    print(f"    Scan target: {my_docs_dir}")

    while time.time() - start < timeout_sec:
        if emergency_manager.is_stop_requested():
            break

        available, all_pdfs = _current_available()
        # current_count = số file hiện có trong My Documents + số file đã di chuyển
        files_in_my_docs = len(available)
        files_moved = len(moved_bases)
        current_count = files_in_my_docs + files_moved

        # Debug moi 5 giay
        if int(time.time() - start) % 5 == 0:
            bases_str = sorted(set(os.path.splitext(n)[0].strip().lower() for n in all_pdfs))[:10]
            print(f"   📋 My Docs: {files_in_my_docs}, Di chuyen: {files_moved}, Tong: {current_count} | {sorted(available.keys())}")

        # Neu current_count = 1: KHONG di chuyen (cho file thu 2 de dam bao file 1 on dinh)
        if current_count == 1:
            print(f"⏳ Chi co 1 file, cho file thu 2 de dam bao on dinh...")
            prev_count = current_count
            time.sleep(poll_interval)
            continue

        # Neu current_count >= 2 va TANG so voi lan truoc
        if current_count >= 2 and current_count > prev_count:
            files_to_move = list(available.items())
            moved_count = 0

            for base, src_path in sorted(files_to_move):
                try:
                    if _is_file_ready(src_path, stable_checks=2, wait_interval=0.3):
                        dst = _get_nonconflict_path(destination_dir, os.path.basename(src_path))
                        shutil.move(src_path, dst)
                        moved_bases.add(base)
                        moved_paths.append(dst)
                        moved_count += 1
                        print(f"✅ Di chuyen (file {len(moved_paths)}): {os.path.basename(src_path)}")
                    else:
                        print(f"⏳ Chua on dinh, giu lai: {os.path.basename(src_path)}")
                except Exception as e:
                    print(f"❌ Loi di chuyen: {os.path.basename(src_path)} / {e}")
                    moved_bases.add(base)

            if moved_count > 0:
                print(f"🔄 Vua di chuyen {moved_count} file, Tong phat hien: {current_count} (My Docs: {files_in_my_docs}, Di chuyen: {files_moved})")

        # Neu khong con file nao khop
        if not available and moved_bases:
            print(f"✅ Het file moi, tong da di chuyen: {len(moved_paths)}")
            break

        prev_count = current_count
        time.sleep(poll_interval)

    print(f"✅ Watcher finished: {len(moved_paths)} files moved in {time.time() - start:.1f}s")
    return moved_paths


# ===========================================================
# Step 2: Print
# ===========================================================

def step2_print_icd(output_dir, excel_name_clean, icd_list=None):
    """
    Step 2 - Print from ICAD to DocuWorks:
      1. Create a new folder in DocuWorks.
      2. Drive ICAD (Alt+F->D, paste output path, select DocuWorks PDF printer,
         Shift+End, Enter).
      3. After printing, poll My Documents using ICD name matching
         and move confirmed PDFs into the DocuWorks folder.
      4. Return summary dict.

    Parameters:
        output_dir:       Output folder path
        excel_name_clean: Cleaned Excel name
        icd_list:         List of source ICD file paths (used for filename matching)
    """
    try:
        # 1. Ensure DocuWorks is running and create folder
        ensure_docuworks_running()
        time.sleep(0.5)

        folder_name, folder_path = create_docuworks_folder_unique(
            excel_name_clean, ensure_docuworks_running
        )

        if not folder_name:
            raise Exception("Failed to create DocuWorks folder.")
        if not folder_path:
            raise Exception(f"Failed to get folder path (folder_name={folder_name})")

        print(f"📂 DocuWorks Folder: {folder_path}")

        # 2. Ensure ICAD is running
        icad_path = r"C:\MC2\bin\icad.exe"
        ensure_icad_running(icad_path)
        time.sleep(1)

        if emergency_manager.is_stop_requested():
            print("⚠ Emergency stop before printer selection. Aborting.")
            return None

        # 3. Open ICAD print dialog (Alt+F -> D)
        pyautogui.keyDown('alt')
        pyautogui.press('f')
        time.sleep(0.2)
        pyautogui.press('d')
        pyautogui.keyUp('alt')
        time.sleep(0.5)

        if emergency_manager.is_stop_requested():
            return None

        # 4. Paste output folder path
        pyperclip.copy(output_dir)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.3)
        pyautogui.press('enter')
        time.sleep(2.0)

        if emergency_manager.is_stop_requested():
            return None

        # 5. Open print options (Alt+O)
        pyautogui.hotkey("alt", "o")
        time.sleep(0.5)

        if emergency_manager.is_stop_requested():
            return None

        # Drop-down in printer dialog
        keyboard.press(Key.alt)
        keyboard.press(Key.down)
        keyboard.release(Key.alt)
        time.sleep(0.5)

        click_one_of_images([IMAGE3_PATH], max_attempts=10, confidence=0.80, wait_time=1)

        if emergency_manager.is_stop_requested():
            return None

        pyautogui.press("enter")
        time.sleep(1)

        # 6. Select all files (Home -> Shift+End -> Enter to start print)
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

        print_started_at = time.time()
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)

        # ICAD job submission grace period
        print("⏳ ICAD print command — waiting 3 s ...")
        time.sleep(3)

        # -----------------------------------------------------------
        # Step 2-B: Incremental watch — chi quet My Documents
        # -----------------------------------------------------------
        my_docs_dir = get_docuworks_pdf_folder()

        if not my_docs_dir or not os.path.isdir(my_docs_dir):
            print("⚠ Khong tim thay thu muc My Documents hop le.")
            return None

        icd_names = list(_normalize_icd_names(icd_list or []))
        if not icd_names:
            print("⚠ ICD list is empty — nothing to move.")

        print(f"👁  Incremental watcher: {len(icd_names)} files expected")
        print(f"    Scan target: {my_docs_dir}")

        moved = _watch_and_move_incremental(
            icd_list=icd_list,
            destination_dir=folder_path,
            my_docs_dir=my_docs_dir,
        )
        printed_pdf_count = len(moved)

        print("✅ Print step finished.")
        if emergency_manager.is_stop_requested():
            return None

        return {
            "folder_name": folder_name,
            "folder_path": folder_path,
            "moved_files_count": printed_pdf_count,
            "printed_pdf_count": printed_pdf_count,
            "print_started_at": print_started_at,
        }

    except Exception as e:
        print(f"❌ Print error: {e}")
        return None
