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
from config.settings import IMAGE3_PATH, DOCUWORKS_TARGET_FOLDER
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
    Incremental "gọi tên và chuyển" watcher.

    Mỗi lần quét My Documents:
      1. Đếm số file PDF khớp tên ICD đã xuất hiện (tracking_count).
      2. Nếu tracking_count tăng lên so với lần trước (ví dụ 1→2):
           → File thứ 1 (cũ nhất) đã có đủ thời gian, kiểm tra và chuyển nó đi.
      3. Lặp lại cho đến khi tất cả file được chuyển hoặc hết timeout.
    """
    if not destination_dir or not os.path.isdir(destination_dir):
        print(f"⚠ Invalid destination: {destination_dir}")
        return []

    if not my_docs_dir or not os.path.isdir(my_docs_dir):
        print(f"⚠ Invalid source (My Documents): {my_docs_dir}")
        return []

    icd_names = _normalize_icd_names(icd_list)
    expected_count = len(icd_names)
    moved_paths = []
    waiting_bases = set()   # file đã xuất hiện nhưng chưa chuyển (theo dõi thứ tự)
    moved_bases = set()     # file đã chuyển xong
    start = time.time()

    def _current_available():
        """Trả về (matched_dict, all_pdf_names_in_source)"""
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

    # Lần quét đầu tiên
    initial, initial_all = _current_available()
    waiting_bases = set(initial.keys())
    prev_count = len(waiting_bases)
    if prev_count > 0:
        print(f"🆕 Ban đầu phát hiện {prev_count} file: {sorted(waiting_bases)}")
    else:
        print(f"ℹ️ Lần quét đầu: chưa thấy file nào (tổng PDF trong folder: {len(initial_all)})")

    while time.time() - start < timeout_sec:
        if emergency_manager.is_stop_requested():
            break

        available, all_pdfs = _current_available()
        current_count = len(available)

        # Debug: in tất cả PDF trong folder mỗi 5 giây
        if int(time.time() - start) % 5 == 0:
            print(f"   📋 Folder có {len(all_pdfs)} PDF: {sorted(set(os.path.splitext(n)[0].strip().lower() for n in all_pdfs))[:10]}")

        # Kiểm tra điều kiện kết thúc
        if not available:
            if waiting_bases:
                print(f"⏳ Hết file mới, {len(waiting_bases)} file đang chờ ổn định...")
                # thử di chuyển lần cuối
                for base in sorted(waiting_bases):
                    # tìm path trong all_pdfs
                    src = None
                    for name in all_pdfs:
                        if os.path.splitext(name)[0].strip().lower() == base:
                            src = os.path.join(my_docs_dir, name)
                            break
                    if src and _is_file_ready(src, stable_checks=2, wait_interval=0.3):
                        try:
                            dst = _get_nonconflict_path(destination_dir, os.path.basename(src))
                            shutil.move(src, dst)
                            moved_bases.add(base)
                            waiting_bases.discard(base)
                            moved_paths.append(dst)
                            print(f"✅ Di chuyển (lần cuối): {os.path.basename(src)}")
                        except Exception as e:
                            print(f"❌ Lỗi: {os.path.splitext(os.path.basename(src))[0]} / {e}")
                            waiting_bases.discard(base)
            if not waiting_bases:
                break
            if len(moved_paths) >= expected_count:
                break
            time.sleep(poll_interval)
            continue

        # Nếu số file hiện có TĂNG lên so với lần trước
        # → file cũ nhất trong waiting_bases đã có thêm thời gian ổn định
        if current_count > prev_count:
            if waiting_bases:
                oldest = next(iter(sorted(waiting_bases)))  # file cũ nhất theo alphabet
                src_path = available.get(oldest)
                if src_path:
                    try:
                        if _is_file_ready(src_path, stable_checks=2, wait_interval=0.3):
                            dst = _get_nonconflict_path(destination_dir, os.path.basename(src_path))
                            shutil.move(src_path, dst)
                            moved_bases.add(oldest)
                            waiting_bases.discard(oldest)
                            moved_paths.append(dst)
                            print(f"✅ Di chuyển (file {len(moved_paths)}/{expected_count}): {os.path.basename(src_path)}")
                        else:
                            print(f"⏳ Chưa ổn định, giữ lại: {os.path.basename(src_path)}")
                    except Exception as e:
                        print(f"❌ Lỗi di chuyển: {os.path.basename(src_path)} / {e}")
                        waiting_bases.discard(oldest)

            # Cập nhật waiting_bases: những file mới chưa từng được ghi nhận
            new_bases = set(available.keys()) - waiting_bases - moved_bases
            waiting_bases |= new_bases
            if new_bases:
                print(f"🆕 Thêm {len(new_bases)} file vào danh sách chờ: {sorted(new_bases)}")

        # Nếu số file không tăng nhưng có file cũ chưa di chuyển,
        # thử di chuyển ngay
        elif waiting_bases and current_count == prev_count:
            still_waiting = sorted(waiting_bases & set(available.keys()))
            if still_waiting:
                oldest = still_waiting[0]
                src_path = available.get(oldest)
                if src_path:
                    try:
                        if _is_file_ready(src_path, stable_checks=2, wait_interval=0.3):
                            dst = _get_nonconflict_path(destination_dir, os.path.basename(src_path))
                            shutil.move(src_path, dst)
                            moved_bases.add(oldest)
                            waiting_bases.discard(oldest)
                            moved_paths.append(dst)
                            print(f"✅ Di chuyển (file {len(moved_paths)}/{expected_count}): {os.path.basename(src_path)}")
                    except Exception as e:
                        print(f"❌ Lỗi di chuyển: {os.path.basename(src_path)} / {e}")
                        waiting_bases.discard(oldest)

        prev_count = current_count

        if len(moved_paths) >= expected_count:
            print(f"✅ Tất cả {expected_count} file đã di chuyển.")
            break

        time.sleep(poll_interval)

    if len(moved_paths) < expected_count:
        print(f"⚠ Chưa đủ: {len(moved_paths)}/{expected_count} (timeout hoặc dừng)")

    return moved_paths


def _watch_and_move_incremental_multi(icd_list, destination_dir, scan_dirs,
                                       poll_interval=1.0, timeout_sec=300):
    """
    Quét nhiều thư mục cùng lúc để tìm file PDF của ICAD.

    Mỗi lần quét:
      - Gộp kết quả từ tất cả scan_dirs → {base_name: (src_path, dir_index)}
      - Phát hiện file MỚI (chưa từng thấy trước đó)
      - Khi số file tăng → lấy file cũ nhất trong danh sách, chờ ổn định, di chuyển
      - Khi không còn file mới sau 3s và đã di chuyển ít nhất 1 file → kết thúc

    scan_dirs: list of folder paths (ưu tiên theo thứ tự trong list)
    """
    if not destination_dir or not os.path.isdir(destination_dir):
        print(f"⚠ Invalid destination: {destination_dir}")
        return []

    icd_names = _normalize_icd_names(icd_list)
    expected_count = len(icd_names)
    moved_paths = []
    waiting_bases = {}   # base_name → (src_path, src_dir)
    moved_bases = set()
    start = time.time()
    prev_total_count = 0
    last_new_sighting = None

    def _collect_all():
        """Trả về (matched_dict{base: (src_path, dir_idx)}, total_pdf_count) — quét cả subfolder"""
        found = {}
        total = 0
        for di, d in enumerate(scan_dirs):
            try:
                for root, _, files in os.walk(d):
                    for name in files:
                        if not name.lower().endswith(".pdf"):
                            continue
                        total += 1
                        base = os.path.splitext(name)[0].strip().lower()
                        if base in icd_names and base not in moved_bases and base not in found:
                            found[base] = (os.path.join(root, name), di)
            except Exception:
                continue
        return found, total

    # Lần quét đầu
    available, total_now = _collect_all()
    waiting_bases = available.copy()
    prev_total_count = total_now
    sources_used = sorted(set(di for _, di in available.values()))
    print(f"👁  Multi-watcher: {len(available)}/{expected_count} file tìm thấy lần đầu (từ thư mục {sources_used}, tổng PDF: {total_now})")

    while time.time() - start < timeout_sec:
        if emergency_manager.is_stop_requested():
            break

        available, total_now = _collect_all()
        current_count = len(available)

        # Debug mỗi 5s
        if int(time.time() - start) % 5 == 0:
            dir_summary = {}
            for base, (path, di) in available.items():
                dir_summary.setdefault(scan_dirs[di], []).append(base)
            for d, bases in dir_summary.items():
                print(f"   📋 {os.path.basename(d)}: {sorted(bases)[:8]}")
            if not dir_summary:
                print(f"   ⚠️ Không tìm thấy file nào trong các thư mục quét.")

        if not available and not waiting_bases:
            break

        # ---------- di chuyển file cũ nhất khi số file TĂNG ----------
        if current_count > prev_total_count:
            # Số file trong tất cả thư mục tăng → file cũ nhất ổn định rồi
            all_waiting = set(waiting_bases) | set(available)
            oldest = next(iter(sorted(all_waiting)))
            src_path, _ = available.get(oldest) or waiting_bases.get(oldest)
            if src_path:
                moved_bases = _try_move(oldest, src_path, destination_dir, moved_bases, moved_paths, expected_count)
            prev_total_count = total_now
            last_new_sighting = time.time()
            # Cập nhật waiting_bases với file mới tìm thấy
            waiting_bases.update(available)
            continue

        # ---------- Nếu không tăng, thử di chuyển file cũ nhất ----------
        if waiting_bases:
            # Ưu tiên file cũ nhất theo thứ tự alphabet
            all_ordered = sorted(set(list(waiting_bases.keys()) + list(available.keys())))
            for base in all_ordered:
                src_path, _ = available.get(base) or waiting_bases.get(base)
                if src_path:
                    result = _try_move(base, src_path, destination_dir, moved_bases, moved_paths, expected_count)
                    if result:
                        break

        # ---------- Cập nhật waiting_bases ----------
        waiting_bases.update(available)
        # Xóa những file đã di chuyển
        waiting_bases = {k: v for k, v in waiting_bases.items() if k not in moved_bases}

        prev_total_count = max(prev_total_count, total_now)

        if len(moved_paths) >= expected_count:
            print(f"✅ Tất cả {expected_count} file đã di chuyển.")
            break

        time.sleep(poll_interval)

    remaining = expected_count - len(moved_paths)
    if remaining > 0:
        print(f"⚠ Chưa đủ: {len(moved_paths)}/{expected_count} (thiếu {remaining})")

    return moved_paths


def _try_move(base_name, src_path, destination_dir, moved_bases, moved_paths, expected_count):
    """Thử di chuyển 1 file. Trả về True nếu di chuyển thành công."""
    try:
        if not os.path.isfile(src_path):
            return False
        if not _is_file_ready(src_path, stable_checks=2, wait_interval=0.3):
            print(f"⏳ Chưa ổn định: {os.path.basename(src_path)}")
            return False
        dst = _get_nonconflict_path(destination_dir, os.path.basename(src_path))
        shutil.move(src_path, dst)
        moved_bases.add(base_name)
        moved_paths.append(dst)
        print(f"✅ Di chuyển ({len(moved_paths)}/{expected_count}): {os.path.basename(src_path)}")
        return True
    except Exception as e:
        print(f"❌ Lỗi di chuyển: {os.path.basename(src_path)} / {e}")
        return False


# ===========================================================
# ✅ Step 2: Print
# ===========================================================

def step2_print_icd(output_dir, excel_name_clean, icd_list=None):
    """
    Step 2 — Print from ICAD to DocuWorks:
      1. Create a new folder in DocuWorks.
      2. Drive ICAD (Alt+F→D, paste output path, select DocuWorks PDF printer,
         Shift+End, Enter).
      3. After printing, poll My Documents using ICD name matching
         and move confirmed PDFs into the DocuWorks folder.
      4. Return summary dict.

    Parameters:
        output_dir:    Output folder path
        excel_name_clean: Cleaned Excel name
        icd_list:      List of source ICD file paths (used for filename matching)
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

        # 3. Open ICAD print dialog  (Alt+F → D)
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

        # 5. Open print options  (Alt+O)
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

        # 6. Select all files  (Home → Shift+End → Enter to start print)
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
        # Step 2-B: Incremental watch — move as soon as each file is stable
        # -----------------------------------------------------------
        # ICAD print output có thể xuất hiện ở nhiều nơi:
        #   DOCUWORKS_TARGET_FOLDER (cấu hình trong settings.py)
        #   → output_dir (đường dẫn ICAD vừa dán)
        #   → My Documents (fallback cuối cùng)
        watch_candidates = [
            d for d in [
                DOCUWORKS_TARGET_FOLDER,
                output_dir,
                get_docuworks_pdf_folder(),
            ] if d and os.path.isdir(d)
        ]

        if not watch_candidates:
            print("⚠ Không tìm thấy thư mục quét nào hợp lệ.")
            return None

        icd_names = list(_normalize_icd_names(icd_list or []))
        if not icd_names:
            print("⚠ ICD list is empty — nothing to move.")

        print(f"👁  Incremental watcher: {len(icd_names)} files expected")
        print(f"    Scan targets: {watch_candidates}")

        moved = _watch_and_move_incremental_multi(
            icd_list=icd_list,
            destination_dir=folder_path,
            scan_dirs=watch_candidates,
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
