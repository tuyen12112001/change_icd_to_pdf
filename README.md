# Shutsuzuu Tool (出図ツール) - ICAD to PDF Automation

## 📖 Giới thiệu
**Shutsuzuu Tool** (Change to PDF Tools v1.0) là công cụ tự động hóa giúp chuyển đổi hàng loạt bản vẽ kỹ thuật từ định dạng **ICAD (.icd)** sang **PDF** thông qua phần mềm trung gian **DocuWorks PDF**.

Công cụ giúp giảm thiểu thao tác thủ công, tự động tìm kiếm bản vẽ, in ấn, chuyển đổi định dạng và kiểm tra tính toàn vẹn của số lượng file đầu ra.

## 🚀 Tính năng chính
*   **Hai chế độ đầu vào:**
    *   **Excel Mode:** Đọc danh sách mã bản vẽ từ file Excel (BOM), lọc theo trạng thái (Bỏ qua "Hàng chờ/保留", xử lý "Thêm/追加").
    *   **Folder Mode:** Quét toàn bộ file `.icd` trong thư mục chỉ định.
*   **Tự động hóa in (Printing):** Điều khiển ICAD để in hàng loạt bản vẽ sang DocuWorks.
*   **Tự động hóa chuyển đổi (Conversion):** Điều khiển DocuWorks để chuyển đổi file XDW sang PDF.
*   **Xử lý file thông minh:**
    *   Tự động đổi tên file (loại bỏ hậu tố `-3D`).
    *   So sánh số lượng file đầu vào/đầu ra để cảnh báo thiếu/thừa.
    *   Dọn dẹp file rác (XDW, file tạm) sau khi hoàn tất.
*   **Giao diện trực quan:** Hỗ trợ Kéo & Thả (Drag & Drop), thanh tiến trình và log chi tiết.
*   **An toàn:** Có nút **Dừng khẩn cấp (Emergency Stop)** để ngắt quy trình ngay lập tức.

## 🛠 Yêu cầu hệ thống

### Phần mềm bắt buộc
1.  **Python 3.x**
2.  **ICAD (Micro Caelum):** Phải được cài đặt (đường dẫn mặc định: `C:\MC2\bin\icad.exe`).
3.  **Fuji Xerox DocuWorks:** Phải được cài đặt để xử lý in và chuyển đổi PDF.

### Thư viện Python
Cài đặt các thư viện cần thiết bằng lệnh sau:

```bash
pip install pandas openpyxl xlrd pywin32 pyautogui pynput opencv-python mss pygetwindow tkinterdnd2
```

## 📂 Cấu trúc dự án

```text
change_to_pdf_tools_version1.0/
├── main.py                     # Điểm khởi chạy ứng dụng
├── app/
│   └── main_app.py             # Giao diện chính (GUI)
├── config/
│   └── settings.py             # Cấu hình (Màu sắc, đường dẫn, text)
├── process/                    # Logic xử lý chính
│   ├── process_manager.py      # Quản lý luồng chạy (Step 1 -> Step 4)
│   ├── create.py               # Step 1: Copy file ICD
│   ├── printing.py             # Step 2: Auto click ICAD để in
│   ├── xdw_collection.py       # Step 3: Xử lý file XDW
│   ├── pdf_collection.py       # Step 3 & 4: Xử lý PDF trong DocuWorks
│   ├── clear.py                # Dọn dẹp file rác
│   └── rename_pdf.py           # Đổi tên file PDF
├── utils/                      # Các hàm tiện ích bổ trợ
│   ├── searchTools.py          # Tìm kiếm file ICD (Standard/Special)
│   ├── UI_helpers.py           # Hỗ trợ giao diện (Log, Loading, Blink)
│   ├── check_ICAD_and_Docuworks.py # Kiểm tra/Khởi động phần mềm
│   ├── docuworks_folder_creator.py # Tạo folder trong DocuWorks
│   ├── emergency_stop.py       # Quản lý dừng khẩn cấp
│   ├── excel_collect.py        # Thu thập file Excel liên quan
│   ├── excel_remove.py         # Lọc/Xóa file trong Excel mode
│   ├── file_compare.py         # So sánh số lượng file ICD/PDF
│   ├── cleanup_pdf.py          # Dọn dẹp file PDF/Rác
│   ├── cleanup_xdw.py          # Dọn dẹp file XDW
│   └── refresh_explore.py      # Làm mới Windows Explorer
└── assets/                     # Ảnh dùng cho Image Recognition
```

## 📖 Hướng dẫn sử dụng

### Bước 1: Khởi chạy
Chạy file `main.py`:
```bash
python main.py
```

### Bước 2: Chọn chế độ và Đầu vào
1.  **Excel File:**
    *   Kéo thả file Excel (BOM) vào ô nhập liệu.
    *   Tool đọc cột **K** (Mã bản vẽ) và cột **AD** (Trạng thái).
    *   *Yêu cầu:* Tên file Excel bắt đầu bằng `LS-`.
2.  **ICD Folder:**
    *   Kéo thả thư mục chứa các file `.icd` vào ô nhập liệu.

### Bước 3: Bắt đầu quy trình
Nhấn nút **"Bắt đầu" (Start)**.

1.  **Step 1 (Copy):** Tool tìm và copy file ICD vào thư mục đầu ra.
2.  **Step 2 (Printing):**
    *   Tool tự động mở ICAD và in sang DocuWorks.
    *   **Lưu ý:** Không chạm vào chuột/bàn phím lúc này.
3.  **Step 3 (Conversion):**
    *   Sau khi in xong, nhấn nút **"In xong" (Print Done)** trên giao diện.
    *   Tool điều khiển DocuWorks để convert sang PDF.
4.  **Step 4 (Exchange & Cleanup):**
    *   Nhấn nút **"Trao đổi xong" (Exchange Done)**.
    *   Tool di chuyển PDF về thư mục đích, đổi tên và dọn dẹp file rác.

### Bước 4: Kiểm tra kết quả
*   Tool hiển thị thông báo so sánh số lượng **ICD đầu vào** vs **PDF đầu ra**.
*   Nếu lệch, tool sẽ cảnh báo và liệt kê file thiếu/thừa.

## ⚠️ Lưu ý quan trọng
1.  **Auto-Click:** Tool chiếm quyền điều khiển chuột/bàn phím ở Step 2 & 3. **KHÔNG** thao tác máy tính trong lúc này.
2.  **Màn hình:** Tool sử dụng nhận diện hình ảnh (`assets/1.png`, `2.png`). Thay đổi độ phân giải màn hình có thể ảnh hưởng đến độ chính xác.
3.  **Dừng khẩn cấp:** Nhấn nút **"Dừng khẩn cấp" (Emergency Stop)** màu đỏ nếu gặp sự cố.

---
*Project developed for internal automation.*

---

# 日本語版 (Japanese Version)

# Shutsuzuu Tool (出図ツール) - ICAD to PDF Automation

## 📖 概要
**Shutsuzuu Tool** (Change to PDF Tools v1.0) は、**ICAD (.icd)** 形式の技術図面を **DocuWorks** を介して **PDF** に一括変換する自動化ツールです。

このツールは、手作業を減らし、図面の検索、印刷、フォーマット変換を自動化し、出力ファイル数の整合性をチェックします。

## 🚀 主な機能
*   **2つの入力モード:**
    *   **Excelモード:** Excelファイル (BOM) から図面番号リストを読み込み、ステータス（"保留"はスキップ、"追加"は処理）でフィルタリングします。
    *   **フォルダモード:** 指定されたフォルダ内のすべての `.icd` ファイルをスキャンします。
*   **印刷の自動化 (Printing):** ICADを制御して、図面をDocuWorksに一括印刷します。
*   **変換の自動化 (Conversion):** DocuWorksを制御して、XDWファイルをPDFに変換します。
*   **スマートなファイル処理:**
    *   ファイル名の自動変更（`-3D` 接尾辞の削除）。
    *   入力ファイル数と出力ファイル数を比較し、不足/過剰を警告します。
    *   完了後に不要なファイル（XDW、一時ファイル）を削除します。
*   **直感的なインターフェース:** ドラッグ＆ドロップ、プログレスバー、詳細ログをサポート。
*   **安全性:** プロセスを即座に中断するための **非常停止 (Emergency Stop)** ボタンがあります。

## 🛠 システム要件

### 必須ソフトウェア
1.  **Python 3.x**
2.  **ICAD (Micro Caelum):** インストール済みであること（デフォルトパス: `C:\MC2\bin\icad.exe`）。
3.  **Fuji Xerox DocuWorks:** 印刷およびPDF変換処理のためにインストール済みであること。

### Pythonライブラリ
以下のコマンドで必要なライブラリをインストールしてください:

```bash
pip install pandas openpyxl xlrd pywin32 pyautogui pynput opencv-python mss pygetwindow tkinterdnd2
```

## 📂 プロジェクト構成

```text
change_to_pdf_tools_version1.0/
├── main.py                     # アプリケーション起動ファイル
├── app/
│   └── main_app.py             # メインGUI
├── config/
│   └── settings.py             # 設定 (色, パス, テキスト)
├── process/                    # メイン処理ロジック
│   ├── process_manager.py      # プロセス管理 (Step 1 -> Step 4)
│   ├── create.py               # Step 1: ICDファイルコピー
│   ├── printing.py             # Step 2: ICAD自動印刷
│   ├── xdw_collection.py       # Step 3: XDWファイル処理
│   ├── pdf_collection.py       # Step 3 & 4: DocuWorksでのPDF処理
│   ├── clear.py                # ゴミファイル削除
│   └── rename_pdf.py           # PDFファイル名変更
├── utils/                      # ユーティリティ関数
│   ├── searchTools.py          # ICDファイル検索 (標準機/専用機)
│   ├── UI_helpers.py           # GUI補助 (ログ, ロード, 点滅)
│   ├── check_ICAD_and_Docuworks.py # ソフトウェア起動・確認
│   ├── docuworks_folder_creator.py # DocuWorksフォルダ作成
│   ├── emergency_stop.py       # 非常停止管理
│   ├── excel_collect.py        # 関連Excelファイル収集
│   ├── excel_remove.py         # Excelモード用ファイル整理
│   ├── file_compare.py         # ICD/PDFファイル数比較
│   ├── cleanup_pdf.py          # PDF/ゴミファイル削除
│   ├── cleanup_xdw.py          # XDWファイル削除
│   └── refresh_explore.py      # エクスプローラー更新
└── assets/                     # 画像認識用アセット
```

## 📖 使用方法

### ステップ 1: 起動
`main.py` を実行します:
```bash
python main.py
```

### ステップ 2: モードと入力の選択
1.  **Excelファイル:**
    *   Excelファイル (BOM) を入力欄にドラッグ＆ドロップします。
    *   ツールは **K列** (図番) と **AD列** (ステータス) を読み取ります。
    *   *要件:* Excelファイル名は `LS-` で始まる必要があります。
2.  **ICDフォルダ:**
    *   `.icd` ファイルを含むフォルダを入力欄にドラッグ＆ドロップします。

### ステップ 3: プロセス開始
**"開始" (Start)** ボタンを押します。

1.  **Step 1 (Copy):** ツールがICDファイルを検索し、出力フォルダにコピーします。
2.  **Step 2 (Printing):**
    *   ツールが自動的にICADを開き、DocuWorksに印刷します。
    *   **注意:** この間はマウスやキーボードに触れないでください。
3.  **Step 3 (Conversion):**
    *   印刷完了後、インターフェース上の **"印刷完了" (Print Done)** ボタンを押します。
    *   ツールがDocuWorksを制御してPDFに変換します。
4.  **Step 4 (Exchange & Cleanup):**
    *   **"交換完了" (Exchange Done)** ボタンを押します。
    *   ツールがPDFを目的のフォルダに移動し、名前を変更し、不要なファイルを削除します。

### ステップ 4: 結果確認
*   ツールは **入力ICD数** と **出力PDF数** を比較して表示します。
*   不一致の場合、警告と不足/過剰ファイルのリストを表示します。

## ⚠️ 重要な注意点
1.  **自動クリック:** ツールはStep 2と3でマウス/キーボードを制御します。この間はPCを操作 **しないで** ください。
2.  **画面解像度:** ツールは画像認識 (`assets/1.png`, `2.png`) を使用します。画面解像度を変更すると精度に影響する可能性があります。
3.  **非常停止:** 問題が発生した場合は、赤色の **"非常停止" (Emergency Stop)** ボタンを押してください。
