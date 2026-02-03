import os
import shutil
import subprocess


def _force_delete_file(file_path):
    """
    ファイルを強制削除する（Windows del /F /Q コマンド使用）
    """
    try:
        # Windows の del コマンドで強制削除
        subprocess.run(
            ['del', '/F', '/Q', file_path],
            shell=True,
            capture_output=True,
            timeout=5,
            check=False
        )
        # subprocess が成功したかどうか確認
        if not os.path.exists(file_path):
            print(f"      ✅ 削除成功: {os.path.basename(file_path)}")
            return True
        else:
            print(f"      ❌ 削除失敗（ファイルが残っている）: {os.path.basename(file_path)}")
            return False
    except Exception as e:
        print(f"      ❌ 削除失敗: {type(e).__name__}: {str(e)}")
        return False


def step4_cleanup_pdf(output_dir):
    """
    ステップ4 (PDFモード用クリーンアップ):
    - 出力フォルダ内のすべてのファイルを確認
    - PDF ファイルと Excel ファイルのみを保持
    - その他のファイル（ICD, config.txt など）を削除
    """
    try:
        if not os.path.isdir(output_dir):
            print(f"❌ 出力フォルダが見つかりません: {output_dir}")
            return False

        kept_files = []
        deleted_files = []
        failed_files = []
        
        print(f"\n📂 クリーンアップ対象フォルダ: {output_dir}")
        print(f"📋 フォルダ内のファイル一覧:")
        
        for filename in os.listdir(output_dir):
            file_path = os.path.join(output_dir, filename)
            print(f"   - {filename}")
            
            # ディレクトリはスキップ
            if os.path.isdir(file_path):
                print(f"      ℹ️ スキップ（フォルダです）")
                continue
            
            # PDF ファイルまたは Excel ファイルは保持
            if filename.lower().endswith(('.pdf', '.xlsx', '.xls')):
                kept_files.append(filename)
                print(f"      ✅ 保持: {filename}")
            else:
                # その他のファイル（ICD, config.txt, 他）は削除
                print(f"      🗑️ 削除試行中...")
                success = _force_delete_file(file_path)
                if success:
                    deleted_files.append(filename)
                else:
                    print(f"      ❌ 削除失敗（諦めました）: {filename}")
                    failed_files.append((filename, "Force delete also failed"))

        print(f"\n📊 クリーンアップ完了:")
        print(f"   - 保持したファイル: {len(kept_files)} 件")
        print(f"   - 削除したファイル: {len(deleted_files)} 件")
        print(f"   - 削除失敗: {len(failed_files)} 件")
        
        if deleted_files:
            print(f"\n削除されたファイル:")
            for f in deleted_files:
                print(f"   - {f}")
        
        if failed_files:
            print(f"\n⚠️ 削除失敗したファイル:")
            for fname, error in failed_files:
                print(f"   - {fname}: {error}")
        
        return True

    except Exception as e:
        print(f"❌ クリーンアップ中にエラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

