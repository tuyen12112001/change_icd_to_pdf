import os
import shutil


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
        
        for filename in os.listdir(output_dir):
            file_path = os.path.join(output_dir, filename)
            
            # ディレクトリはスキップ
            if os.path.isdir(file_path):
                continue
            
            # PDF ファイルまたは Excel ファイルは保持
            if filename.lower().endswith(('.pdf', '.xlsx', '.xls')):
                kept_files.append(filename)
                print(f"✅ 保持: {filename}")
            else:
                # その他のファイルは削除
                try:
                    os.remove(file_path)
                    deleted_files.append(filename)
                    print(f"❌ 削除: {filename}")
                except Exception as e:
                    print(f"⚠️ 削除失敗: {filename} - {str(e)}")

        print(f"\n📊 クリーンアップ完了:")
        print(f"   - 保持したファイル: {len(kept_files)} 件")
        print(f"   - 削除したファイル: {len(deleted_files)} 件")
        
        if deleted_files:
            print(f"\n削除されたファイル:")
            for f in deleted_files:
                print(f"   - {f}")
        
        return True

    except Exception as e:
        print(f"❌ クリーンアップ中にエラーが発生しました: {str(e)}")
        return False
