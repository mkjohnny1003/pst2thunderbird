"""
Command-line interface for pst2thunderbird.
Supports automation and headless migration.
"""
import os
import sys
import argparse
from datetime import datetime

from core.detector import ThunderbirdDetector
from core.migrator import PSTMigrator, MigrationProgress


def main():
    parser = argparse.ArgumentParser(
        description="pst2thunderbird: 將 Outlook PST 郵件、附件與目錄樹無痛遷移至 Thunderbird Local Folders。"
    )
    parser.add_argument("-i", "--input", required=True, help="來源 PST 檔案路徑")
    parser.add_argument("-o", "--output", help="Thunderbird Local Folders 目錄路徑 (若不填寫則自動偵測)")
    parser.add_argument("-n", "--name", help="匯入在 Local Folders 底下的根資料夾名稱 (預設 Outlook_備份_YYYYMMDD)")
    parser.add_argument("--engine", choices=["auto", "mapi", "pff"], default="auto", help="指定提取引擎 (auto/mapi/pff)")
    parser.add_argument("--list-profiles", action="store_true", help="列出本機偵測到的所有 Thunderbird 設定檔並退出")

    args = parser.parse_args()

    if args.list_profiles:
        profiles = ThunderbirdDetector.list_profiles()
        print("=== 偵測到的 Thunderbird Profiles ===")
        for p in profiles:
            default_mark = "[預設] " if p["is_default"] else ""
            print(f"- {default_mark}{p['name']}: {p['profile_path']}")
            print(f"  Local Folders -> {p['local_folders_path']}")
        return 0

    pst_path = os.path.abspath(args.input)
    if not os.path.isfile(pst_path):
        print(f"錯誤：找不到 PST 檔案 {pst_path}", file=sys.stderr)
        return 1

    tb_output = args.output
    if not tb_output:
        tb_output = ThunderbirdDetector.get_default_local_folders()
        if not tb_output:
            print("錯誤：無法自動偵測 Thunderbird Local Folders 路徑，請使用 -o 參數手動指定。", file=sys.stderr)
            return 1
        print(f"自動定位 Thunderbird 目標目錄: {tb_output}")

    today_str = datetime.now().strftime("%Y%m%d")
    root_name = args.name or f"Outlook_備份_{today_str}"

    print(f"開始遷移:")
    print(f"  PST 來源: {pst_path}")
    print(f"  目標目錄: {tb_output}")
    print(f"  根資料夾: {root_name}")
    print(f"  使用引擎: {args.engine}")
    print("-" * 50)

    migrator = PSTMigrator(
        pst_path=pst_path,
        tb_local_folders_path=tb_output,
        root_folder_name=root_name,
        preferred_engine=args.engine
    )

    def on_progress(p: MigrationProgress):
        sys.stdout.write(f"\r處理中 [{p.current_folder}] - 已匯入 {p.total_emails} 封郵件...")
        sys.stdout.flush()

    try:
        progress = migrator.run(progress_callback=on_progress)
        print("\n" + "=" * 50)
        print(f"遷移完成！共處理 {progress.total_folders} 個資料夾，成功匯入 {progress.total_emails} 封郵件。")
        print("請開啟 Thunderbird，在「本地資料夾」中即可檢視！")
        return 0
    except Exception as e:
        print(f"\n遷移失敗: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
