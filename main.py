"""
Application entry point.
Launches the Single-window GUI by default, or redirects to CLI if command-line arguments are provided.
"""
import sys
import traceback


def main():
    if len(sys.argv) > 1 and sys.argv[1] not in ["--help", "-h"]:
        # 如果有傳入其他參數，啟動命令列模式
        from cli import main as cli_main
        sys.exit(cli_main())
    else:
        # 預設啟動桌面 GUI 視窗
        from ui.window import run_gui
        run_gui()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("=" * 60)
        print("啟動時發生未預期的錯誤:")
        print("=" * 60)
        traceback.print_exc()
        print("=" * 60)
        input("按 Enter 鍵關閉此視窗...")
        sys.exit(1)
