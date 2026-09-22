import sys, traceback
sys.path.insert(0, ".")
try:
    print("Step 1: importing core.sanitizer...")
    from core.sanitizer import Sanitizer
    print("  OK")

    print("Step 2: importing core.detector...")
    from core.detector import ThunderbirdDetector
    print("  OK")

    print("Step 3: importing core.process_guard...")
    from core.process_guard import ProcessGuard
    print("  OK")

    print("Step 4: importing core.mbox_writer...")
    from core.mbox_writer import ThunderbirdMboxWriter
    print("  OK")

    print("Step 5: importing core.extractors.base...")
    from core.extractors.base import BasePSTExtractor, ExtractedMessage
    print("  OK")

    print("Step 6: importing core.migrator...")
    from core.migrator import PSTMigrator
    print("  OK")

    print("Step 7: importing ui.window...")
    from ui.window import MainWindow, run_gui
    print("  OK")

    print("Step 8: import tkinter...")
    import tkinter as tk
    print("  OK")

    print("\n=== ALL IMPORTS PASSED ===")
except Exception as e:
    print(f"\n!!! FAILED: {e}")
    traceback.print_exc()

input("\nPress Enter to exit...")
