"""
Single-window modern GUI for PST to Thunderbird Migration.
Uses standard library tkinter/ttk with modern styling and non-blocking background threading.
"""
import os
import sys
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from core.detector import ThunderbirdDetector
from core.process_guard import ProcessGuard
from core.migrator import PSTMigrator, MigrationProgress


class MainWindow:
    """單一視窗 GUI 主介面"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PST to Thunderbird 郵件一體化遷移工具 v1.0.0")
        self.root.geometry("720x650")
        self.root.minsize(640, 580)

        # 啟用 Windows High DPI
        try:
            if sys.platform.startswith("win"):
                from ctypes import windll
                windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

        self.migrator = None
        self.worker_thread = None

        self._init_styles()
        self._build_ui()
        self._refresh_tb_detection()
        self._check_process_status()

    def _init_styles(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 14, "bold"))
        style.configure("Subtitle.TLabel", font=("Segoe UI", 9), foreground="#555555")
        style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))

    def _build_ui(self):
        main_frame = ttk.Frame(self.root, padding="16")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 頂部標題區
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 12))
        ttk.Label(header_frame, text="Outlook PST 轉 Thunderbird 郵件無痛遷移", style="Title.TLabel").pack(anchor="w")
        ttk.Label(header_frame, text="專注於純郵件、附件與目錄樹無損轉移，直接寫入 Thunderbird 本地資料夾，免裝任何外掛。", style="Subtitle.TLabel").pack(anchor="w")

        # 1. PST 來源檔案設定區
        src_group = ttk.LabelFrame(main_frame, text=" 1. PST 來源檔案 ", padding="10")
        src_group.pack(fill=tk.X, pady=6)

        src_box = ttk.Frame(src_group)
        src_box.pack(fill=tk.X)
        self.pst_path_var = tk.StringVar()
        pst_entry = ttk.Entry(src_box, textvariable=self.pst_path_var, font=("Segoe UI", 9))
        pst_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(src_box, text="瀏覽 (PST)...", command=self._browse_pst).pack(side=tk.RIGHT)

        # 2. Thunderbird 目標位置設定區
        dst_group = ttk.LabelFrame(main_frame, text=" 2. Thunderbird 目標資料夾 ", padding="10")
        dst_group.pack(fill=tk.X, pady=6)

        tb_box = ttk.Frame(dst_group)
        tb_box.pack(fill=tk.X, pady=(0, 6))
        self.tb_path_var = tk.StringVar()
        tb_entry = ttk.Entry(tb_box, textvariable=self.tb_path_var, font=("Segoe UI", 9))
        tb_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(tb_box, text="重新偵測", command=self._refresh_tb_detection).pack(side=tk.RIGHT, padx=(0, 4))
        ttk.Button(tb_box, text="手動選擇...", command=self._browse_tb_folder).pack(side=tk.RIGHT)

        name_box = ttk.Frame(dst_group)
        name_box.pack(fill=tk.X, pady=(4, 0))
        ttk.Label(name_box, text="匯入根資料夾名稱:").pack(side=tk.LEFT, padx=(0, 8))
        today_str = datetime.now().strftime("%Y%m%d")
        self.folder_name_var = tk.StringVar(value=f"Outlook_備份_{today_str}")
        ttk.Entry(name_box, textvariable=self.folder_name_var, width=30).pack(side=tk.LEFT)

        # 3. 安全與環境檢查橫幅
        self.status_banner = ttk.Label(main_frame, text="正在檢查系統環境...", foreground="#333333", font=("Segoe UI", 9, "italic"))
        self.status_banner.pack(anchor="w", pady=(2, 6))

        # 4. 執行與進度區
        progress_group = ttk.LabelFrame(main_frame, text=" 3. 遷移進度與日誌 ", padding="10")
        progress_group.pack(fill=tk.BOTH, expand=True, pady=6)

        # 進度條與計數
        info_box = ttk.Frame(progress_group)
        info_box.pack(fill=tk.X, pady=(0, 4))
        self.folder_label = ttk.Label(info_box, text="準備就緒", font=("Segoe UI", 9))
        self.folder_label.pack(side=tk.LEFT)
        self.count_label = ttk.Label(info_box, text="0 封郵件", font=("Segoe UI", 9, "bold"))
        self.count_label.pack(side=tk.RIGHT)

        self.progress_bar = ttk.Progressbar(progress_group, mode="indeterminate")
        self.progress_bar.pack(fill=tk.X, pady=(0, 8))

        # 即時日誌文字框
        log_box = ttk.Frame(progress_group)
        log_box.pack(fill=tk.BOTH, expand=True)
        self.log_text = tk.Text(log_box, height=8, wrap="word", font=("Consolas", 8), bg="#f8f9fa")
        scrollbar = ttk.Scrollbar(log_box, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 5. 底部按鈕區
        action_box = ttk.Frame(main_frame)
        action_box.pack(fill=tk.X, pady=(10, 0))

        self.start_btn = ttk.Button(
            action_box,
            text="開始一鍵無痛匯入",
            style="Primary.TButton",
            command=self._start_migration
        )
        self.start_btn.pack(side=tk.RIGHT, padx=(6, 0))

        self.cancel_btn = ttk.Button(
            action_box,
            text="取消",
            state=tk.DISABLED,
            command=self._cancel_migration
        )
        self.cancel_btn.pack(side=tk.RIGHT)

    def _browse_pst(self):
        filename = filedialog.askopenfilename(
            title="選擇 Outlook PST 檔案",
            filetypes=[("Outlook PST 檔案", "*.pst"), ("所有檔案", "*.*")]
        )
        if filename:
            self.pst_path_var.set(filename)
            self._log(f"已選取來源 PST: {filename}")

    def _browse_tb_folder(self):
        folder = filedialog.askdirectory(title="選擇 Thunderbird 的 Local Folders 目錄")
        if folder:
            self.tb_path_var.set(folder)
            self._log(f"自訂目標目錄: {folder}")

    def _refresh_tb_detection(self):
        detected = ThunderbirdDetector.get_default_local_folders()
        if detected:
            self.tb_path_var.set(detected)
            self._log(f"自動偵測到 Thunderbird Local Folders:\n  -> {detected}")
        else:
            self._log("未自動偵測到 Thunderbird 設定檔，請確認已安裝或手動指定路徑。")

    def _check_process_status(self):
        is_running, msg = ProcessGuard.is_thunderbird_running()
        if is_running:
            self.status_banner.config(
                text="⚠️ 警告：Thunderbird 正在執行中！匯入前請務必完全關閉 Thunderbird。",
                foreground="#d9534f"
            )
        else:
            self.status_banner.config(
                text="✔ 檢查通過：Thunderbird 未執行，可以安全寫入本地資料夾。",
                foreground="#28a745"
            )

    def _log(self, text: str):
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)

    def _start_migration(self):
        pst_path = self.pst_path_var.get().strip()
        tb_path = self.tb_path_var.get().strip()
        folder_name = self.folder_name_var.get().strip()

        if not pst_path or not os.path.isfile(pst_path):
            messagebox.showerror("錯誤", "請先選擇有效的 PST 檔案！")
            return

        if not tb_path or not os.path.isdir(tb_path):
            messagebox.showerror("錯誤", "找不到有效的 Thunderbird 目標資料夾！")
            return

        self._check_process_status()
        is_running, _ = ProcessGuard.is_thunderbird_running()
        if is_running:
            if not messagebox.askyesno("警告", "偵測到 Thunderbird 正在執行！\n強烈建議先關閉它以防快取覆蓋。是否仍要繼續？"):
                return

        self.start_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.progress_bar.start(10)
        self._log(">>> 遷移任務已啟動...")

        self.migrator = PSTMigrator(
            pst_path=pst_path,
            tb_local_folders_path=tb_path,
            root_folder_name=folder_name,
            preferred_engine="auto"
        )

        self.worker_thread = threading.Thread(target=self._run_migration_worker, daemon=True)
        self.worker_thread.start()

    def _run_migration_worker(self):
        try:
            progress = self.migrator.run(progress_callback=self._on_progress_update)
            self.root.after(0, self._on_migration_success, progress)
        except Exception as e:
            self.root.after(0, self._on_migration_error, str(e))

    def _on_progress_update(self, progress: MigrationProgress):
        # 跨執行緒安全更新 UI
        def update():
            self.folder_label.config(text=f"處理中: {progress.current_folder}")
            self.count_label.config(text=f"{progress.total_emails} 封郵件")
            self._log(f"[{progress.current_folder}] 已寫入 {progress.total_emails} 封")
        self.root.after(0, update)

    def _on_migration_success(self, progress: MigrationProgress):
        self.progress_bar.stop()
        self.start_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.DISABLED)
        self.folder_label.config(text="匯入完成！")
        self._log(f"🎉 遷移成功！共匯入 {progress.total_folders} 個資料夾，{progress.total_emails} 封郵件。")
        messagebox.showinfo(
            "遷移完成",
            f"成功完成郵件遷移！\n共匯入 {progress.total_emails} 封郵件。\n\n現在您可以打開 Thunderbird，在「本地資料夾 (Local Folders)」下即可查看所有郵件！"
        )

    def _on_migration_error(self, err_msg: str):
        self.progress_bar.stop()
        self.start_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.DISABLED)
        self.folder_label.config(text="發生錯誤")
        self._log(f"❌ 錯誤: {err_msg}")
        messagebox.showerror("遷移失敗", f"執行過程中發生錯誤：\n{err_msg}")

    def _cancel_migration(self):
        if self.migrator:
            self.migrator.cancel()
            self._log("使用者請求取消中斷...")
            self.cancel_btn.config(state=tk.DISABLED)


def run_gui():
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()
