"""
Migration orchestrator.
Coordinates PST extraction, MIME transformation, and Thunderbird Mbox injection.
"""
import sys
from typing import Callable, Optional
from core.extractors.base import BasePSTExtractor
from core.extractors.mapi_extractor import MAPIExtractor
from core.extractors.pff_extractor import PFFExtractor
from core.mbox_writer import ThunderbirdMboxWriter
from core.process_guard import ProcessGuard


class MigrationProgress:
    def __init__(self):
        self.current_folder: str = ""
        self.total_emails: int = 0
        self.total_folders: int = 0
        self.is_finished: bool = False
        self.error_message: Optional[str] = None


class PSTMigrator:
    """協調提取器與寫入器的遷移核心類別"""

    def __init__(
        self,
        pst_path: str,
        tb_local_folders_path: str,
        root_folder_name: str = "Outlook_Import",
        preferred_engine: str = "auto"  # 'auto', 'mapi', 'pff'
    ):
        self.pst_path = pst_path
        self.tb_local_folders_path = tb_local_folders_path
        self.root_folder_name = root_folder_name
        self.preferred_engine = preferred_engine
        self._cancel_requested = False

    def select_extractor(self) -> BasePSTExtractor:
        """依據作業系統與安裝環境選擇最適當的提取引擎"""
        if self.preferred_engine == "mapi":
            return MAPIExtractor(self.pst_path)
        elif self.preferred_engine == "pff":
            return PFFExtractor(self.pst_path)

        # auto 模式：在 Windows 下優先嘗試 MAPI，若不可用再 fallback 到 PFF
        if sys.platform.startswith("win"):
            try:
                import win32com.client
                return MAPIExtractor(self.pst_path)
            except ImportError:
                pass

        return PFFExtractor(self.pst_path)

    def cancel(self):
        """請求中斷遷移作業"""
        self._cancel_requested = True

    def run(self, progress_callback: Optional[Callable[[MigrationProgress], None]] = None) -> MigrationProgress:
        """
        執行遷移程序
        progress_callback: 接收 MigrationProgress 物件的回調函式
        """
        progress = MigrationProgress()
        self._cancel_requested = False

        # 1. 檢查 Thunderbird 行程保護
        is_running, msg = ProcessGuard.is_thunderbird_running()
        if is_running:
            raise RuntimeError(f"無法執行匯入：{msg} 請先完全關閉 Thunderbird，以防資料快取覆蓋。")

        writer = ThunderbirdMboxWriter(self.tb_local_folders_path, self.root_folder_name)
        extractor = self.select_extractor()

        try:
            with extractor:
                for folder_path, message_generator in extractor.walk_folders():
                    if self._cancel_requested:
                        break

                    folder_display = " / ".join(folder_path)
                    progress.current_folder = folder_display
                    progress.total_folders += 1

                    if progress_callback:
                        progress_callback(progress)

                    mbox = writer.open_folder_mbox(folder_path)
                    try:
                        for ext_msg in message_generator:
                            if self._cancel_requested:
                                break

                            mime_msg = ext_msg.to_mime()
                            writer.write_message(mbox, mime_msg)
                            progress.total_emails += 1

                            if progress_callback:
                                progress_callback(progress)
                    finally:
                        writer.close_folder_mbox(mbox)

            progress.is_finished = True
        except Exception as e:
            progress.error_message = str(e)
            raise e

        return progress
