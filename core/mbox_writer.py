"""
Thunderbird Mbox and .sbd directory structure generator.
Streams emails to prevent memory exhaustion.
"""
import os
import mailbox
import email.message
from typing import Optional
from core.sanitizer import Sanitizer


class ThunderbirdMboxWriter:
    """負責在 Thunderbird Local Folders 底下構建 .sbd 樹狀目錄並寫入 Mbox 郵件"""

    def __init__(self, local_folders_path: str, root_folder_name: str = "Outlook_Import"):
        self.local_folders_path = os.path.abspath(local_folders_path)
        self.clean_root_name = Sanitizer.sanitize_folder_name(root_folder_name, "Outlook_Import")
        
        # 建立根目錄與標識用空 Mbox 檔案
        os.makedirs(self.local_folders_path, exist_ok=True)
        self.root_mbox_path = os.path.join(self.local_folders_path, self.clean_root_name)
        self.root_sbd_path = os.path.join(self.local_folders_path, f"{self.clean_root_name}.sbd")

        self._ensure_file(self.root_mbox_path)
        os.makedirs(self.root_sbd_path, exist_ok=True)

    @staticmethod
    def _ensure_file(file_path: str):
        """確保檔案存在（若不存在則建立空檔）"""
        if not os.path.exists(file_path):
            with open(file_path, "ab"):
                pass

    def get_target_paths(self, folder_hierarchy: list) -> tuple:
        """
        給予階層路徑清單，例如 ["收件匣", "2024專案", "重要"]
        計算出：
        1. 該資料夾的 Mbox 檔案路徑
        2. 該資料夾若有子資料夾時的 .sbd 目錄路徑
        """
        current_dir = self.root_sbd_path

        for i, part in enumerate(folder_hierarchy):
            clean_part = Sanitizer.sanitize_folder_name(part)
            mbox_path = os.path.join(current_dir, clean_part)
            sbd_path = os.path.join(current_dir, f"{clean_part}.sbd")

            self._ensure_file(mbox_path)

            if i == len(folder_hierarchy) - 1:
                # 這是目標資料夾
                return mbox_path, sbd_path
            else:
                # 中繼資料夾，確保其 .sbd 存在
                os.makedirs(sbd_path, exist_ok=True)
                current_dir = sbd_path

        return self.root_mbox_path, self.root_sbd_path

    def open_folder_mbox(self, folder_hierarchy: list) -> mailbox.mbox:
        """打開或建立指定階層的 Mbox 實體"""
        mbox_path, sbd_path = self.get_target_paths(folder_hierarchy)
        mbox = mailbox.mbox(mbox_path)
        mbox.lock()
        return mbox

    def write_message(self, mbox: mailbox.mbox, mime_msg: email.message.EmailMessage):
        """串流寫入單一郵件至 mbox"""
        mbox.add(mime_msg)

    def close_folder_mbox(self, mbox: mailbox.mbox):
        """關閉並釋放 mbox 檔案鎖"""
        try:
            mbox.flush()
        finally:
            mbox.unlock()
            mbox.close()
