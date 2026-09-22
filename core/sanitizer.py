"""
Sanitization utilities for folder names, email addresses, and MIME headers.
"""
import re
from datetime import datetime
import email.utils


class Sanitizer:
    """清理非法檔名、路徑與電子郵件標頭"""

    ILLEGAL_FILENAME_CHARS = re.compile(r'[\\/*?:"<>|]')

    @classmethod
    def sanitize_folder_name(cls, name: str, fallback="Folder") -> str:
        """
        過濾作業系統保留字元，避免在建立 Mbox 檔案與 .sbd 目錄時出錯。
        """
        if not name or not isinstance(name, str):
            return fallback

        # 替換非法字元為底線
        cleaned = cls.ILLEGAL_FILENAME_CHARS.sub("_", name).strip()

        # 移除前後點與空白 (Windows 不允許檔名以點或空格結尾)
        cleaned = cleaned.strip(". ")

        # 檢查 Windows 保留字
        reserved_names = {
            "CON", "PRN", "AUX", "NUL",
            "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
        }
        if cleaned.upper() in reserved_names:
            cleaned = f"{cleaned}_folder"

        return cleaned or fallback

    @classmethod
    def format_email_address(cls, display_name: str, raw_address: str = None) -> str:
        """
        清洗微軟 Exchange 內部地址 (X.500 DN)，格式化為合規的 RFC 5322 地址：
        "Display Name" <username@domain>
        """
        name = (display_name or "").strip().replace('"', "'")
        addr = (raw_address or "").strip()

        # 若是微軟 Exchange 內部 X.500 格式 (例如 /O=COMPANY/OU=...) 或無效 address
        is_exchange_dn = "/O=" in addr.upper() or "/OU=" in addr.upper() or "EX:" in addr.upper()

        if is_exchange_dn or not addr or "@" not in addr:
            # 建立相容的虛擬地址
            alias_source = name or addr or "user"
            safe_alias = re.sub(r'[^a-zA-Z0-9._-]', '_', alias_source).strip('_').lower()
            if not safe_alias:
                safe_alias = "unknown_sender"
            final_email = f"{safe_alias}@exchange.local"
        else:
            final_email = addr

        if name:
            return f'"{name}" <{final_email}>'
        return f"<{final_email}>"

    @classmethod
    def format_date(cls, dt) -> str:
        """轉換日期為標準 RFC 2822 日期字串"""
        if isinstance(dt, datetime):
            try:
                return email.utils.format_datetime(dt)
            except Exception:
                pass
        elif isinstance(dt, str) and dt:
            return dt

        return email.utils.format_datetime(datetime.now())
