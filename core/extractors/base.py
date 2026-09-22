"""
Abstract base class for PST extractors.
"""
from abc import ABC, abstractmethod
from typing import Generator, Dict, Any, List, Optional, Tuple
import email.message
from core.sanitizer import Sanitizer


class ExtractedMessage:
    """封裝單封郵件的標準中介資料結構"""

    def __init__(self):
        self.subject: str = ""
        self.sender_name: str = ""
        self.sender_email: str = ""
        self.to_recipients: List[str] = []
        self.cc_recipients: List[str] = []
        self.bcc_recipients: List[str] = []
        self.delivery_time = None
        self.body_plain: str = ""
        self.body_html: str = ""
        self.attachments: List[Dict[str, Any]] = []  # dict with 'filename', 'content', 'content_id'
        self.message_class: str = "IPM.Note"

    def to_mime(self) -> email.message.EmailMessage:
        """將提取的郵件轉換為標準 RFC 822 (MIME) 物件"""
        msg = email.message.EmailMessage()

        # 標題
        msg["Subject"] = self.subject or "(無主旨)"

        # 寄件人
        msg["From"] = Sanitizer.format_email_address(self.sender_name, self.sender_email)

        # 收件人
        if self.to_recipients:
            msg["To"] = ", ".join(self.to_recipients)
        if self.cc_recipients:
            msg["Cc"] = ", ".join(self.cc_recipients)

        # 日期
        msg["Date"] = Sanitizer.format_date(self.delivery_time)

        # 內文處理 (HTML 優先，保有格式；若無 HTML 則純文字)
        if self.body_html:
            msg.set_content(self.body_plain or "")
            msg.add_alternative(self.body_html, subtype="html")
        else:
            msg.set_content(self.body_plain or "")

        # 附件處理
        for att in self.attachments:
            filename = Sanitizer.sanitize_folder_name(att.get("filename", "attachment.bin"))
            content = att.get("content", b"")
            content_id = att.get("content_id")

            if content:
                if content_id:
                    # 內嵌圖片 (CID)
                    msg.add_attachment(
                        content,
                        maintype="image",
                        subtype="octet-stream",
                        filename=filename,
                        cid=f"<{content_id}>"
                    )
                else:
                    # 一般附件
                    msg.add_attachment(
                        content,
                        maintype="application",
                        subtype="octet-stream",
                        filename=filename
                    )

        return msg


class BasePSTExtractor(ABC):
    """PST 解析器基礎類別"""

    def __init__(self, pst_path: str):
        self.pst_path = pst_path

    @abstractmethod
    def open(self):
        """開啟 PST 檔案"""
        pass

    @abstractmethod
    def close(self):
        """關閉 PST 檔案並釋放資源"""
        pass

    @abstractmethod
    def walk_folders(self) -> Generator[Tuple[List[str], Generator[ExtractedMessage, None, None]], None, None]:
        """
        遞迴走訪資料夾樹。
        Yields:
            (folder_hierarchy: List[str], message_generator)
        """
        pass

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
