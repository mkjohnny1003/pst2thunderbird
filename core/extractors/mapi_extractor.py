"""
Windows native MAPI extractor for PST files using pywin32 (Outlook COM Automation).
High fidelity, supports HTML, attachments, and native Exchange address resolution.
"""
import os
import tempfile
from typing import Generator, List, Tuple
from core.extractors.base import BasePSTExtractor, ExtractedMessage


class MAPIExtractor(BasePSTExtractor):
    """使用 Windows 原生 Outlook MAPI 引擎讀取 PST 檔案"""

    def __init__(self, pst_path: str):
        super().__init__(pst_path)
        self.abs_path = os.path.abspath(pst_path)
        self.outlook_app = None
        self.namespace = None
        self.target_store = None
        self.root_folder = None

    def open(self):
        try:
            import win32com.client
        except ImportError:
            raise RuntimeError(
                "未安裝 pywin32 套件，請執行 'pip install pywin32' 或使用 PFF 提取引擎。"
            )

        try:
            # 建立 Outlook MAPI Namespace
            self.outlook_app = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook_app.GetNamespace("MAPI")
            
            # 掛載 PST 檔案
            self.namespace.AddStore(self.abs_path)
            
            # 依檔案路徑尋找剛掛載的 Store
            for store in self.namespace.Stores:
                try:
                    if store.FilePath and os.path.abspath(store.FilePath).lower() == self.abs_path.lower():
                        self.target_store = store
                        self.root_folder = store.GetRootFolder()
                        break
                except Exception:
                    continue

            if not self.root_folder:
                raise RuntimeError(f"無法在 Outlook 中成功定位掛載的 PST 檔案: {self.abs_path}")

        except Exception as e:
            raise RuntimeError(f"透過 Outlook MAPI 開啟 PST 失敗: {e}")

    def close(self):
        """卸載 PST 避免檔案被鎖死"""
        if self.namespace and self.root_folder:
            try:
                self.namespace.RemoveStore(self.root_folder)
            except Exception:
                pass
        self.target_store = None
        self.root_folder = None
        self.namespace = None
        self.outlook_app = None

    def walk_folders(self) -> Generator[Tuple[List[str], Generator[ExtractedMessage, None, None]], None, None]:
        if not self.root_folder:
            return

        yield from self._traverse_folder(self.root_folder, [])

    def _traverse_folder(self, folder, current_path: List[str]):
        name = folder.Name
        # 排除微軟頂層重複的根資料夾名稱
        folder_path = current_path + [name] if current_path or name != self.target_store.DisplayName else current_path

        # 提供該資料夾的郵件生成器
        yield (folder_path or [name], self._extract_messages_from_folder(folder))

        # 遞迴處理子資料夾
        for i in range(1, folder.Folders.Count + 1):
            try:
                sub_folder = folder.Folders.Item(i)
                yield from self._traverse_folder(sub_folder, folder_path)
            except Exception:
                continue

    def _extract_messages_from_folder(self, folder) -> Generator[ExtractedMessage, None, None]:
        items = folder.Items
        count = items.Count
        for i in range(1, count + 1):
            try:
                item = items.Item(i)
                # olMail = 43; 且確認 MessageClass 是郵件
                msg_class = getattr(item, "MessageClass", "")
                item_class = getattr(item, "Class", 0)

                if item_class != 43 and not msg_class.startswith("IPM.Note"):
                    continue  # 忽略非郵件項目 (日曆、聯絡人、待辦等)

                msg = ExtractedMessage()
                msg.message_class = msg_class
                msg.subject = getattr(item, "Subject", "")
                msg.sender_name = getattr(item, "SenderName", "")
                
                # 嘗試讀取 SMTP 地址
                sender_email = ""
                try:
                    sender_email = getattr(item, "SenderEmailAddress", "")
                    # 如果是 Exchange DN，嘗試透過 PropertyAccessor 讀取 PR_SMTP_ADDRESS (0x39FE001E)
                    if "@" not in sender_email and hasattr(item, "PropertyAccessor"):
                        prop = item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
                        if prop and "@" in str(prop):
                            sender_email = str(prop)
                except Exception:
                    pass

                msg.sender_email = sender_email
                msg.delivery_time = getattr(item, "ReceivedTime", None) or getattr(item, "SentOn", None)
                msg.body_plain = getattr(item, "Body", "")
                msg.body_html = getattr(item, "HTMLBody", "")

                # 收件人
                try:
                    to_recips = []
                    cc_recips = []
                    for r_idx in range(1, item.Recipients.Count + 1):
                        recip = item.Recipients.Item(r_idx)
                        r_name = getattr(recip, "Name", "")
                        r_addr = getattr(recip, "Address", "")
                        formatted_addr = f'"{r_name}" <{r_addr}>' if "@" in r_addr else r_name
                        if recip.Type == 1:  # olTo
                            to_recips.append(formatted_addr)
                        elif recip.Type == 2:  # olCC
                            cc_recips.append(formatted_addr)
                    msg.to_recipients = to_recips
                    msg.cc_recipients = cc_recips
                except Exception:
                    pass

                # 附件
                try:
                    att_count = item.Attachments.Count
                    if att_count > 0:
                        for a_idx in range(1, att_count + 1):
                            att = item.Attachments.Item(a_idx)
                            att_name = getattr(att, "FileName", f"att_{a_idx}.bin")
                            
                            # 讀取附件二進位
                            with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
                                tmp_path = tmp_file.name
                            try:
                                att.SaveAsFile(tmp_path)
                                with open(tmp_path, "rb") as f:
                                    att_data = f.read()
                                msg.attachments.append({
                                    "filename": att_name,
                                    "content": att_data,
                                    "content_id": None
                                })
                            finally:
                                if os.path.exists(tmp_path):
                                    try:
                                        os.remove(tmp_path)
                                    except Exception:
                                        pass
                except Exception:
                    pass

                yield msg

            except Exception:
                continue
