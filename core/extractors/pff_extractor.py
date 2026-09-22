"""
Extractor using pypff (libpff python bindings) for standalone PST parsing without Outlook.
"""
from typing import Generator, List, Tuple
from core.extractors.base import BasePSTExtractor, ExtractedMessage


class PFFExtractor(BasePSTExtractor):
    """使用 pypff 函式庫解析 PST 檔案（不需安裝 Outlook）"""

    def __init__(self, pst_path: str):
        super().__init__(pst_path)
        self.pff_file = None

    def open(self):
        try:
            import pypff
        except ImportError:
            raise RuntimeError(
                "未安裝 pypff 套件。若在 Windows 且有 Outlook，請改用 MAPI 引擎；或安裝 pypff-python。"
            )

        self.pff_file = pypff.file()
        self.pff_file.open(self.pst_path)

    def close(self):
        if self.pff_file:
            try:
                self.pff_file.close()
            except Exception:
                pass
            self.pff_file = None

    def walk_folders(self) -> Generator[Tuple[List[str], Generator[ExtractedMessage, None, None]], None, None]:
        if not self.pff_file:
            return

        root_folder = self.pff_file.get_root_folder()
        yield from self._traverse_folder(root_folder, [])

    def _traverse_folder(self, folder, current_path: List[str]):
        name = folder.get_name() or "Folder"
        folder_path = current_path + [name] if current_path or name != "Top of Outlook data file" else current_path

        yield (folder_path or [name], self._extract_messages(folder))

        num_sub = folder.get_number_of_sub_folders()
        for i in range(num_sub):
            try:
                sub_folder = folder.get_sub_folder(i)
                yield from self._traverse_folder(sub_folder, folder_path)
            except Exception:
                continue

    def _extract_messages(self, folder) -> Generator[ExtractedMessage, None, None]:
        num_msgs = folder.get_number_of_sub_messages()
        for i in range(num_msgs):
            try:
                pff_msg = folder.get_sub_message(i)
                msg_class = pff_msg.get_message_class() or "IPM.Note"

                if not msg_class.startswith("IPM.Note"):
                    continue  # 跳過非郵件項目

                msg = ExtractedMessage()
                msg.message_class = msg_class
                msg.subject = pff_msg.get_subject() or ""
                msg.sender_name = pff_msg.get_sender_name() or ""
                msg.delivery_time = pff_msg.get_delivery_time()

                # 內文
                plain = pff_msg.get_plain_text_body()
                html = pff_msg.get_html_body()
                msg.body_plain = plain.decode("utf-8", errors="replace") if plain else ""
                msg.body_html = html.decode("utf-8", errors="replace") if html else ""

                # 附件
                num_att = pff_msg.get_number_of_attachments()
                for a in range(num_att):
                    try:
                        att = pff_msg.get_attachment(a)
                        size = att.get_size()
                        if size > 0:
                            att_name = att.get_name() or f"attachment_{a}.bin"
                            att_data = att.read_buffer(size)
                            msg.attachments.append({
                                "filename": att_name,
                                "content": att_data,
                                "content_id": None
                            })
                    except Exception:
                        continue

                yield msg
            except Exception:
                continue
