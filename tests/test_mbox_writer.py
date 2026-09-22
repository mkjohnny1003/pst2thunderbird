import os
import shutil
import tempfile
import unittest
import email.message
import mailbox
from core.mbox_writer import ThunderbirdMboxWriter


class TestThunderbirdMboxWriter(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()

    def tearDown(self):
        if os.path.exists(self.test_dir):
            shutil.rmtree(self.test_dir)

    def test_mbox_and_sbd_structure_creation(self):
        writer = ThunderbirdMboxWriter(self.test_dir, root_folder_name="Test_Outlook")

        # 驗證根目錄結構已建立
        self.assertTrue(os.path.exists(writer.root_mbox_path))
        self.assertTrue(os.path.exists(writer.root_sbd_path))

        # 模擬階層: ["Inbox", "SubProject"]
        mbox_path, sbd_path = writer.get_target_paths(["Inbox", "SubProject"])
        self.assertTrue(os.path.exists(mbox_path))
        self.assertTrue(os.path.exists(os.path.dirname(mbox_path)))

        # 寫入測試郵件
        mbox = writer.open_folder_mbox(["Inbox", "SubProject"])
        msg = email.message.EmailMessage()
        msg["Subject"] = "Unit Test Mail"
        msg["From"] = "sender@example.com"
        msg["To"] = "receiver@example.com"
        msg.set_content("Hello from Unit Test!")

        writer.write_message(mbox, msg)
        writer.close_folder_mbox(mbox)

        # 重新讀取 mbox 驗證內容
        read_mbox = mailbox.mbox(mbox_path)
        self.assertEqual(len(read_mbox), 1)
        first_msg = read_mbox[0]
        self.assertEqual(first_msg["Subject"], "Unit Test Mail")
        self.assertEqual(first_msg["From"], "sender@example.com")
        read_mbox.close()


if __name__ == "__main__":
    unittest.main()
