import unittest
from datetime import datetime
from core.sanitizer import Sanitizer


class TestSanitizer(unittest.TestCase):

    def test_sanitize_folder_name_illegal_chars(self):
        raw = 'Project: 2024 / Q1 *Important* <Confidential>?"|'
        cleaned = Sanitizer.sanitize_folder_name(raw)
        self.assertNotIn(":", cleaned)
        self.assertNotIn("/", cleaned)
        self.assertNotIn("*", cleaned)
        self.assertNotIn("<", cleaned)
        self.assertNotIn(">", cleaned)
        self.assertNotIn("?", cleaned)
        self.assertNotIn('"', cleaned)
        self.assertNotIn("|", cleaned)

    def test_sanitize_folder_name_windows_reserved(self):
        self.assertEqual(Sanitizer.sanitize_folder_name("CON"), "CON_folder")
        self.assertEqual(Sanitizer.sanitize_folder_name("prn"), "prn_folder")
        self.assertEqual(Sanitizer.sanitize_folder_name("NUL"), "NUL_folder")
        self.assertEqual(Sanitizer.sanitize_folder_name("COM1"), "COM1_folder")

    def test_sanitize_folder_name_empty_or_dots(self):
        self.assertEqual(Sanitizer.sanitize_folder_name("..."), "Folder")
        self.assertEqual(Sanitizer.sanitize_folder_name("", fallback="FallbackFolder"), "FallbackFolder")

    def test_format_email_address_standard(self):
        res = Sanitizer.format_email_address("Alice Wang", "alice@example.com")
        self.assertEqual(res, '"Alice Wang" <alice@example.com>')

    def test_format_email_address_exchange_dn(self):
        dn = "/O=EXAMPLE/OU=EXCHANGE ADMINISTRATIVE GROUP/CN=RECIPIENTS/CN=ALICE"
        res = Sanitizer.format_email_address("Alice Wang", dn)
        self.assertTrue(res.startswith('"Alice Wang" <'))
        self.assertTrue(res.endswith('@exchange.local>'))

    def test_format_date(self):
        dt = datetime(2026, 9, 23, 10, 0, 0)
        formatted = Sanitizer.format_date(dt)
        self.assertIn("2026", formatted)
        self.assertIn("Sep", formatted)


if __name__ == "__main__":
    unittest.main()
