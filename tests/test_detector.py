import unittest
from core.detector import ThunderbirdDetector


class TestThunderbirdDetector(unittest.TestCase):

    def test_get_base_dir_runs_without_exception(self):
        base_dir = ThunderbirdDetector.get_thunderbird_base_dir()
        if base_dir:
            self.assertIsInstance(base_dir, str)

    def test_list_profiles_returns_list(self):
        profiles = ThunderbirdDetector.list_profiles()
        self.assertIsInstance(profiles, list)
        for p in profiles:
            self.assertIn("name", p)
            self.assertIn("profile_path", p)
            self.assertIn("local_folders_path", p)
            self.assertIn("is_default", p)


if __name__ == "__main__":
    unittest.main()
