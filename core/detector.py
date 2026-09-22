"""
Thunderbird Profile and Local Folders auto-detector.
Supports Windows, Linux, and macOS.
"""
import os
import sys
import configparser
from typing import Optional, List, Dict


class ThunderbirdDetector:
    """自動探測本地安裝的 Thunderbird Profile 與 Local Folders 目錄"""

    @staticmethod
    def get_thunderbird_base_dir() -> Optional[str]:
        """依據作業系統取得 Thunderbird 設定檔基底路徑"""
        if sys.platform.startswith("win"):
            appdata = os.environ.get("APPDATA")
            if appdata:
                path = os.path.join(appdata, "Thunderbird")
                if os.path.isdir(path):
                    return path
        elif sys.platform.startswith("darwin"):
            home = os.path.expanduser("~")
            path = os.path.join(home, "Library", "Thunderbird")
            if os.path.isdir(path):
                return path
        else:  # Linux / Unix
            home = os.path.expanduser("~")
            # Thunderbird 在 Linux 上可能是 ~/.thunderbird 或 ~/.mozilla-thunderbird
            for candidate in [".thunderbird", ".mozilla-thunderbird"]:
                path = os.path.join(home, candidate)
                if os.path.isdir(path):
                    return path
        return None

    @classmethod
    def list_profiles(cls) -> List[Dict[str, str]]:
        """
        解析 profiles.ini，列出所有找到的 Thunderbird Profile。
        回傳清單，每個項目包含：name, path, is_default, local_folders_path
        """
        base_dir = cls.get_thunderbird_base_dir()
        if not base_dir:
            return []

        ini_path = os.path.join(base_dir, "profiles.ini")
        if not os.path.isfile(ini_path):
            return []

        config = configparser.ConfigParser()
        try:
            config.read(ini_path, encoding="utf-8")
        except Exception:
            try:
                config.read(ini_path, encoding="latin-1")
            except Exception:
                return []

        profiles = []
        default_install_path = None

        # 檢查 [Install*] 區塊 (Thunderbird 68+ 引入)
        for section in config.sections():
            if section.startswith("Install") and "Default" in config[section]:
                default_install_path = config[section]["Default"]
                break

        for section in config.sections():
            if section.startswith("Profile"):
                p_name = config[section].get("Name", "Unknown")
                p_path = config[section].get("Path", "")
                is_relative = config[section].get("IsRelative", "1") == "1"
                is_default_flag = config[section].get("Default", "0") == "1"

                if not p_path:
                    continue

                full_path = os.path.join(base_dir, p_path) if is_relative else p_path

                # 判斷是否為預設 Profile
                is_default = False
                if default_install_path and (p_path == default_install_path or full_path == default_install_path):
                    is_default = True
                elif is_default_flag:
                    is_default = True

                # 計算 Local Folders 路徑
                local_folders = os.path.join(full_path, "Mail", "Local Folders")

                profiles.append({
                    "name": p_name,
                    "profile_path": full_path,
                    "is_default": is_default,
                    "local_folders_path": local_folders,
                    "exists": os.path.isdir(full_path)
                })

        # 排序：預設的排在最前
        profiles.sort(key=lambda x: not x["is_default"])
        return profiles

    @classmethod
    def get_default_local_folders(cls) -> Optional[str]:
        """取得預設 Profile 的 Local Folders 路徑"""
        profiles = cls.list_profiles()
        if not profiles:
            return None
        # 第一個通常是預設 profile
        return profiles[0]["local_folders_path"]
