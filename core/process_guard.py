"""
Process guard to detect if Thunderbird is currently running.
Prevents writing to Local Folders while Thunderbird has open file handles.
"""
import sys
import subprocess
from typing import Tuple


class ProcessGuard:
    """檢查相關行程狀態，防範快取覆蓋與檔案鎖死"""

    @staticmethod
    def is_thunderbird_running() -> Tuple[bool, str]:
        """
        檢查 thunderbird 是否正在運行。
        回傳 (是否在運行, 提示字串)
        """
        return ProcessGuard._is_process_running(["thunderbird.exe", "thunderbird"])

    @staticmethod
    def is_outlook_running() -> Tuple[bool, str]:
        """檢查 outlook 是否正在運行"""
        return ProcessGuard._is_process_running(["outlook.exe", "outlook"])

    @staticmethod
    def _is_process_running(process_names) -> Tuple[bool, str]:
        """檢查行程名稱是否存在於系統行程列表中"""
        if sys.platform.startswith("win"):
            try:
                # 使用 Windows 內建 tasklist，不依賴額外第三方套件 (如 psutil)
                output = subprocess.check_output(
                    ["tasklist", "/FO", "CSV", "/NH"],
                    creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, "CREATE_NO_WINDOW") else 0,
                    text=True,
                    errors="ignore"
                )
                output_lower = output.lower()
                for name in process_names:
                    if f'"{name.lower()}"' in output_lower or name.lower() in output_lower:
                        return True, f"偵測到 {name} 正在執行中！"
                return False, "未執行"
            except Exception as e:
                return False, f"檢查行程時發生例外: {e}"
        else:
            try:
                output = subprocess.check_output(["ps", "-A"], text=True, errors="ignore").lower()
                for name in process_names:
                    clean_name = name.replace(".exe", "")
                    if clean_name in output:
                        return True, f"偵測到 {clean_name} 正在執行中！"
                return False, "未執行"
            except Exception as e:
                return False, f"檢查行程時發生例外: {e}"
