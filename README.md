# pst2thunderbird 📬

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.8+](https://img.shields.io/badge/Python-3.8%2B-brightgreen.svg)](https://www.python.org/)
[![Platform: Windows | Linux | macOS](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey.svg)]()
[![Status: Beta Testing](https://img.shields.io/badge/Status-Beta%20Testing-orange.svg)]()

> **All-in-One Seamless Outlook PST to Mozilla Thunderbird Mail Migration Tool**  
> Focused on high-speed, lossless transfer of pure emails, attachments, and folder hierarchies directly into Thunderbird Local Folders — zero third-party add-ons required!

> [!CAUTION]
> **⚠️ IMPORTANT BETA TESTING NOTICE / DATA SAFETY GUIDELINES**  
> This project is currently in **Beta Testing**. To safeguard your precious email data, please adhere to the following rules:  
> 1. **Back up your PST file**: Always make a copy of your original `.pst` file to a secure location (e.g., external drive or cloud storage) before performing any migration.  
> 2. **Keep the original file**: After migration, inspect your emails, attachments, and folder structures in Thunderbird thoroughly.  
> 3. **Never delete the original prematurely**: Only delete the original PST file once you are 100% satisfied that all data has been accurately transferred.

---

## 🌐 English Documentation

### 💡 Why this project?

Migrating from Microsoft Outlook to Mozilla Thunderbird is often frustrating due to several major pain points:
1. **Commercial Software Paywalls**: The market is saturated with expensive proprietary tools ($49–$99 USD), whose trial versions often limit exports to just 25 emails.
2. **Fragile & Complex Manual Steps**: Exporting to EML files and then importing via Thunderbird extensions (like `ImportExportTools NG`) is prone to crashes on large folders, garbled character encoding, and corrupted timestamps.
3. **Overcomplicated PST Structures**: PST files contain emails, contacts, calendars, and tasks. Attempting to parse everything simultaneously often leads to database corruption and crashes.

**Core Design Philosophy of `pst2thunderbird`**:
* 🎯 **Focused Scope (80/20 Rule)**: Strictly targets emails (`IPM.Note`), inline HTML/plain text, attachments, and folder trees.
* ⚡ **Direct-Inject Architecture**: Automatically locates your Thunderbird profile and builds standard `.sbd` directory trees and Mbox files directly inside `Mail/Local Folders`.
* 🛡️ **Zero Add-ons Required**: No Thunderbird extension needed. Launch the tool, migrate, open Thunderbird, and your emails are ready with indices rebuilt automatically.

---

### 🏗️ Architecture

```
+------------------+         +-------------------------------+
| Outlook PST File |  --->   | pst2thunderbird Core Engine   |
+------------------+         | 1. Message Filter (IPM.Note)  |
                             | 2. Exchange X.500 DN Cleanup  |
                             | 3. Attachment & MIME Builder  |
                             +---------------+---------------+
                                             |
                                             v (Direct Streaming Write)
                             +-------------------------------+
                             |  Thunderbird Local Folders    |
                             |   - Outlook_Backup.sbd/       |
                             |     ├── Inbox (Mbox)          |
                             |     ├── Inbox.sbd/            |
                             |     └── Sent (Mbox)           |
                             +-------------------------------+
                                             |
                                             v (Auto-Index on Startup)
                             +-------------------------------+
                             | Mozilla Thunderbird Mailbox   |
                             +-------------------------------+
```

---

### ✨ Key Features

- **Single-Window Intuitive GUI**: Simple one-screen interface with PST file browsing, automatic profile detection, and live dual progress bars.
- **Dual-Engine Support**:
  - **MAPI Engine (Windows)**: Uses native Outlook MAPI COM automation for 100% fidelity without compiling external C dependencies.
  - **PFF Engine (Cross-Platform)**: Supports standalone PST parsing in environments without Outlook installed.
- **Lossless Folder Hierarchy**: Mirrors complex multi-level folder trees while sanitizing forbidden characters (`\ / : * ? " < > |`, `CON`, `NUL`, etc.).
- **Low Memory Footprint (Streaming)**: Streams messages individually to disk, maintaining a steady memory usage (~100MB) even with 30GB+ PST files.
- **Process Guard**: Automatically verifies that `thunderbird.exe` is closed before writing to prevent file locks and cache conflicts.
- **Dual Modes (GUI & CLI)**: Built-in desktop GUI and headless CLI for batch automation.

---

### 🚀 Quick Start

#### 1. Requirements & Setup

Requires **Python 3.8+**. On Windows with Outlook installed, `pywin32` is recommended:

```bash
git clone https://github.com/mkjohnny1003/pst2thunderbird.git
cd pst2thunderbird

pip install -r requirements.txt
```

#### 2. Launch Desktop GUI

```bash
python main.py
# Or simply double-click start_gui.bat on Windows
```

1. Click **Browse** to select your `.pst` file.
2. Verify the automatically detected Thunderbird `Local Folders` path.
3. Ensure the status banner indicates Thunderbird is closed.
4. Click **Start Seamless Import**!

#### 3. Command-Line Interface (CLI)

```bash
# Auto-detect target and run
python cli.py -i C:\path\to\archive.pst

# Specify custom target directory and custom root folder name
python cli.py -i archive.pst -o "C:\Users\User\AppData\Roaming\Thunderbird\Profiles\xxx.default\Mail\Local Folders" -n "Old_Emails_2024"

# List all detected Thunderbird profiles
python cli.py --list-profiles
```

---

### 📦 Build Standalone Windows EXE

To create a standalone `.exe` for distribution without requiring Python:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name pst2thunderbird main.py
```
The output executable will be located in `dist/pst2thunderbird.exe`.

---

## ☕ Sponsor / Donate

If this tool saved you time and hassle, consider buying me a coffee!

[![Donate via PayPal](https://img.shields.io/badge/Donate-PayPal-blue.svg?logo=paypal)](https://www.paypal.me/mkjohnny1003)

**PayPal**: [https://www.paypal.me/mkjohnny1003](https://www.paypal.me/mkjohnny1003)

Your support helps keep this project free, open-source, and actively maintained. Thank you! 🙏

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).

**Author**: [mkjohnny1003](https://github.com/mkjohnny1003)

---
---

## 🌐 中文說明 (Traditional Chinese)

### 💡 為什麼需要這個專案？

當使用者想從 Microsoft Outlook 遷移至 Mozilla Thunderbird 時，通常會遇到以下困境：
1. **商業軟體陷阱**：市面上充斥著收費高昂（$49~$99 美金）的封閉軟體，試用版甚至限制只能轉換 25 封郵件。
2. **手動流程繁瑣且脆弱**：透過 EML 中繼轉出，再到 Thunderbird 安裝 `ImportExportTools NG` 擴充套件，不僅步驟繁雜，且常發生大資料夾匯入當機、中文目錄亂碼、日期時間錯亂等問題。
3. **PST 結構過度複雜**：PST 內夾雜聯絡人、行事曆、待辦事項等多種資料庫結構。若不收斂範圍，往往導致遷移工具極易崩潰。

**`pst2thunderbird` 的核心設計哲學**：
* 🎯 **收斂範疇（80/20 原則）**：只專注處理價值最高的「歷史郵件、附件與階層目錄」，將穩定度做到極致。
* ⚡ **直寫儲存庫（Direct-Inject）**：自動探測本機 Thunderbird 的 Profile，直接建立標準 `.sbd` 樹狀目錄與 Mbox 檔案寫入 `Mail/Local Folders`。
* 🛡️ **開箱即用**：使用者無需在 Thunderbird 安裝任何外掛，只要在工具點擊「開始匯入」，完成後打開 Thunderbird 即見完整郵件與樹狀資料夾！

---

### ✨ 核心特性

- **單一窗口直觀介面**：免繁瑣設定，支援 PST 檔案選取、自動偵測本機 Thunderbird 設定檔、一鍵遷移。
- **雙引擎架構 (Dual-Engine)**：
  - **MAPI 原生引擎 (Windows)**：調用本機 Outlook 原生 MAPI，100% 還原內文排版、附件與時間，零外部 C 編譯負擔。
  - **PFF 引擎 (跨平台)**：支援在無安裝 Outlook 的獨立環境下解析。
- **無損目錄樹鏡像**：完整還原多層子資料夾結構，自動過濾非法檔名保留字（Windows 保留字與特殊符號）。
- **記憶體友善（串流寫入）**：採用 Streaming 寫入機制，即使 30GB+ 的巨型 PST 檔案，記憶體佔用亦維持在 100MB 左右。
- **行程安全守護 (Process Guard)**：自動檢查 Thunderbird 是否處於開啟狀態，防範快取衝突與檔案鎖死。
- **雙模式支援**：同時提供桌面視窗 GUI 與命令列 CLI，方便批次自動化運作。

---

### 🚀 快速開始

#### 1. 安裝環境需求
支援 **Python 3.8+**。在 Windows 且有安裝 Outlook 的環境下建議安裝 `pywin32`：

```bash
git clone https://github.com/mkjohnny1003/pst2thunderbird.git
cd pst2thunderbird

pip install -r requirements.txt
```

#### 2. 啟動桌面視窗 (GUI)
```bash
python main.py
# 或在 Windows 上直接雙擊 start_gui.bat
```

#### 3. 命令列模式 (CLI)
```bash
# 自動偵測目標並執行
python cli.py -i C:\path\to\archive.pst

# 指定自訂目標目錄與根資料夾名稱
python cli.py -i archive.pst -o "C:\path\to\Local Folders" -n "2024備份"
```

---

### ☕ 贊助支持

如果這個工具為您節省了寶貴的時間與金錢，歡迎請我喝杯咖啡！

[![Donate via PayPal](https://img.shields.io/badge/Donate-PayPal-blue.svg?logo=paypal)](https://www.paypal.me/mkjohnny1003)

**PayPal**: [https://www.paypal.me/mkjohnny1003](https://www.paypal.me/mkjohnny1003)
