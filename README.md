# pst2thunderbird 📬

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python 3.8+](https://img.shields.io/badge/Python-3.8%2B-brightgreen.svg)](https://www.python.org/)
[![Platform: Windows | Linux | macOS](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey.svg)]()
[![Status: Beta Testing](https://img.shields.io/badge/Status-Beta%20Testing-orange.svg)]()

> **Outlook PST 轉 Mozilla Thunderbird 郵件一體化無痛遷移工具**  
> 專注於「純郵件、附件與目錄樹」的極速轉移，直接寫入 Thunderbird 本地資料夾，免裝任何第三方外掛！

> [!CAUTION]
> **本專案目前處於 Beta 測試階段，尚未經過大規模實戰驗證。**
>
> 使用前請務必遵守以下安全守則：
> 1. **備份您的 PST 檔案**：在執行遷移前，請先將原始 `.pst` 檔案複製一份到安全的位置（例如外接硬碟或雲端）。
> 2. **保留原始檔案**：遷移完成後，請先在 Thunderbird 中逐一確認郵件內容、附件與資料夾結構是否完整正確。
> 3. **確認無誤後才刪除**：只有在您 100% 確認所有郵件都已正確匯入後，才可考慮刪除原始 PST 檔案，避免造成不可挽回的資料遺失。
>
> 歡迎回報任何問題至 [Issues](https://github.com/mkjohnny1003/pst2thunderbird/issues)！

---

## 💡 為什麼需要這個專案？(Why this project?)

當使用者想從 Microsoft Outlook 遷移至 Mozilla Thunderbird 時，通常會遇到以下困境：
1. **商業軟體陷阱**：市面上充斥著收費高昂（$49~$99 美金）的封閉軟體，試用版甚至限制只能轉換 25 封郵件。
2. **手動流程繁瑣且脆弱**：透過 EML 中繼轉出，再到 Thunderbird 安裝 `ImportExportTools NG` 擴充套件，不僅步驟繁雜，且常發生大資料夾匯入當機、中文目錄亂碼、日期時間錯亂等問題。
3. **PST 結構過度複雜**：PST 內夾雜聯絡人、行事曆、待辦事項等多種資料庫結構。若不收斂範圍，往往導致遷移工具極易崩潰。

**`pst2thunderbird` 的核心設計哲學**：
* 🎯 **收斂範疇（80/20 原則）**：只專注處理價值最高的「歷史郵件、附件與階層目錄」，將穩定度做到極致。
* ⚡ **直寫儲存庫（Direct-Inject）**：自動探測本機 Thunderbird 的 Profile，直接建立標準 `.sbd` 樹狀目錄與 Mbox 檔案寫入 `Mail/Local Folders`。
* 🛡️ **開箱即用**：使用者無需在 Thunderbird 安裝任何外掛，只要在工具點擊「開始匯入」，完成後打開 Thunderbird 即見完整郵件與樹狀資料夾！

---

## 🏗️ 運作架構 (Architecture)

```
+------------------+         +----------------------------+
|  Outlook PST 檔  |  --->   | pst2thunderbird 核心引擎   |
+------------------+         | 1. 訊息類別過濾 (IPM.Note) |
                             | 2. Exchange X.500 地址清洗 |
                             | 3. 附件與 HTML 內文重組    |
                             +--------------+-------------+
                                            |
                                            v (串流直接寫入)
                             +----------------------------+
                             | Thunderbird Local Folders  |
                             |  - Outlook_備份.sbd/       |
                             |    ├── 收件匣 (Mbox)       |
                             |    ├── 收件匣.sbd/         |
                             |    └── 寄件備份 (Mbox)     |
                             +----------------------------+
                                            |
                                            v (啟動自動重建索引)
                             +----------------------------+
                             | Mozilla Thunderbird 介面   |
                             +----------------------------+
```

---

## ✨ 核心特性

- **單一窗口直觀介面**：免繁瑣設定，支援 PST 檔案選取、自動偵測本機 Thunderbird 設定檔、一鍵遷移。
- **雙引擎架構 (Dual-Engine)**：
  - **MAPI 原生引擎 (Windows)**：調用本機 Outlook 原生 MAPI，100% 還原內文排版、附件與時間，零外部 C 編譯負擔。
  - **PFF 引擎 (跨平台)**：支援在無安裝 Outlook 的獨立環境下解析。
- **無損目錄樹鏡像**：完整還原多層子資料夾結構，自動過濾非法檔名保留字（Windows 保留字與特殊符號）。
- **記憶體友善（串流寫入）**：採用 Streaming 寫入機制，即使 30GB+ 的巨型 PST 檔案，記憶體佔用亦維持在 100MB 左右。
- **行程安全守護 (Process Guard)**：自動檢查 Thunderbird 是否處於開啟狀態，防範快取衝突與檔案鎖死。
- **雙模式支援**：同時提供桌面視窗 GUI 與命令列 CLI，方便批次自動化運作。

---

## 🚀 快速開始 (Quick Start)

### 1. 安裝環境需求

本工具支援 Python 3.8+。若在 Windows 環境且已安裝 Outlook，建議安裝 `pywin32` 獲得最佳體驗：

```bash
# 下載專案
git clone https://github.com/mkjohnny1003/pst2thunderbird.git
cd pst2thunderbird

# 安裝依賴 (Windows)
pip install -r requirements.txt
```

### 2. 啟動桌面視窗 (GUI)

直接執行：
```bash
python main.py
```
視窗啟動後：
1. 點擊「瀏覽」選擇你的 `.pst` 檔案。
2. 系統會自動填入本機偵測到的 Thunderbird `Local Folders` 路徑。
3. 確保狀態列顯示「Thunderbird 未執行」。
4. 點擊 **「開始一鍵無痛匯入」**！

### 3. 命令列模式 (CLI)

適用於伺服器管理員或批次作業：

```bash
# 自動偵測目標並執行
python cli.py -i C:\path\to\archive.pst

# 指定目標資料夾名稱與自訂目標目錄
python cli.py -i archive.pst -o "C:\Users\User\AppData\Roaming\Thunderbird\Profiles\xxx.default\Mail\Local Folders" -n "2024年舊信"

# 列出本機偵測到的所有 Thunderbird 設定檔
python cli.py --list-profiles
```

---

## 📦 打包成獨立 Windows EXE 執行檔

若希望讓沒有 Python 環境的同仁或使用者雙擊即用，可透過 PyInstaller 一鍵打包：

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name pst2thunderbird main.py
```
打包完成後的單一執行檔將位於 `dist/pst2thunderbird.exe`。

---

## 🧪 單元測試

本專案附帶完整單元測試，涵蓋路徑探測、命名與標頭清洗、Mbox 生成器：

```bash
python -m unittest discover tests
```

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
