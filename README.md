# 健康資料管理系統 (Local Health Data Management System)

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Python Version](https://img.shields.io/badge/Python-3.10%2B-brightgreen.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20(Single%20Executable)-lightgrey.svg)]()
[![Privacy](https://img.shields.io/badge/Privacy-100%25%20Air--Gapped-success.svg)]()

一套專為校園、活動或企業團體設計的**輕量級定期健康數據追蹤與管理系統**。系統結合 Python 後端（Flask + openpyxl）與極簡原生前端，提供**完全離線（Air-gapped）、零個資外洩、隨插即用（Zero-Config）**的單一獨立桌面軟體體驗。

---

## 🌟 核心特色

### 1. 絕對離線與個資安全防護 (Air-Gapped & Privacy First)
* **100% 本機閉環運作**：後端嚴格綁定 `127.0.0.1` 迴路，不使用任何外部資料庫或雲端 API，區域網路與外網皆無法連入。
* **零外部資源追蹤**：前端無任何第三方遠端 CDN、字型或追蹤腳本，拔除網路線亦可完整運作。
* **防偷查個資雙重驗證**：若他人嘗試透過「名字」或「暱稱」搜尋，系統立即跳出身分驗證彈窗，強制輸入**學號/工號**核對；輸入過程具備密碼遮罩且徹底禁用瀏覽器歷史紀錄，身分不符即刻拒絕存取。
* **多人使用自動重刷（Auto-Reset）**：填報或修改完成後自動重新整理頁面，徹底銷毀螢幕暫存，確保排隊下一位使用者看不到前人個資。

### 2. 生理指標與體脂率智慧檢核
* **支援性別與體脂率**：針對男女生理常態進行醫學邏輯檢核（如極端異常區間或與 BMI 嚴重矛盾）。
* **Excel 異常高亮警示**：若檢測到異常體脂率數據，寫入 Excel 時該儲存格自動填入**亮黃色背景（Yellow Fill）**，便於主辦人員統計追蹤。
* **原生 BMI 公式連動**：Excel 自動嵌入公式 `=IFERROR(ROUND(...), 1)`，數值即時連動。
* **自動週次排序**：補登歷史資料時，系統依週次從小到大（第1週 $\to$ 第2週...）自然排列，維護時間軸可觀性。

### 3. 直覺的人性化填報體驗
* **場域動態適配**：首次啟動時可自選為「學校單位（學號）」或「公司企業（工號）」。
* **擇二自動補齊**：初次填報完整建立；之後每週僅需「名字/暱稱擇一」＋「學號/工號擇一」，系統自動帶出其餘資訊。
* **市售血壓機直覺排版**：上格收縮壓（高壓/SYS）、下格舒張壓（低壓/DIA），避免同學混淆專有名詞。
* **防呆確認機制**：初次建立名單送出前跳出二次核對，並提示姓名與學號一旦鎖定不得隨意更改。

### 4. 隱密管理員特權模式 (Stealth Admin Mode)
* **無痕後門驗證**：前端無任何特權入口提示，管理員於「名字」輸入帳號、「暱稱」輸入密碼並送出（其他欄位皆可留白），即刻解鎖 1 分鐘特權模式。
* **ASCII 加密設定檔**：帳號密碼經位移加密儲存於外部 `admin_config.txt`，管理員可隨時以 `true`/`false` 啟閉特權。
* **全功能解鎖與補登**：特權期間可任意調閱歷史月份、解鎖名字與學號編輯，並能自由指定月份與第 1~5 週補登數據。
* **安全時效與手動登出**：特權有效期限 1 分鐘，閒置無操作或重整頁面自動失效，右上角亦提供【登出特權】按鈕。

### 5. 原生桌面軟體架構 (Single Instance Lock)
* 透過 `pywebview` 包裝為純粹桌面視窗，關閉視窗即徹底終止背景行程，絕不霸佔連線 Port。
* 內建單一實例鎖（Single Instance Lock），重複點擊 `.exe` 不會多開衝突，永久固定使用單一 Port。

---

## 📁 專案目錄結構

```text
health_system/
│
├── data/                    # 執行時自動建立：存放 records.xlsx 與 audit_log.txt
├── static/
│   └── style.css            # 純本地端 UI 樣式
├── templates/
│   └── index.html           # 前端介面、防窺探驗證與開機精靈
│
├── admin_guard.py           # 特權驗證、ASCII 簡易加解密與組態管理
├── excel_manager.py         # Excel 讀寫、體脂率檢核、表頭排序與日誌
├── app.py                   # 後端 API 伺服器 (Flask)
├── main.py                  # 桌面應用程式啟動入口 (pywebview + Port 鎖)
├── requirements.txt         # 開發依賴套件清單
├── .gitignore               # Git 版本控制忽略設定
├── LICENSE                  # MIT 授權條款
└── README.md                # 專案說明文件
