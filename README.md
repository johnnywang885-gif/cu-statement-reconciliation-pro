# account-statement-tool

> 對帳單工具 (進階版) — 儲互社對帳單資料處理與異常社員偵測自動化系統 (Private repository)

## 功能

- **異常社員偵測**：自動篩選、計算每位社員的異常評分（包含股金貸款差異、新入社放款、利率不符等指標）。
- **資料庫稽核登記**：支援社員餘額對帳核對、稽核登記與放款合併檢查。
- **對帳回覆統計與非郵寄管理**：維護非郵寄名單與對帳回覆數據。

## 使用方式

請於 Windows PowerShell 環境下，執行批次檔啟動系統：
```bash
# 雙擊執行或在終端機中執行
.\對帳單.bat
```

## 環境需求

- Windows OS
- Windows PowerShell 5.1+
- Microsoft Access Database Engine 2016 (32-bit)
- 依賴資料庫：`CUB.MDB` (約 289MB 儲互社資料庫)

## 專案結構

```text
/
├── 對帳單.bat             # 系統啟動批次檔 (Big5 編碼)
├── find_anomaly_members.ps1  # 異常社員偵測核心引擎 (自動判定 32-bit 環境)
├── audit_register.ps1     # 稽核登記處理
├── balance_check.ps1      # 社員餘額核對
├── loan_merge_check.ps1   # 放款合併檢查
├── manage_nonmail.ps1     # 非郵寄社員名單管理
└── lib/                   # 共用 PowerShell 函式模組
```

## License

MIT
