# 對帳單工具 — AGENTS.md

## 架構概覽

- **入口**: `對帳單.bat` — 選單式批次檔，呼叫各 PowerShell 腳本
- **核心模組**: `lib/AnomalyScore.psm1` — 異常分數計算引擎
- **資料庫**: `CUB.MDB` (Access, 需密碼, 289MB, `.gitignore` 排除)
- **輸出CSV**: `CUB_異常社員_對帳單排序_*.csv`

## 必要設定

- 須安裝 **Access Database Engine 2016** (32-bit) 才能使用 DAO.DBEngine.120
- 所有 `.ps1` 會自動偵測 32-bit 環境並重啟 `SysWOW64\powershell.exe`
- CUB.MDB 密碼透過 `-CubPassword` 參數傳入或執行時安全提示輸入，**不硬編碼**

## 已知陷阱

| 問題 | 說明 |
|------|------|
| `SEPN0` vs `SEPNO` | BOROW 表的欄位名是 `SEPNO` (字母O)，程式曾誤用 `SEPN0` (數字0) |
| `DELETE * FROM` | **Access SQL 不支援** `DELETE * FROM`，正確是 `DELETE FROM` |
| `Get-MemberValue` | `AnomalyScore.psm1` **未匯出**此函式，各腳本需自行定義區域版本 |

## 執行順序 (選單)

```
1. find_anomaly_members.ps1  → 異常偵測 → 寫入 M_對帳明細 + CSV
2. filter_by_percent.ps1     → 依比例/分數/貸款/帳號篩選 → 更新 M_對帳明細
3. manage_nonmail.ps1        → 不寄發管理
4. reconcile_reply.ps1       → 回覆登錄
5. audit_register.ps1        → 查核登記匯入/查詢
6. balance_check.ps1         → 資產負債表/損益表
7. loan_merge_check.ps1      → 貸款查核合併
```

## 開發指令

```powershell
# 執行異常偵測
powershell -ExecutionPolicy Bypass -File find_anomaly_members.ps1 -CubPassword "密碼"

# 所有腳本都可用 -CubPassword 略過密碼提示
```

## 資料表操作

- `M_對帳明細` — 對帳單明細主表，用 `DELETE FROM` (勿用 `DELETE * FROM`)
- `查核登記表` — 查核記錄，自 `audit_register.ps1` 建立
- `k_對帳單回覆` — 回覆記錄
- `st_科目別` — 科目設定 (需先從 Access 建立)
- `BOROW` — 放款主檔，欄位 `SEPNO` (字母O)
