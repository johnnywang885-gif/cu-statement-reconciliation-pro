# 對帳單工具 — AGENTS.md

## 入口

- `對帳單.bat` — CP950/Big5 編碼、CRLF 換行，選單入口（選項 1→`find_anomaly_members.ps1`，選項 2→`filter_by_percent.ps1`，選項 8→重選 CUB.MDB，選項 9→結束）
- 所有 `.ps1` **必須用 `pwsh` (PowerShell 7+)** 執行。Windows PowerShell 5.1 會把 SQL 的 `<` 解析為重新導向而失敗
- `find_anomaly_members.ps1` 和 `loan_merge_check.ps1` 支援 `-CubPath`；其餘腳本固定讀取同目錄的 `CUB.MDB`。所有腳本皆支援 `-CubPassword`
- `find_anomaly_members.ps1`（主引擎）→ 產出 CSV + 寫入 M_對帳明細
- `filter_by_percent.ps1` → 依百分比篩選異常社員、產出 CSV（不回寫 M_對帳明細）
- `balance_check.ps1` → 科目餘額 / 資產負債表 / 損益表
- `audit_register.ps1` → 查核登記管理（List / Import from Excel）
- `loan_merge_check.ps1` → 放款合併（-MergeSource / -SubstantiveReview / -ApplicationReview）
- `manage_nonmail.ps1` → 加入/移除不寄發名單
- `reconcile_reply.ps1` → 對帳回覆登錄查詢
- `pack.ps1` → 打包發布 ZIP

## 必要環境

- **Access Database Engine 2016 (32-bit)** — 才能建立 `DAO.DBEngine.120` COM 物件
- **PowerShell 7+** (`pwsh`)
- `.ps1` 內建 32-bit 檢查，若在 64-bit pwsh 下偵測不到 DAO 會自動用 `SysWOW64\powershell.exe` 重啟自身
- CUB.MDB 密碼透過 `-CubPassword` 傳入或執行時 `Read-Host -AsSecureString` 輸入，**不硬編碼**
- ⚠ 診斷腳本 (`check_lgr.ps1` / `check_lgr2.ps1` / `inspect_cub_schema.ps1`) 含硬編碼測試密碼 — 勿用於正式環境或提交至 git

## 執行

```powershell
pwsh -ExecutionPolicy Bypass -File find_anomaly_members.ps1 -CubPassword "<密碼>"
```

## 異常偵測權重 (AnomalyScore.psm1)

`$script:AnomalyWeights` 定義 17 項，僅保留與挪用相關者：

| 權重 | 項數 | 項目 |
|:----:|:----:|------|
| **5** | 4 | 股金差額、貸款差額、備轉金差額、非社員貸款 (7E) |
| **4** | 3 | 貸款結餘為負 (PREBO<0)、先放後審 (DAT<COUNCILDAT)、休眠戶激活 (7S) |
| **3** | 3 | 關係人放款 (7G)、董監事超限 (7P)、新帳戶爆發 (7U) |
| **2** | 7 | 同地址多戶 (7D)、新入社貸款 (7I)、借據不符 (7M)、審查未入帳 (7N)、重複交易 (7Q)、整數金額 (7T) |
| **1** | 1 | 超過約定還款日期 (7O) |

分級：Score≥10→High, 5-9→Mid, 1-4→Low。

## 已知陷阱

| 問題 | 說明 |
|------|------|
| `SEPN0` vs `SEPNO` | BOROW 欄位是 `SEPNO` (字母O)，非 `SEPN0` (數字0) |
| `DELETE * FROM` | **Access SQL 不支援**，須用 `DELETE FROM` |
| `Get-MemberValue` | `AnomalyScore.psm1` 未匯出此函式，`find_anomaly_members.ps1` 自訂了區域版本 |
| OUTDAT 過濾 | 在社社員用 `(OUTDAT IS NULL OR OUTDAT='')` 判定，非 `TYPE='1'` |
| 權重表 | `lib/AnomalyScore.psm1` 的 `$script:AnomalyWeights`，新增旗標須同時更新此處和 `Get-AnomalyScore` |
| CSV 編碼 | `Export-Csv` 須用 **`-Encoding utf8BOM`**（Excel 需要 BOM 才能正確辨識 UTF-8） |
| SQL 日期 | 中華民國曆 (yyy/MM/dd)，須加 1911 轉西元；用 `Convert-ROCDate` 輔助函式 |
| `Compress-Archive` | `pack.ps1` 切換到暫存目錄再壓縮，避免路徑問題 |
| `"`"$PSCommandPath`""` | 32-bit 重啟時勿用嵌入引號，pwsh 7.3+ 會原樣傳遞；用 `$PSCommandPath` 即可 |
| 借貸方向 | 資產/費用科目借方為正，負債/權益/收入貸方為正；`IIf(DC='C',-1,1)` 或 `IIf(DC='C',1,-1)` 須依科目範圍選用 |
| `BeginTrans` | DELETE+AddNew 批次操作須用 `BeginTrans/CommitTrans` 包覆，失敗 Rollback |

## 驗收檢查

1. CSV 欄位含 `NegativeLoanBalance`、`LoanBeforeApproval`、`DormantActivation`、`RoundAmountTxn`、`NewAccountBurst`
2. 分數分布：High / Mid / Low 三級皆有資料
3. 無退社社員混入（`OUTDAT IS NULL` 過濾）

## 資料表重點

- `M_對帳明細` — `find_anomaly_members.ps1` 會自動建立（若不存在），寫入用 `DELETE FROM` (勿用 `DELETE * FROM`)
- `BOROW` — 放款主檔，含 `OWEBR`(結餘)、`DAT`(撥款日)、`COUNCILDAT`(審核日)、`BRNO`(擔保品號)、`DLY`(滯納天數)、`ALLN`(放款餘額)、**`SEPNO`(字母O)**
- `SER` — 社員主檔，`OUTDAT`(退社日)、`GRPNO`(戶號)、`INDAT`(入社日)、`PREBO`(上期貸款結餘)
- `LGR` — 總帳，含 `DAT` `ACCNO` `MNY` `DC` `USR`，**無 TIME/CONTRA 欄位**
- `k_para` — 系統參數，含 `bDate`、`eDate`（查核起訖）、`OverdueDays`（滯納天數上限）
- 無 `WOFF`/`RECOVER` 表

## 版本控制

- `.gitignore` 排除 `CUB.MDB` 和 `CUB_異常社員_排序_*.csv`（機敏資料與產出檔）
- `pack.ps1` 自動打包發布 ZIP（排除自身、CUB.MDB、CSV、.git）
