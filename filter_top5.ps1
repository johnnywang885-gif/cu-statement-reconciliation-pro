# filter_top5.ps1 — 一鍵篩選異常分數前 5%，含股金/貸款/備轉金結餘
# 用法:
#   .\filter_top5.ps1              (預設前 5%)
#   .\filter_top5.ps1 -Percent 10  (指定百分比)
# 前置: 須先執行 find_anomaly_members.ps1 產出 CSV

param([int]$Percent = 5)

Set-Location (Split-Path -Parent $MyInvocation.MyCommand.Path)

if ($Percent -lt 1 -or $Percent -gt 100) {
    Write-Host "百分比須為 1-100 (收到 $Percent)" -ForegroundColor Red
    exit 1
}

# 尋找最新 CSV
$csvFile = Get-ChildItem "CUB_異常社員_對帳單排序_*.csv" `
    -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (-not $csvFile) {
    Write-Host "找不到 CUB_異常社員_對帳單排序_*.csv，請先執行選項 1 (異常偵測)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "讀取 $($csvFile.Name) ..." -ForegroundColor Yellow
$allRows = Import-Csv $csvFile.FullName -Encoding utf8
$total = $allRows.Count

if ($total -eq 0) {
    Write-Host "CSV 無資料" -ForegroundColor Red
    exit 1
}

# 篩選前 N%
$pct = [Math]::Max(1, [int]($total * $Percent / 100))
$filtered = $allRows |
    Sort-Object { [int]$_.Score } -Descending |
    Select-Object -First $pct |
    Sort-Object { $_.AccNo }

Write-Host "  共 $total 筆，篩選前 $Percent% = $pct 筆" -ForegroundColor Cyan

# 選取輸出欄位（含結餘）
$outRows = $filtered | Select-Object `
    AccNo,
    Name1,
    Name2,
    ADDR,
    TEL,
    Score,
    Flags_cn,
    LGR_Share,
    SER_Share,
    DiffShare,
    LGR_Loan,
    SER_Loan,
    DiffLoan,
    LGR_Reserve,
    SER_Reserve,
    DiffReserve,
    BadLoans,
    BadPrincipal,
    BadInterest,
    LoanCount,
    RecentTxn,
    NonMember,
    SameAddrMul,
    RelatedParty,
    NewJoinLoan,
    AuditOverdue,
    ExceedPayDate,
    DirectorOverLimit,
    DuplicateLoan,
    DormantActivation,
    RoundAmountTxn,
    NewAccountBurst,
    NegativeLoanBalance,
    LoanBeforeApproval

# 輸出
$ts = (Get-Date).ToString('yyyyMMdd_HHmmss')
$outFile = "CUB_異常社員_Top${Percent}Pct_$ts.csv"
$outRows | Export-Csv -Path $outFile -Encoding utf8BOM -NoTypeInformation

Write-Host ""
Write-Host "  輸出: $outFile" -ForegroundColor Green
Write-Host ""

# 統計摘要
$high = @($outRows | Where-Object { [int]$_.Score -ge 10 }).Count
$mid  = @($outRows | Where-Object { [int]$_.Score -ge 5 -and [int]$_.Score -lt 10 }).Count
$low  = @($outRows | Where-Object { [int]$_.Score -ge 1 -and [int]$_.Score -lt 5 }).Count

Write-Host "=== Top $Percent% 摘要 ===" -ForegroundColor Cyan
Write-Host "  總篩選: $($outRows.Count) 筆" -ForegroundColor White
Write-Host "  [High] 分數>=10: $high 筆" -ForegroundColor Red
Write-Host "  [Mid]  分數 5-9: $mid 筆" -ForegroundColor Yellow
Write-Host "  [Low]  分數 1-4: $low 筆" -ForegroundColor DarkGray
Write-Host ""

# 前 10 名預覽
Write-Host "=== 前 10 名 ===" -ForegroundColor Cyan
Write-Host ("{0,-4} {1,-6} {2,-10} {3,4}  {4,10}  {5,10}  {6,10}  {7}" -f `
    "排名", "會員", "戶名", "分數", "股金差額", "貸款差額", "備轉差額", "異常說明") -ForegroundColor DarkGray
Write-Host ("-" * 100) -ForegroundColor DarkGray

$rank = 1
foreach ($r in $outRows | Select-Object -First 10) {
    $desc = $r.Flags_cn
    if ($desc.Length -gt 40) { $desc = $desc.Substring(0, 37) + "..." }
    Write-Host ("{0,3}. {1,-6} {2,-10} {3,4}  {4,10:N0}  {5,10:N0}  {6,10:N0}  {7}" -f `
        $rank, $r.AccNo, $r.Name1, $r.Score,
        [long]$r.DiffShare, [long]$r.DiffLoan, [long]$r.DiffReserve,
        $desc)
    $rank++
}

Write-Host ""
Write-Host "  以 Excel 開啟 $outFile 即可檢視完整明細" -ForegroundColor Green
Write-Host ""
