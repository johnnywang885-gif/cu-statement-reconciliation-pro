# filter_by_percent.ps1 — 依比例篩選 CSV（簡化版）
# 用法: .\filter_by_percent.ps1 -Percent 30

param([int]$Percent = 0)

Set-Location (Split-Path -Parent $MyInvocation.MyCommand.Path)

# 驗證參數
if ($Percent -lt 1 -or $Percent -gt 100) {
    Write-Host "請指定 -Percent N (1-100)" -ForegroundColor Red
    exit 1
}

# 尋找最新 CSV
$csvFile = Get-ChildItem "CUB_異常社員_對帳單排序_*.csv" `
    -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (-not $csvFile) {
    Write-Host "找不到 CUB_異常社員_對帳單排序_*.csv，請先執行選項 1" -ForegroundColor Red
    exit 1
}

# 讀取 CSV
Write-Host "`n  讀取 $($csvFile.Name) ..." -ForegroundColor Yellow
$allRows = Import-Csv $csvFile.FullName -Encoding utf8
$total = $allRows.Count
Write-Host "  共 $total 筆" -ForegroundColor DarkGray

# 篩選前 N%
$pct = [Math]::Max(1, [int]($total * $Percent / 100))
$filtered = $allRows | Select-Object -First $pct
Write-Host "  篩選: 前 $Percent% ($pct 筆)" -ForegroundColor Cyan

# 輸出
$outFile = "CUB_異常社員_篩選結果.csv"
$filtered | Export-Csv -Path $outFile -Encoding utf8BOM -NoTypeInformation
Write-Host "  已輸出 $($filtered.Count) 筆至 $outFile" -ForegroundColor Green
Write-Host "  Done!" -ForegroundColor Green
