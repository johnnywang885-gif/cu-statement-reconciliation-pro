# filter_by_percent.ps1 — 依條件篩選 M_對帳明細
# 用法:
#   .\filter_by_percent.ps1 -Percent 30                     (依比例篩選前 N%)
#   .\filter_by_percent.ps1 -ByScore High                   (依異常分數 High/Mid/Low)
#   .\filter_by_percent.ps1 -ByLoan -Min 100000              (貸款餘額 >= 100000)
#   .\filter_by_percent.ps1 -ByShare -Max 50000              (股金 <= 50000)
#   .\filter_by_percent.ps1 -ByAccount A00001,A00002         (指定帳號清單)
#   .\filter_by_percent.ps1 -Percent 20 -ByScore High        (組合: High 分數取前 20%)
#   .\filter_by_percent.ps1 -CubPassword <密碼>
#
# 若無參數則進入互動模式

param(
    [int]$Percent = 0,
    [string]$ByScore = "",
    [switch]$ByLoan,
    [switch]$ByShare,
    [string]$ByAccount = "",
    [double]$Min = 0,
    [double]$Max = 0,
    [string]$CubPassword = ""
)

Set-Location (Split-Path -Parent $MyInvocation.MyCommand.Path)

$libPath = Join-Path $PSScriptRoot 'lib\AnomalyScore.psm1'
if (Test-Path $libPath) { Import-Module $libPath -Force -DisableNameChecking }

if ([string]::IsNullOrEmpty($CubPassword)) {
    Write-Host "════════════════════════════════════════════════════════════════════════════════════" -ForegroundColor Yellow
    Write-Host "║   CUB.MDB 需要密碼                                          ║" -ForegroundColor Yellow
    Write-Host "════════════════════════════════════════════════════════════════════════════════════" -ForegroundColor Yellow
    $secure = Read-Host -Prompt "  請輸入 CUB.MDB 密碼" -AsSecureString
    $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try { $CubPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($ptr) } finally { [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) | Out-Null }
    Write-Host ""
}

$daoAvailable = $false
try { $null = New-Object -ComObject DAO.DBEngine.120; $daoAvailable = $true } catch {}

if (-not $daoAvailable -and -not $env:DAO_RESTARTED) {
    $ps32 = "$env:SystemRoot\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    if (Test-Path $ps32) {
        Write-Host "  偵測到 32-bit 需求，自動切換..." -ForegroundColor Yellow
        $env:DAO_RESTARTED = '1'
        $argList = @()
        if ($Percent -gt 0) { $argList += '-Percent', $Percent }
        if ($ByScore -ne '') { $argList += '-ByScore', $ByScore }
        if ($ByLoan) { $argList += '-ByLoan' }
        if ($ByShare) { $argList += '-ByShare' }
        if ($ByAccount -ne '') { $argList += '-ByAccount', "`"$ByAccount`"" }
        if ($Min -gt 0) { $argList += '-Min', $Min }
        if ($Max -gt 0) { $argList += '-Max', $Max }
        if ($CubPassword -ne '') { $argList += '-CubPassword', $CubPassword }
        & $ps32 -ExecutionPolicy Bypass -File "`"$PSCommandPath`"" @argList
        exit $LASTEXITCODE
    }
}

if (-not $daoAvailable) {
    Write-Host "  DAO.DBEngine 無法初始化，請安裝 Access Database Engine 2016" -ForegroundColor Red
    exit 1
}

$dbe = New-Object -ComObject DAO.DBEngine.120
$cubPath = Join-Path $PSScriptRoot "CUB.MDB"
if (-not (Test-Path $cubPath)) { Write-Host "  找不到 CUB.MDB" -ForegroundColor Red; exit 1 }
$connectStr = ";PWD=$CubPassword"

$hasMode = ($Percent -gt 0) -or ($ByScore -ne '') -or ($ByLoan) -or ($ByShare) -or ($ByAccount -ne '')
if (-not $hasMode) {
    Write-Host "════════════════════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  篩選模式" -ForegroundColor White
    Write-Host "  1. 依百分比 (Percent)" -ForegroundColor DarkGray
    Write-Host "  2. 依異常分數 (Score)" -ForegroundColor DarkGray
    Write-Host "  3. 依貸款餘額 (Loan)" -ForegroundColor DarkGray
    Write-Host "  4. 依股金餘額 (Share)" -ForegroundColor DarkGray
    Write-Host "  5. 依指定帳號 (Account)" -ForegroundColor DarkGray
    Write-Host "  6. 結束" -ForegroundColor DarkGray
    Write-Host "════════════════════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    $choice = Read-Host "  請選擇模式 (1-6)"
    switch ($choice) {
        "1" {
            $p = Read-Host "  輸入百分比 (1-100)"
            $Percent = [int]$p
        }
        "2" {
            Write-Host "  分數等級: High (>=10), Mid (5-9), Low (1-4)" -ForegroundColor DarkGray
            $ByScore = Read-Host "  輸入等級 (High/Mid/Low)"
        }
        "3" { $ByLoan = $true; $Min = [double](Read-Host "  最低貸款餘額 (0=不限)") }
        "4" { $ByShare = $true; $Max = [double](Read-Host "  最高股金 (0=不限)") }
        "5" {
            $accs = Read-Host "  輸入帳號 (逗號分隔)"
            $ByAccount = $accs
        }
        "6" { exit }
        default { Write-Host "  無效選擇" -ForegroundColor Red; exit }
    }
}

# Step 1: 讀取 M_對帳明細
Write-Host "`n  讀取 CUB.MDB 中..." -ForegroundColor Yellow
$db = $dbe.OpenDatabase($cubPath, $false, $true, $connectStr)

$hasTable = @($db.TableDefs | Where-Object { $_.Name -eq "M_對帳明細" }).Count -gt 0
if (-not $hasTable) {
    Write-Host "  M_對帳明細 不存在，請先執行選項 1" -ForegroundColor Red
    $db.Close(); exit 1
}

$rs = $db.OpenRecordset("SELECT * FROM [M_對帳明細] ORDER BY 帳號")
$allRows = @()
$totalAll = 0
while (-not $rs.EOF) {
    $allRows += [PSCustomObject]@{
        基準日 = if ($rs.Fields("基準日").Value) { [string]$rs.Fields("基準日").Value } else { "" }
        社號  = if ($rs.Fields("社號").Value) { [string]$rs.Fields("社號").Value } else { "" }
        帳號  = if ($rs.Fields("帳號").Value) { [string]$rs.Fields("帳號").Value } else { "" }
        姓名  = if ($rs.Fields("姓名").Value) { [string]$rs.Fields("姓名").Value } else { "" }
        股金  = if ($rs.Fields("股金").Value) { [double]$rs.Fields("股金").Value } else { 0 }
        貸款  = if ($rs.Fields("貸款").Value) { [double]$rs.Fields("貸款").Value } else { 0 }
        備轉金 = if ($rs.Fields("備轉金").Value) { [double]$rs.Fields("備轉金").Value } else { 0 }
        寄發  = if ($rs.Fields("寄發").Value) { $true } else { $false }
        不寄發 = if ($rs.Fields("不寄發").Value) { $true } else { $false }
    }
    $totalAll++
    $rs.MoveNext()
}
$rs.Close()
$db.Close()

Write-Host "  M_對帳明細 共 $totalAll 筆" -ForegroundColor DarkGray

# Step 2: 篩選
$filtered = $allRows

# ByAccount: 指定帳號清單
if ($ByAccount -ne "") {
    $accList = $ByAccount.Split(',', [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() }
    $filtered = $filtered | Where-Object { $_.帳號 -in $accList }
    Write-Host "  帳號篩選: $($accList.Count) 個帳號" -ForegroundColor DarkGray
}

# ByScore: 從 CSV 讀取異常分數
if ($ByScore -ne "") {
    $csvFiles = Get-ChildItem "CUB_異常社員_對帳單排序_*.csv" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
    if ($csvFiles.Count -eq 0) {
        Write-Host "  找不到異常分數 CSV，請先執行選項 1" -ForegroundColor Red
        exit 1
    }
    $scoreCsv = Import-Csv $csvFiles[0].FullName
    $level = $ByScore.ToLower()
    $targetAccs = switch ($level) {
        "high" { $scoreCsv | Where-Object { (Get-ScoreCategory -Score [int]$_.Score) -eq 'High' } | ForEach-Object { $_.AccNo } }
        "mid" { $scoreCsv | Where-Object { (Get-ScoreCategory -Score [int]$_.Score) -eq 'Mid' } | ForEach-Object { $_.AccNo } }
        "low" { $scoreCsv | Where-Object { (Get-ScoreCategory -Score [int]$_.Score) -eq 'Low' } | ForEach-Object { $_.AccNo } }
        default {
            Write-Host "  未知分數等級: $ByScore (請用 High/Mid/Low)" -ForegroundColor Red
            exit 1
        }
    }
    $filtered = $filtered | Where-Object { $_.帳號 -in $targetAccs }
    Write-Host "  分數篩選 ($ByScore): $($targetAccs.Count) 個帳號" -ForegroundColor DarkGray
}

# ByLoan: 貸款餘額 >= Min
if ($ByLoan -and $Min -gt 0) {
    $filtered = $filtered | Where-Object { $_.貸款 -ge $Min }
    Write-Host "  貸款篩選: >= $Min" -ForegroundColor DarkGray
}

# ByShare: 股金 <= Max
if ($ByShare -and $Max -gt 0) {
    $filtered = $filtered | Where-Object { $_.股金 -le $Max }
    Write-Host "  股金篩選: <= $Max" -ForegroundColor DarkGray
}

# Percent: 再從結果取前 N%
if ($Percent -gt 0) {
    $total = $filtered.Count
    if ($total -eq 0) { Write-Host "  篩選結果為空" -ForegroundColor Red; exit }
    $pct = [Math]::Max(1, [int]($total * $Percent / 100))
    $filtered = $filtered | Select-Object -First $pct
    Write-Host "  百分比篩選: 前 $Percent% ($pct 筆)" -ForegroundColor DarkGray
}

Write-Host "  篩選結果: $($filtered.Count) 筆" -ForegroundColor Cyan

# Step 3: 寫回 M_對帳明細（只保留篩選結果）
if ($filtered.Count -gt 0) {
    $db = $dbe.OpenDatabase($cubPath, $false, $false, $connectStr)
    $ws = $dbe.Workspaces(0)
    $ws.BeginTrans()
    try {
        $db.Execute("DELETE FROM [M_對帳明細]")
        $rs = $db.OpenRecordset("SELECT * FROM [M_對帳明細] WHERE False")
        $written = 0
        foreach ($r in $filtered) {
            $rs.AddNew()
            $rs.Fields("基準日").Value = $r.基準日
            $rs.Fields("社號").Value = $r.社號
            $rs.Fields("帳號").Value = $r.帳號
            $rs.Fields("姓名").Value = $r.姓名
            $rs.Fields("股金").Value = $r.股金
            $rs.Fields("貸款").Value = $r.貸款
            $rs.Fields("備轉金").Value = $r.備轉金
            if ($r.寄發) { $rs.Fields("寄發").Value = $true }
            if ($r.不寄發) { $rs.Fields("不寄發").Value = $true }
            $rs.Update()
            $written++
        }
        $rs.Close()
        $ws.CommitTrans()
        $db.Close()
        Write-Host "  已更新 M_對帳明細: $written 筆" -ForegroundColor Green
        Write-Host "  開啟 Access → 對帳單作業 → 列印" -ForegroundColor Green
    }
    catch {
        $ws.Rollback()
        $db.Close()
        Write-Host "  更新失敗，已復原: $_" -ForegroundColor Red
    }
}
else {
    Write-Host "  無符合條件的資料，M_對帳明細未修改" -ForegroundColor Yellow
}

[Runtime.InteropServices.Marshal]::ReleaseComObject($dbe) | Out-Null
Write-Host "`n  Done!" -ForegroundColor Green

