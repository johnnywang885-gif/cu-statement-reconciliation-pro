# balance_check.ps1 — 科目年度結餘產出
# 用法:
#   .\balance_check.ps1 -BalanceSheet                     (產出資產負債表)
#   .\balance_check.ps1 -IncomeStatement                   (產出損益表)
#   .\balance_check.ps1 -ListAccounts                      (列出科目)
#   .\balance_check.ps1 -SelectCategory Assets              (選取資產類科目)
#   .\balance_check.ps1 -CubPassword <密碼>
#
#   類別: Assets, Liabilities, Equity, Incomes, Expenses, All

param(
    [string]$CubPath = "",
    [switch]$BalanceSheet,
    [switch]$IncomeStatement,
    [switch]$ListAccounts,
    [string]$SelectCategory = "",
    [string]$CubPassword = ""
)

Set-Location (Split-Path -Parent $MyInvocation.MyCommand.Path)

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
        if ($CubPath -ne '') { $argList += '-CubPath', $CubPath }
        if ($BalanceSheet) { $argList += '-BalanceSheet' }
        if ($IncomeStatement) { $argList += '-IncomeStatement' }
        if ($ListAccounts) { $argList += '-ListAccounts' }
        if ($SelectCategory) { $argList += '-SelectCategory', $SelectCategory }
        if ($CubPassword -ne '') { $argList += '-CubPassword', $CubPassword }
        & $ps32 -ExecutionPolicy Bypass -File "`"$PSCommandPath`"" @argList
        exit $LASTEXITCODE
    }
}

if (-not $daoAvailable) {
    Write-Host "  DAO.DBEngine 無法初始化，請安裝 Access Database Engine 2016" -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrEmpty($CubPath)) { $CubPath = Join-Path $PSScriptRoot "CUB.MDB" }
$CubPath = [System.IO.Path]::GetFullPath($CubPath)
if (-not (Test-Path $CubPath)) { Write-Host "  找不到 CUB.MDB: $CubPath" -ForegroundColor Red; exit 1 }

$connectStr = ";PWD=$CubPassword"

try {
    $dbe = New-Object -ComObject DAO.DBEngine.120
    $needWrite = ($SelectCategory -ne "")
    $db = $dbe.OpenDatabase($CubPath, $false, (-not $needWrite), $connectStr)

    # 確保 st_科目別 存在
    $hasTable = @($db.TableDefs | Where-Object { $_.Name -eq "st_科目別" }).Count -gt 0
    if (-not $hasTable) {
        Write-Host "  st_科目別 不存在於 CUB.MDB" -ForegroundColor Red
        Write-Host "  請先從 Access 中設定科目表" -ForegroundColor Yellow
        $db.Close(); exit 1
    }

    # 列出科目
    if ($ListAccounts) {
        $rs = $db.OpenRecordset("SELECT * FROM [st_科目別] ORDER BY 科目代號")
        Write-Host "`n=== 科目清單 ===" -ForegroundColor Cyan
        $count = 0
        while (-not $rs.EOF) {
            $code = if ($rs.Fields("科目代號").Value) { $rs.Fields("科目代號").Value } else { "" }
            $name = if ($rs.Fields("科目名稱").Value) { $rs.Fields("科目名稱").Value } else { "" }
            $sel = if ($rs.Fields("Selected").Value) { "V" } else { "" }
            Write-Host ("  {0,-8} {1,-20} [{2}]" -f $code, $name, $sel)
            $count++
            $rs.MoveNext()
        }
        $rs.Close()
        Write-Host "  共 $count 筆" -ForegroundColor Green
    }

    # 選取科目類別
    if ($SelectCategory -ne "") {
        if (-not $needWrite) { $db.Close(); $db = $dbe.OpenDatabase($CubPath, $false, $false, $connectStr) }
        switch ($SelectCategory) {
            "All" { $sql = "UPDATE [st_科目別] SET Selected=True" }
            "Assets" { $sql = "UPDATE [st_科目別] SET Selected=True WHERE 科目代號 < '2'" }
            "Liabilities" { $sql = "UPDATE [st_科目別] SET Selected=True WHERE 科目代號 > '2' AND 科目代號 < '4'" }
            "Equity" { $sql = "UPDATE [st_科目別] SET Selected=True WHERE 科目代號 >= '3' AND 科目代號 < '4'" }
            "Incomes" { $sql = "UPDATE [st_科目別] SET Selected=True WHERE 科目代號 >= '326' AND 科目代號 < '5'" }
            "Expenses" { $sql = "UPDATE [st_科目別] SET Selected=True WHERE 科目代號 > '5'" }
            "None" { $sql = "UPDATE [st_科目別] SET Selected=False" }
            default {
                Write-Host "  未知類別: $SelectCategory" -ForegroundColor Red
                Write-Host "  可用: All, Assets, Liabilities, Equity, Incomes, Expenses, None" -ForegroundColor Yellow
                $db.Close(); exit 1
            }
        }
        $db.Execute($sql)
        Write-Host "  已選取類別: $SelectCategory" -ForegroundColor Green
    }

    # 產出資產負債表
    if ($BalanceSheet) {
        Write-Host "`n=== 資產負債表 ===" -ForegroundColor Cyan
        $rs = $db.OpenRecordset("SELECT 科目代號, 科目名稱 FROM [st_科目別] WHERE Selected=True AND 科目代號 < '4' ORDER BY 科目代號")
        $totalAssets = 0; $totalLiabEquity = 0
        while (-not $rs.EOF) {
            $code = [string]$rs.Fields("科目代號").Value
            $name = [string]$rs.Fields("科目名稱").Value
            # 從 LGR 抓餘額
            $rs2 = $db.OpenRecordset("SELECT Sum(IIf(DC='C',1,-1)*MNY) AS BLN FROM LGR WHERE Left(ACNO,3)='" + $code + "' AND DCHK='Y'")
            $bln = if (-not $rs2.EOF -and $rs2.Fields("BLN").Value) { [double]$rs2.Fields("BLN").Value } else { 0 }
            $rs2.Close()

            if ($code -lt '2') {
                $totalAssets += $bln
                Write-Host ("  資產 {0,-6} {1,-20} {2,12:N0}" -f $code, $name, $bln) -ForegroundColor White
            }
            else {
                $totalLiabEquity += $bln
                Write-Host ("  負債/權益 {0,-6} {1,-20} {2,12:N0}" -f $code, $name, $bln) -ForegroundColor White
            }
            $rs.MoveNext()
        }
        $rs.Close()
        Write-Host ("  ─────────────────────────────────") -ForegroundColor DarkGray
        Write-Host ("  資產總計: {0,12:N0}" -f $totalAssets) -ForegroundColor Green
        Write-Host ("  負債權益總計: {0,12:N0}" -f $totalLiabEquity) -ForegroundColor Green
        Write-Host ("  差額: {0,12:N0}" -f ($totalAssets - $totalLiabEquity)) -ForegroundColor Yellow
    }

    # 產出損益表
    if ($IncomeStatement) {
        Write-Host "`n=== 損益表 ===" -ForegroundColor Cyan
        $rs = $db.OpenRecordset("SELECT 科目代號, 科目名稱 FROM [st_科目別] WHERE Selected=True AND 科目代號 >= '326' ORDER BY 科目代號")
        $totalIncome = 0; $totalExpense = 0
        while (-not $rs.EOF) {
            $code = [string]$rs.Fields("科目代號").Value
            $name = [string]$rs.Fields("科目名稱").Value
            $rs2 = $db.OpenRecordset("SELECT Sum(IIf(DC='C',1,-1)*MNY) AS BLN FROM LGR WHERE Left(ACNO,Len('$code'))='$code' AND DCHK='Y'")
            $bln = if (-not $rs2.EOF -and $rs2.Fields("BLN").Value) { [double]$rs2.Fields("BLN").Value } else { 0 }
            $rs2.Close()

            if ($code -ge '326' -and $code -lt '5') {
                $totalIncome += $bln
                Write-Host ("  收入 {0,-6} {1,-20} {2,12:N0}" -f $code, $name, $bln) -ForegroundColor White
            }
            else {
                $totalExpense += $bln
                Write-Host ("  費用 {0,-6} {1,-20} {2,12:N0}" -f $code, $name, $bln) -ForegroundColor White
            }
            $rs.MoveNext()
        }
        $rs.Close()
        Write-Host ("  ─────────────────────────────────") -ForegroundColor DarkGray
        Write-Host ("  收入總計: {0,12:N0}" -f $totalIncome) -ForegroundColor Green
        Write-Host ("  費用總計: {0,12:N0}" -f $totalExpense) -ForegroundColor Green
        Write-Host ("  本期損益: {0,12:N0}" -f ($totalIncome - $totalExpense)) -ForegroundColor Yellow
    }

    $db.Close()
}
catch {
    Write-Host "  錯誤: $_" -ForegroundColor Red
}
finally {
    if ($dbe) { [Runtime.InteropServices.Marshal]::ReleaseComObject($dbe) | Out-Null }
}

if (-not $BalanceSheet -and -not $IncomeStatement -and -not $ListAccounts -and $SelectCategory -eq "") {
    Write-Host "`n用法:" -ForegroundColor Yellow
    Write-Host "  .\balance_check.ps1 -BalanceSheet             資產負債表" -ForegroundColor DarkGray
    Write-Host "  .\balance_check.ps1 -IncomeStatement          損益表" -ForegroundColor DarkGray
    Write-Host "  .\balance_check.ps1 -ListAccounts             列出科目" -ForegroundColor DarkGray
    Write-Host "  .\balance_check.ps1 -SelectCategory Assets    選取類別" -ForegroundColor DarkGray
    Write-Host "`n按任意鍵結束..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

