# reconcile_reply.ps1 — 對帳單回覆登錄
# 用法:
#   .\reconcile_reply.ps1 -List                               (列出所有回覆)
#   .\reconcile_reply.ps1 -Find ACCNO                         (查詢特定帳號回覆)
#   .\reconcile_reply.ps1 -FilterYearMonth 11305              (依年月篩選)
#   .\reconcile_reply.ps1 -Register ACCNO -YearMonth 11305    (登錄回覆)
#   .\reconcile_reply.ps1 -CubPassword <密碼>

param(
    [switch]$List,
    [string]$Find = "",
    [string]$FilterYearMonth = "",
    [string]$Register = "",
    [string]$YearMonth = "",
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
        if ($List) { $argList += '-List' }
        if ($Find) { $argList += '-Find', $Find }
        if ($FilterYearMonth) { $argList += '-FilterYearMonth', $FilterYearMonth }
        if ($Register) { $argList += '-Register', $Register }
        if ($YearMonth) { $argList += '-YearMonth', $YearMonth }
        if ($CubPassword -ne '') { $argList += '-CubPassword', $CubPassword }
        & $ps32 -ExecutionPolicy Bypass -File "`"$PSCommandPath`"" @argList
        exit $LASTEXITCODE
    }
}

if (-not $daoAvailable) {
    Write-Host "  DAO.DBEngine 無法初始化，請安裝 Access Database Engine 2016" -ForegroundColor Red
    exit 1
}

$cubPath = Join-Path $PSScriptRoot "CUB.MDB"
if (-not (Test-Path $cubPath)) { Write-Host "  找不到 CUB.MDB" -ForegroundColor Red; exit 1 }

$connectStr = ";PWD=$CubPassword"

try {
    $dbe = New-Object -ComObject DAO.DBEngine.120
    $needWrite = ($Register -ne "")
    $db = $dbe.OpenDatabase($cubPath, $false, (-not $needWrite), $connectStr)

    # 檢查 k_對帳單回覆 是否存在，若不存在則建立（需可寫入）
    $hasTable = @($db.TableDefs | Where-Object { $_.Name -eq "k_對帳單回覆" }).Count -gt 0
    if (-not $hasTable) {
        if (-not $needWrite) { $db.Close(); $db = $dbe.OpenDatabase($cubPath, $false, $false, $connectStr); $needWrite = $true }
        Write-Host "  k_對帳單回覆 不存在於 CUB.MDB，嘗試建立..." -ForegroundColor Yellow
        try {
            $td = $db.CreateTableDef("k_對帳單回覆")
            $td.Fields.Append($td.CreateField("年月", 10, 6))
            $td.Fields.Append($td.CreateField("科目", 10, 6))
            $td.Fields.Append($td.CreateField("帳號", 10, 6))
            $td.Fields.Append($td.CreateField("戶名", 10, 10))
            $td.Fields.Append($td.CreateField("回覆日期", 8))
            $td.Fields.Append($td.CreateField("備註", 10, 50))
            $db.TableDefs.Append($td)
            Write-Host "  已建立 k_對帳單回覆 表" -ForegroundColor Green
        }
        catch {
            Write-Host "  無法建立 k_對帳單回覆: $_" -ForegroundColor Red
            $db.Close(); exit 1
        }
    }
    # 若 Register 需要寫入但當前是唯讀，重新開啟
    if ($Register -ne "" -and -not $needWrite) {
        $db.Close(); $db = $dbe.OpenDatabase($cubPath, $false, $false, $connectStr)
    }

    $where = @()
    if ($FilterYearMonth -ne "") { $where += "年月='" + $FilterYearMonth + "'" }
    if ($Find -ne "") {
        $padded = Pad-Account $Find
        $where += "帳號='" + $padded + "'"
    }

    $sql = "SELECT * FROM [k_對帳單回覆]"
    if ($where.Count -gt 0) {
        $sql += " WHERE " + ($where -join " AND ")
    }
    $sql += " ORDER BY 年月, 帳號"

    $rs = $db.OpenRecordset($sql)
    Write-Host "`n=== 對帳單回覆清單 ===" -ForegroundColor Cyan
    $count = 0
    while (-not $rs.EOF) {
        $ym = if ($rs.Fields("年月").Value) { $rs.Fields("年月").Value } else { "" }
        $acc = if ($rs.Fields("帳號").Value) { $rs.Fields("帳號").Value } else { "" }
        $name = if ($rs.Fields("戶名").Value) { $rs.Fields("戶名").Value } else { "" }
        $rdate = if ($rs.Fields("回覆日期").Value) { $rs.Fields("回覆日期").Value } else { "" }
        $memo = if ($rs.Fields("備註").Value) { $rs.Fields("備註").Value } else { "" }
        Write-Host ("  {0,-8} {1,-8} {2,-10} {3,-12} {4}" -f $ym, $acc, $name, $rdate, $memo)
        $count++
        $rs.MoveNext()
    }
    $rs.Close()
    Write-Host "  共 $count 筆" -ForegroundColor Green

    # 登錄新回覆
    if ($Register -ne "") {
        $ym = if ($YearMonth -ne "") { $YearMonth } else { (Get-Date).ToString("yyyMM") }
        $padded = Pad-Account $Register
        $today = (Get-Date).ToString("yyyy/MM/dd")

        # 查戶名
        $rs2 = $db.OpenRecordset("SELECT ACCNM FROM SER WHERE ACCNO='$padded'")
        $name = if (-not $rs2.EOF) { $rs2.Fields("ACCNM").Value } else { "" }
        $rs2.Close()

        $rs2 = $db.OpenRecordset("SELECT * FROM [k_對帳單回覆] WHERE 年月='$ym' AND 帳號='$padded'")
        if ($rs2.EOF) {
            $rs2.AddNew()
            $rs2.Fields("年月").Value = $ym
            $rs2.Fields("科目").Value = ""
            $rs2.Fields("帳號").Value = $padded
            $rs2.Fields("戶名").Value = $name
            $rs2.Fields("回覆日期").Value = $today
            $rs2.Fields("備註").Value = ""
            $rs2.Update()
            Write-Host "  已登錄回覆: $padded ($name) 於 $ym" -ForegroundColor Green
        }
        else {
            Write-Host "  此筆已存在: $padded 於 $ym" -ForegroundColor Yellow
        }
        $rs2.Close()
    }

    $db.Close()
}
catch {
    Write-Host "  錯誤: $_" -ForegroundColor Red
}
finally {
    if ($dbe) { [Runtime.InteropServices.Marshal]::ReleaseComObject($dbe) | Out-Null }
}

function Pad-Account {
    param([string]$Acc)
    if ($Acc.Length -eq 0) { return $Acc }
    $first = $Acc[0]
    $rest = if ($Acc.Length -gt 1) { $Acc.Substring(1) } else { "" }
    if ([char]::IsDigit($first)) {
        return $first.ToString() + $rest.PadLeft(5, '0')
    }
    else {
        return $first.ToString() + $rest.PadLeft(5, '0')
    }
}

if (-not $List -and $Find -eq "" -and $FilterYearMonth -eq "" -and $Register -eq "") {
    Write-Host "`n用法:" -ForegroundColor Yellow
    Write-Host "  .\reconcile_reply.ps1 -List                        列出所有回覆" -ForegroundColor DarkGray
    Write-Host "  .\reconcile_reply.ps1 -Find ACCNO                  查詢特定帳號" -ForegroundColor DarkGray
    Write-Host "  .\reconcile_reply.ps1 -FilterYearMonth 11305       依年月篩選" -ForegroundColor DarkGray
    Write-Host "  .\reconcile_reply.ps1 -Register ACCNO -YearMonth 11305  登錄回覆" -ForegroundColor DarkGray
    Write-Host "`n按任意鍵結束..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

