# manage_nonmail.ps1 — 不寄發對帳單登錄管理
# 用法:
#   .\manage_nonmail.ps1                           (互動模式)
#   .\manage_nonmail.ps1 -List                     (列出不寄發清單)
#   .\manage_nonmail.ps1 -Add ACCNO                (加入不寄發)
#   .\manage_nonmail.ps1 -Remove ACCNO             (移除不寄發)
#   .\manage_nonmail.ps1 -CubPassword <密碼>

param(
    [string]$CubPath = "",
    [switch]$List,
    [string]$Add = "",
    [string]$Remove = "",
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
        if ($List) { $argList += '-List' }
        if ($Add) { $argList += '-Add', $Add }
        if ($Remove) { $argList += '-Remove', $Remove }
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

    # List 模式用唯讀，Add/Remove 用可寫入
    $needWrite = ($Add -ne "") -or ($Remove -ne "")
    $db = $dbe.OpenDatabase($CubPath, $false, (-not $needWrite), $connectStr)

    $has明細 = @($db.TableDefs | Where-Object { $_.Name -eq "M_對帳明細" }).Count -gt 0

    if (-not $has明細) {
        Write-Host "  M_對帳明細 不存在 CUB.MDB 中" -ForegroundColor Red
        $db.Close()
        exit 1
    }

    if ($List -or ($Add -eq "" -and $Remove -eq "")) {
        $rs = $db.OpenRecordset("SELECT 帳號, 姓名, 不寄發, 不寄發原因 FROM [M_對帳明細] WHERE 不寄發=True ORDER BY 帳號")
        Write-Host "`n=== 不寄發對帳單清單 ===" -ForegroundColor Cyan
        $count = 0
        while (-not $rs.EOF) {
            $acc = $rs.Fields("帳號").Value
            $name = $rs.Fields("姓名").Value
            $reason = if ($rs.Fields("不寄發原因").Value) { $rs.Fields("不寄發原因").Value } else { "" }
            Write-Host ("  {0,-8} {1,-10}  原因: {2}" -f $acc, $name, $reason)
            $count++
            $rs.MoveNext()
        }
        $rs.Close()
        Write-Host "  共 $count 筆" -ForegroundColor Green

        if ($Add -eq "" -and $Remove -eq "") {
            Write-Host "`n  指令:" -ForegroundColor Yellow
            Write-Host "    .\manage_nonmail.ps1 -Add A00001     加入不寄發" -ForegroundColor DarkGray
            Write-Host "    .\manage_nonmail.ps1 -Remove A00001  移除不寄發" -ForegroundColor DarkGray
        }
    }

    if ($Add -ne "") {
        if (-not $needWrite) { $db.Close(); $db = $dbe.OpenDatabase($CubPath, $false, $false, $connectStr) }
        $qdef = $db.CreateQueryDef("", "SELECT 帳號, 姓名 FROM [M_對帳明細] WHERE 帳號=[p_acc]")
        $qdef.Parameters("p_acc").Value = $Add
        $rs = $qdef.OpenRecordset()
        if ($rs.EOF) {
            Write-Host "  帳號 $Add 不存在 M_對帳明細 中" -ForegroundColor Red
        }
        else {
            $qdef = $db.CreateQueryDef("", "UPDATE [M_對帳明細] SET 不寄發=True, 不寄發原因='手動登錄' WHERE 帳號=[p_acc]")
            $qdef.Parameters("p_acc").Value = $Add
            $qdef.Execute()
            Write-Host "  已將 $Add ($($rs.Fields("姓名").Value)) 加入不寄發清單" -ForegroundColor Green
        }
        $rs.Close()
    }

    if ($Remove -ne "") {
        if (-not $needWrite) { $db.Close(); $db = $dbe.OpenDatabase($CubPath, $false, $false, $connectStr) }
        $qdef = $db.CreateQueryDef("", "UPDATE [M_對帳明細] SET 不寄發=False, 不寄發原因=Null WHERE 帳號=[p_acc]")
        $qdef.Parameters("p_acc").Value = $Remove
        $qdef.Execute()
        Write-Host "  已將 $Remove 移除不寄發清單" -ForegroundColor Green
    }

    $db.Close()
}
catch {
    Write-Host "  錯誤: $_" -ForegroundColor Red
}
finally {
    if ($dbe) { [Runtime.InteropServices.Marshal]::ReleaseComObject($dbe) | Out-Null }
}

if ($Add -eq "" -and $Remove -eq "" -and -not $List) {
    Write-Host "`n按任意鍵結束..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

