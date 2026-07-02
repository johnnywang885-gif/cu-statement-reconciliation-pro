# audit_register.ps1 — 查核登記管理
# 用法:
#   .\audit_register.ps1 -List                                  (列出查核登記)
#   .\audit_register.ps1 -List -Year 113                        (依年度篩選)
#   .\audit_register.ps1 -List -CUNO A00001                     (依社員編號篩選)
#   .\audit_register.ps1 -Import "C:\path\to\score.xls"         (從 Excel 匯入)
#   .\audit_register.ps1 -CubPassword <密碼>

param(
    [string]$CubPath = "",
    [switch]$List,
    [string]$Year = "",
    [string]$CUNO = "",
    [string]$Import = "",
    [string]$CubPassword = ""
)

Set-Location (Split-Path -Parent $MyInvocation.MyCommand.Path)

# 內建 Get-MemberValue（避免依賴 AnomalyScore.psm1）
function Get-MemberValue {
    param($Value, $Default = 0)
    if ($null -eq $Value) { return $Default }
    return $Value
}

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
        if ($Year) { $argList += '-Year', $Year }
        if ($CUNO) { $argList += '-CUNO', $CUNO }
        if ($Import) { $argList += '-Import', "`"$Import`"" }
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
    $needWrite = ($Import -ne "")
    $db = $dbe.OpenDatabase($CubPath, $false, (-not $needWrite), $connectStr)

    # 檢查 查核登記表 是否存在，若無則建立
    $hasTable = @($db.TableDefs | Where-Object { $_.Name -eq "查核登記表" }).Count -gt 0
    if (-not $hasTable) {
        Write-Host "  查核登記表 不存在，建立中..." -ForegroundColor Yellow
        try {
            if (-not $needWrite) { $db.Close(); $db = $dbe.OpenDatabase($CubPath, $false, $false, $connectStr); $needWrite = $true }
            $td = $db.CreateTableDef("查核登記表")
            $td.Fields.Append($td.CreateField("SN", 4))       # 自動編號
            $td.Fields.Append($td.CreateField("YR", 3))        # 年度
            $td.Fields.Append($td.CreateField("CUNO", 10, 6))  # 社員編號
            $td.Fields.Append($td.CreateField("CUNM", 10, 10)) # 社員姓名
            $td.Fields.Append($td.CreateField("ChkDate", 8))   # 查核日期
            $td.Fields.Append($td.CreateField("BasDate", 8))   # 基準日期
            $td.Fields.Append($td.CreateField("SendDate", 8))  # 寄發日期
            $td.Fields.Append($td.CreateField("DocNO", 10, 20))# 文號
            $td.Fields.Append($td.CreateField("一", 4))        # 分數一
            $td.Fields.Append($td.CreateField("二", 4))        # 分數二
            $td.Fields.Append($td.CreateField("三", 4))        # 分數三
            $td.Fields.Append($td.CreateField("四", 4))        # 分數四
            $td.Fields.Append($td.CreateField("五", 4))        # 分數五
            $td.Fields.Append($td.CreateField("合計", 4))      # 合計
            $td.Fields.Append($td.CreateField("總評", 4))      # 總評
            $td.Fields.Append($td.CreateField("備註", 10, 100))
            $db.TableDefs.Append($td)
            Write-Host "  已建立 查核登記表" -ForegroundColor Green
        }
        catch {
            Write-Host "  無法建立查核登記表: $_" -ForegroundColor Red
        }
    }

    # Import from Excel score file (比照原 VBA cmdImport_Click)
    if ($Import -ne "") {
        if (-not $needWrite) { $db.Close(); $db = $dbe.OpenDatabase($CubPath, $false, $false, $connectStr) }
        if (-not (Test-Path $Import)) {
            Write-Host "  找不到檔案: $Import" -ForegroundColor Red
            $db.Close(); exit 1
        }

        Write-Host "  從 Excel 匯入查核資料: $Import" -ForegroundColor Yellow
        try {
            $db.Execute("DELETE FROM [查核登記表]")

            # 使用 DAO 連結 Excel
            $xlsPath = (Resolve-Path $Import).Path
            $linkName = "xls_Score_Link"

            # 移除舊連結
            try { $db.TableDefs.Delete($linkName) } catch {}

            $td = $db.CreateTableDef($linkName)
            $td.Connect = "Excel 12.0;HDR=YES;Database=$xlsPath"
            $td.SourceTableName = "xls_Score"
            $db.TableDefs.Append($td)

            $rs = $db.OpenRecordset("SELECT * FROM [$linkName]")
            $imported = 0
            while (-not $rs.EOF) {
                $cuno = if ($rs.Fields("CUNO").Value) { [string]$rs.Fields("CUNO").Value } else { "" }
                if ($cuno -eq "") { $rs.MoveNext(); continue }

                $chkDate = if ($rs.Fields("ChkDate").Value) { [string]$rs.Fields("ChkDate").Value } else { "" }
                $basDate = if ($rs.Fields("BasDate").Value) { [string]$rs.Fields("BasDate").Value } else { "" }
                $cunm = if ($rs.Fields("CUNM").Value) { [string]$rs.Fields("CUNM").Value } else { "" }
                $docNo = if ($rs.Fields("DocNO").Value) { [string]$rs.Fields("DocNO").Value } else { "" }
                $a1 = [double](Get-MemberValue $rs.Fields("一").Value 0)
                $a2 = [double](Get-MemberValue $rs.Fields("二").Value 0)
                $a3 = [double](Get-MemberValue $rs.Fields("三").Value 0)
                $a4 = [double](Get-MemberValue $rs.Fields("四").Value 0)
                $a5 = [double](Get-MemberValue $rs.Fields("五").Value 0)
                $sum = $a1 + $a2 + $a3 + $a4 + $a5
                $total = $sum
                $yr = if ($chkDate.Length -ge 7) { [int]($chkDate.Substring(0, 3)) } else { 0 }

                $qdef = $db.CreateQueryDef("", "SELECT SN FROM [查核登記表] WHERE CUNO=[p_cuno] AND ChkDate=[p_chkdate]")
                $qdef.Parameters("p_cuno").Value = $cuno
                $chkDateObj = try { [datetime]$chkDate } catch { $null }
                $qdef.Parameters("p_chkdate").Value = if ($chkDateObj) { $chkDateObj } else { $chkDate }
                $rs2 = $qdef.OpenRecordset()
                if ($rs2.EOF) {
                    $rs3 = $db.OpenRecordset("SELECT * FROM [查核登記表] WHERE False")
                    $rs3.AddNew()
                    $rs3.Fields("YR").Value = $yr
                    $rs3.Fields("CUNO").Value = $cuno
                    $rs3.Fields("CUNM").Value = $cunm
                    if ($chkDate) { $rs3.Fields("ChkDate").Value = $chkDate }
                    if ($basDate) { $rs3.Fields("BasDate").Value = $basDate }
                    if ($docNo) { $rs3.Fields("DocNO").Value = $docNo }
                    $rs3.Fields("一").Value = $a1
                    $rs3.Fields("二").Value = $a2
                    $rs3.Fields("三").Value = $a3
                    $rs3.Fields("四").Value = $a4
                    $rs3.Fields("五").Value = $a5
                    $rs3.Fields("合計").Value = $sum
                    $rs3.Fields("總評").Value = $total
                    $rs3.Update()
                    $rs3.Close()
                    $imported++
                }
                $rs2.Close()
                $rs.MoveNext()
            }
            $rs.Close()

            try { $db.TableDefs.Delete($linkName) } catch {}

            Write-Host "  已匯入 $imported 筆查核記錄" -ForegroundColor Green
        }
        catch {
            Write-Host "  匯入失敗: $_" -ForegroundColor Red
        }
    }

    # List audit register
    if ($List) {
        $whereClauses = @()
        $params = @{}
        if ($Year -ne "") { $whereClauses += "YR=" + [int]$Year }
        if ($CUNO -ne "") {
            $whereClauses += "CUNO=[p_cuno]"
            $params["p_cuno"] = $CUNO
        }
        $sql = "SELECT * FROM [查核登記表]"
        if ($whereClauses.Count -gt 0) {
            $sql += " WHERE " + ($whereClauses -join " AND ")
        }
        $sql += " ORDER BY CUNO, ChkDate"

        $qdef = $db.CreateQueryDef("", $sql)
        foreach ($k in $params.Keys) { $qdef.Parameters($k).Value = $params[$k] }
        $rs = $qdef.OpenRecordset()
        Write-Host "`n=== 查核登記清單 ===" -ForegroundColor Cyan
        $count = 0
        while (-not $rs.EOF) {
            $yr = if ($rs.Fields("YR").Value) { $rs.Fields("YR").Value } else { "" }
            $cuno = if ($rs.Fields("CUNO").Value) { $rs.Fields("CUNO").Value } else { "" }
            $cunm = if ($rs.Fields("CUNM").Value) { $rs.Fields("CUNM").Value } else { "" }
            $chk = if ($rs.Fields("ChkDate").Value) { $rs.Fields("ChkDate").Value } else { "" }
            $sum = if ($rs.Fields("合計").Value) { $rs.Fields("合計").Value } else { 0 }
            $total = if ($rs.Fields("總評").Value) { $rs.Fields("總評").Value } else { 0 }
            Write-Host ("  YR={0,-4} {1,-8} {2,-10} 查核日={3,-10} 合計={4,4} 總評={5,4}" -f $yr, $cuno, $cunm, $chk, $sum, $total)
            $count++
            $rs.MoveNext()
        }
        $rs.Close()
        Write-Host "  共 $count 筆" -ForegroundColor Green
    }

    $db.Close()
}
catch {
    Write-Host "  錯誤: $_" -ForegroundColor Red
}
finally {
    if ($dbe) { [Runtime.InteropServices.Marshal]::ReleaseComObject($dbe) | Out-Null }
}

if (-not $List -and $Import -eq "") {
    Write-Host "`n用法:" -ForegroundColor Yellow
    Write-Host "  .\audit_register.ps1 -List                    列出查核登記" -ForegroundColor DarkGray
    Write-Host "  .\audit_register.ps1 -List -Year 113          依年度篩選" -ForegroundColor DarkGray
    Write-Host "  .\audit_register.ps1 -Import score.xls        從 Excel 匯入" -ForegroundColor DarkGray
    Write-Host "`n按任意鍵結束..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

