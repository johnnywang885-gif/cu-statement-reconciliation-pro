# loan_merge_check.ps1 — 貸款查核合併作業
# 用法:
#   .\loan_merge_check.ps1 -MergeSource "D:\audit.mdb" [-MergeTarget "CUB.MDB"] [-SubstantiveReview] [-ApplicationReview]
#   .\loan_merge_check.ps1 -CompareList                         (列出待合併比對資料)
#   .\loan_merge_check.ps1 -CubPassword <密碼>
#
# 參數:
#   -MergeSource <path>      來源 MDB (含 M_Loan_All 表)
#   -MergeTarget <path>      目標 MDB (預設為目前目錄的 CUB.MDB)
#   -SubstantiveReview       實質審核模式 (合併審查意見、簽章等)
#   -ApplicationReview       時機申請書模式 (合併申請書/時機資料)
#   -CompareList             列出合併比對結果

param(
    [string]$MergeSource = "",
    [string]$MergeTarget = "",
    [switch]$SubstantiveReview,
    [switch]$ApplicationReview,
    [switch]$CompareList,
    [string]$SourcePassword = "",
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
        if ($MergeSource) { $argList += '-MergeSource', "`"$MergeSource`"" }
        if ($MergeTarget) { $argList += '-MergeTarget', "`"$MergeTarget`"" }
        if ($SubstantiveReview) { $argList += '-SubstantiveReview' }
        if ($ApplicationReview) { $argList += '-ApplicationReview' }
        if ($CompareList) { $argList += '-CompareList' }
        if ($CubPassword -ne '') { $argList += '-CubPassword', $CubPassword }
        if ($SourcePassword -ne '') { $argList += '-SourcePassword', $SourcePassword }
        & $ps32 -ExecutionPolicy Bypass -File "`"$PSCommandPath`"" @argList
        exit $LASTEXITCODE
    }
}

if (-not $daoAvailable) {
    Write-Host "  DAO.DBEngine 無法初始化，請安裝 Access Database Engine 2016" -ForegroundColor Red
    exit 1
}

$cubPath = if ($MergeTarget) { $MergeTarget } else { Join-Path $PSScriptRoot "CUB.MDB" }
$cubPath = [System.IO.Path]::GetFullPath($cubPath)

if (-not (Test-Path $cubPath)) { Write-Host "  找不到目標 MDB: $cubPath" -ForegroundColor Red; exit 1 }

$connectStr = ";PWD=$CubPassword"

try {
    $dbe = New-Object -ComObject DAO.DBEngine.120
    $db = $dbe.OpenDatabase($cubPath, $false, $false, $connectStr)  # 可寫入

    # 檢查 M_Loan_All 是否存在於目標
    $hasTargetLoanTable = @($db.TableDefs | Where-Object { $_.Name -eq "M_Loan_All" }).Count -gt 0
    if (-not $hasTargetLoanTable) {
        Write-Host "  目標 MDB 無 M_Loan_All 表，建立中..." -ForegroundColor Yellow
        try {
            $td = $db.CreateTableDef("M_Loan_All")
            $td.Fields.Append($td.CreateField("帳號", 10, 6))
            $td.Fields.Append($td.CreateField("戶號", 10, 12))
            $td.Fields.Append($td.CreateField("審查意見", 10, 50))
            $td.Fields.Append($td.CreateField("簽章", 10, 20))
            $td.Fields.Append($td.CreateField("備註", 10, 100))
            $td.Fields.Append($td.CreateField("擔保品總額", 4))
            $td.Fields.Append($td.CreateField("CA", 1))
            $td.Fields.Append($td.CreateField("CI", 1))
            $td.Fields.Append($td.CreateField("時機", 10, 20))
            $td.Fields.Append($td.CreateField("申請書", 10, 20))
            $td.Fields.Append($td.CreateField("ToBeShowed", 1))
            $td.Fields.Append($td.CreateField("Repeat", 1))
            $db.TableDefs.Append($td)
            Write-Host "  已建立 M_Loan_All" -ForegroundColor Green
        }
        catch {
            Write-Host "  無法建立 M_Loan_All: $_" -ForegroundColor Red
            $db.Close(); exit 1
        }
    }

    # 合併模式
    if ($MergeSource -ne "") {
        if (-not (Test-Path $MergeSource)) {
            Write-Host "  找不到來源 MDB: $MergeSource" -ForegroundColor Red
            $db.Close(); exit 1
        }

        $sourcePath = (Resolve-Path $MergeSource).Path
        Write-Host "  來源 MDB: $sourcePath" -ForegroundColor DarkGray

        # 連結來源 MDB 為 TM 表
        $linkName = "TM"
        try { $db.TableDefs.Delete($linkName) } catch {}

        $td = $db.CreateTableDef($linkName)
        $srcPwd = if ($SourcePassword) { ";PWD=$SourcePassword" } else { "" }
        $td.Connect = ";DATABASE=$sourcePath$srcPwd"
        $td.SourceTableName = "M_Loan_All"
        $db.TableDefs.Append($td)

        # 檢查 TM 是否有資料
        $rs = $db.OpenRecordset("SELECT COUNT(*) AS CNT FROM [TM]")
        $cnt = $rs.Fields("CNT").Value
        $rs.Close()

        if ($cnt -eq 0) {
            Write-Host "  來源 MDB 的 M_Loan_All 無資料" -ForegroundColor Yellow
        }
        else {
            Write-Host "  來源 MDB 有 $cnt 筆資料" -ForegroundColor DarkGray

            if ($SubstantiveReview) {
                Write-Host "`n  [實質審核模式] 合併審查意見、簽章、備註、擔保品..." -ForegroundColor Yellow
                try {
                    $db.Execute(@"
UPDATE M_Loan_All INNER JOIN TM ON (M_Loan_All.帳號 = TM.帳號) AND (M_Loan_All.戶號 = TM.戶號)
SET M_Loan_All.審查意見 = TM.審查意見,
    M_Loan_All.簽章 = TM.簽章,
    M_Loan_All.備註 = TM.備註,
    M_Loan_All.擔保品總額 = TM.擔保品總額,
    M_Loan_All.ToBeShowed = TM.ToBeShowed,
    M_Loan_All.Repeat = TM.Repeat
WHERE (TM.審查意見 <> '' OR TM.簽章 <> '' OR TM.備註 <> '' OR TM.擔保品總額 > 0)
"@)
                    Write-Host "  實質審核資料合併完成" -ForegroundColor Green
                }
                catch {
                    Write-Host "  實質審核合併失敗: $_" -ForegroundColor Red
                }
            }

            if ($ApplicationReview) {
                Write-Host "`n  [時機申請書模式] 合併時機/申請書資料..." -ForegroundColor Yellow
                try {
                    $db.Execute(@"
UPDATE M_Loan_All INNER JOIN TM ON (M_Loan_All.帳號 = TM.帳號) AND (M_Loan_All.戶號 = TM.戶號)
SET M_Loan_All.時機 = TM.時機,
    M_Loan_All.申請書 = TM.申請書,
    M_Loan_All.CA = TM.CA,
    M_Loan_All.CI = TM.CI,
    M_Loan_All.ToBeShowed = TM.ToBeShowed,
    M_Loan_All.Repeat = TM.Repeat
WHERE (TM.CA='v' OR TM.CI='v')
"@)
                    Write-Host "  時機申請書資料合併完成" -ForegroundColor Green
                }
                catch {
                    Write-Host "  時機申請書合併失敗: $_" -ForegroundColor Red
                }
            }

            # 若無指定模式，則列出差異
            if (-not $SubstantiveReview -and -not $ApplicationReview) {
                Write-Host "  請指定合併模式: -SubstantiveReview 或 -ApplicationReview" -ForegroundColor Yellow
                $rs = $db.OpenRecordset(@"
SELECT TM.帳號, TM.戶號, TM.審查意見, TM.簽章, TM.CA, TM.CI
FROM TM
WHERE (TM.審查意見 <> '' OR TM.簽章 <> '' OR TM.CA='v' OR TM.CI='v')
ORDER BY TM.帳號
"@)
                Write-Host "`n  待合併資料:" -ForegroundColor Cyan
                while (-not $rs.EOF) {
                    $acc = if ($rs.Fields("帳號").Value) { $rs.Fields("帳號").Value } else { "" }
                    $brno = if ($rs.Fields("戶號").Value) { $rs.Fields("戶號").Value } else { "" }
                    $opinion = if ($rs.Fields("審查意見").Value) { "審" } else { "" }
                    $sign = if ($rs.Fields("簽章").Value) { "簽" } else { "" }
                    $ca = if ($rs.Fields("CA").Value -eq 'v') { "CA" } else { "" }
                    $ci = if ($rs.Fields("CI").Value -eq 'v') { "CI" } else { "" }
                    $tags = @($opinion, $sign, $ca, $ci) | Where-Object { $_ -ne "" } | ForEach-Object { $_ }
                    Write-Host "    $acc / $brno [$($tags -join ',')]" -ForegroundColor DarkGray
                    $rs.MoveNext()
                }
                $rs.Close()
            }
        }

        # 清除 TM 連結
        try { $db.TableDefs.Delete($linkName) } catch {}
    }

    # 列出比對結果
    if ($CompareList) {
        $rs = $db.OpenRecordset("SELECT 帳號, 戶號, 審查意見, 簽章, 備註, 擔保品總額, CA, CI FROM [M_Loan_All] ORDER BY 帳號")
        Write-Host "`n=== M_Loan_All 合併資料清單 ===" -ForegroundColor Cyan
        $count = 0
            while (-not $rs.EOF) {
                $acc = if ($rs.Fields("帳號").Value) { $rs.Fields("帳號").Value } else { "" }
                $brno = if ($rs.Fields("戶號").Value) { $rs.Fields("戶號").Value } else { "" }
                $opinion = if ($rs.Fields("審查意見").Value) { [string]$rs.Fields("審查意見").Value } else { "" }
                $sign = if ($rs.Fields("簽章").Value) { [string]$rs.Fields("簽章").Value } else { "" }
                $collateral = if ($rs.Fields("擔保品總額").Value) { [double]$rs.Fields("擔保品總額").Value } else { 0 }
                $ca = if ($rs.Fields("CA").Value) { [string]$rs.Fields("CA").Value } else { "" }
                $ci = if ($rs.Fields("CI").Value) { [string]$rs.Fields("CI").Value } else { "" }
                Write-Host ("  {0,-8} {1,-8} 審={2,-8} 簽={3,-8} 擔={4,10:N0} CA={5,-2} CI={6,-2}" -f $acc, $brno, $opinion, $sign, $collateral, $ca, $ci)
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

if ($MergeSource -eq "" -and -not $CompareList) {
    Write-Host "`n用法:" -ForegroundColor Yellow
    Write-Host "  .\loan_merge_check.ps1 -MergeSource audit.mdb -SubstantiveReview" -ForegroundColor DarkGray
    Write-Host "  .\loan_merge_check.ps1 -MergeSource audit.mdb -ApplicationReview" -ForegroundColor DarkGray
    Write-Host "  .\loan_merge_check.ps1 -CompareList" -ForegroundColor DarkGray
    Write-Host "`n按任意鍵結束..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

