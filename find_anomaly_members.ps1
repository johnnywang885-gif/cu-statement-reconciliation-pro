# find_anomaly_members.ps1 — 異常社員偵測引擎，排序輸出，建議優先追查社員
# 用法:
#   .\find_anomaly_members.ps1                      (自動偵測尋找 CUB.MDB)
#   .\find_anomaly_members.ps1 -CubPath D:\CUB.MDB  (指定路徑)
# 輸出: CUB_異常社員_對帳單排序_YYYYMMDD_HHmmss.csv

param(
    [string]$CubPath = "",
    [string]$CubPassword = ""
)

Set-Location (Split-Path -Parent $MyInvocation.MyCommand.Path)

$libPath = Join-Path $PSScriptRoot 'lib\AnomalyScore.psm1'
if (Test-Path $libPath) { Import-Module $libPath -Force -DisableNameChecking }

# 內建 Get-MemberValue（避免依賴 AnomalyScore.psm1）
function Get-MemberValue {
    param($Value, $Default = 0)
    if ($null -eq $Value) { return $Default }
    return $Value
}

if ([string]::IsNullOrEmpty($CubPassword)) {
    Write-Host "════════════════════════════════════════════════════════════════════════════════════" -ForegroundColor Yellow
    Write-Host "║   CUB.MDB 需要密碼                                          ║" -ForegroundColor Yellow
    Write-Host "║   提示: 也可用 -CubPassword <密碼> 略過此詢問                ║" -ForegroundColor Yellow
    Write-Host "════════════════════════════════════════════════════════════════════════════════════" -ForegroundColor Yellow
    $secure = Read-Host -Prompt "  請輸入 CUB.MDB 密碼" -AsSecureString
    $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try {
        $CubPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($ptr)
    }
    finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr) | Out-Null
    }
    Write-Host ""
}

# ── 檢查 0：DAO.DBEngine 初始化（自動切換 64/32-bit）──────────────────────────
$daoAvailable = $false
try { $null = New-Object -ComObject DAO.DBEngine.120; $daoAvailable = $true } catch {}

if (-not $daoAvailable -and -not $env:DAO_RESTARTED) {
    $ps32 = "$env:SystemRoot\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    if (Test-Path $ps32) {
        Write-Host "  偵測到 32-bit 需求，自動切換..." -ForegroundColor Yellow
        $env:DAO_RESTARTED = '1'
        $argList = @()
        if ($CubPath) { $argList += '-CubPath', $CubPath }
        if ($CubPassword -ne '') { $argList += '-CubPassword', $CubPassword }
        & $ps32 -ExecutionPolicy Bypass -File $PSCommandPath @argList
        exit $LASTEXITCODE
    }
}

if (-not $daoAvailable) {
    Write-Host "════════════════════════════════════════════════════════════════════════════════════" -ForegroundColor Red
    Write-Host "║   DAO.DBEngine 無法初始化                               ║" -ForegroundColor Red
    Write-Host "║   請安裝 Microsoft Access Database Engine 2016 (可再發行)    ║" -ForegroundColor Red
    Write-Host("║   下載: https://www.microsoft.com/download/details.aspx?id=54920") -ForegroundColor Red
    Write-Host "╚═══════════════════════════════════════════════════════════════════════════════════" -ForegroundColor Red
    Write-Host "`n按任意鍵結束..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# ── 檢查 1：CUB.MDB 是否存在 ──────────────────────────────────────────────────────
if ($CubPath -eq "") {
    $CubPath = Join-Path $PSScriptRoot "CUB.MDB"
}
$CubPath = [System.IO.Path]::GetFullPath($CubPath)

if (-not (Test-Path $CubPath)) {
    Write-Host "════════════════════════════════════════════════════════════════════════════════════" -ForegroundColor Red
    Write-Host "║   找不到 CUB.MDB                                       ║" -ForegroundColor Red
    Write-Host("║   預設路徑: $CubPath") -ForegroundColor Red
    Write-Host "║   請把 CUB.MDB 放在此腳本同一個資料夾?             ║" -ForegroundColor Red
    Write-Host "║   或用 -CubPath 參數指定路徑                           ║" -ForegroundColor Red
    Write-Host "╚═══════════════════════════════════════════════════════════════════════════════════" -ForegroundColor Red
    Write-Host "`n按任意鍵結束..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# ── 連線字串（密碼相容）──────────────────────────────────────────────────────
$connectStr = if ($CubPassword) { ";PWD=$CubPassword" } else { "" }

# ── 開始 ────────────────────────────────────────────────────────────────────────────
Write-Host "════════════════════════════════════════════════════════════════════════════════════" -ForegroundColor White
Write-Host "║       異常社員偵測引擎 — 對帳單優先追查工具          ║" -ForegroundColor White
Write-Host "════════════════════════════════════════════════════════════════════════════════════" -ForegroundColor White
Write-Host ("  CUB: $CubPath") -ForegroundColor DarkGray
if ($CubPassword) { Write-Host "  密碼: ****" -ForegroundColor DarkGray }
Write-Host ""

try {
    $dbe = New-Object -ComObject DAO.DBEngine.120

    # ── Step 1: 讀取查核期間 ───────────────────────────────────────────────────────────
    Write-Host "[1/8] 讀取查核期間..." -ForegroundColor Yellow
    try {
        $db = $dbe.OpenDatabase($CubPath, $false, $true, $connectStr)
    }
    catch {
        $errMsg = $_.Exception.Message
        if ($errMsg -match "密碼") {
            Write-Host "════════════════════════════════════════════════════════════════════════════════════" -ForegroundColor Red
            Write-Host "║   CUB.MDB 密碼有誤，無法開啟                         ║" -ForegroundColor Red
            Write-Host "║   請用 -CubPassword 參數指定密碼                       ║" -ForegroundColor Red
            Write-Host "║   範例: powershell -File find_anomaly_members.ps1 -CubPassword <密碼>" -ForegroundColor Red
            Write-Host "╚═══════════════════════════════════════════════════════════════════════════════════" -ForegroundColor Red
        }
        else {
            Write-Host "  CUB.MDB 無法開啟，可能不是有效的 Access 資料庫" -ForegroundColor Red
        }
        throw $_
    }
    $rs = $db.OpenRecordset("SELECT Item, Para FROM k_para")
    $kPara = @{}
    while (-not $rs.EOF) {
        $kPara[$rs.Fields["Item"].Value] = $rs.Fields["Para"].Value
        $rs.MoveNext()
    }
    $rs.Close(); $db.Close()

    $bDate = $kPara["bDate"]; $eDate = $kPara["eDate"]
    if ([string]::IsNullOrWhiteSpace($bDate) -or [string]::IsNullOrWhiteSpace($eDate)) {
        Write-Host "════════════════════════════════════════════════════════════════════════════════════" -ForegroundColor Red
        Write-Host "║   CUB.MDB 的 k_para 缺少 bDate 或 eDate               ║" -ForegroundColor Red
        Write-Host "║   請確認 CUB 是否已正確設定                            ║" -ForegroundColor Red
        Write-Host "╚═══════════════════════════════════════════════════════════════════════════════════" -ForegroundColor Red
        throw "k_para 參數不足"
    }
    Write-Host "  查核期間: $bDate 到 $eDate" -ForegroundColor DarkGray

    # ── 讀取 k_para 其他系統參數 ──────────────────────────────────────────
    $滯納天數上限 = 90
    if ($kPara.ContainsKey("OverdueDays")) {
        $v = $kPara["OverdueDays"]; if ($v -and $v -ne '') { $滯納天數上限 = [int]$v }
    }
    Write-Host "  系統設定滯納天數上限: $滯納天數上限" -ForegroundColor DarkGray

    # ── Step 2: 讀取社員主檔 SER ────────────────────────────────────────────────────────
    Write-Host "[2/8] 讀取社員主檔..." -ForegroundColor Yellow
    $db = $dbe.OpenDatabase($CubPath, $false, $true, $connectStr)
    $rs = $db.OpenRecordset("SELECT ACCNO, ACCNM, NAM, PRESTK, SSAV, NOSTK, PREBO, BORO, LNMNY, PREFO, FO, NOFO, ADDR, TEL, BIRTH, SEX, GRPNO, INDAT, TYPE FROM SER WHERE (OUTDAT IS NULL OR OUTDAT='') ORDER BY ACCNO")
    $memberMap = @{}
    $memberCount = 0
    while (-not $rs.EOF) {
        $a = $rs.Fields["ACCNO"].Value
        $memberMap[$a] = @{
            ACCNO                  = $a
            ACCNM                  = $rs.Fields["ACCNM"].Value
            NAM                    = $rs.Fields["NAM"].Value
            ADDR                   = $rs.Fields["ADDR"].Value
            TEL                    = $rs.Fields["TEL"].Value
            BIRTH                  = [string]$rs.Fields["BIRTH"].Value
            SEX                    = [string]$rs.Fields["SEX"].Value
            GRPNO                  = [string]$rs.Fields["GRPNO"].Value
            INDAT                  = [string]$rs.Fields["INDAT"].Value
            TYPE                   = [string]$rs.Fields["TYPE"].Value
            SER_Share              = [double]($rs.Fields["PRESTK"].Value + $rs.Fields["SSAV"].Value - $rs.Fields["NOSTK"].Value)
            SER_Loan               = [double]($rs.Fields["PREBO"].Value + $rs.Fields["BORO"].Value - $rs.Fields["LNMNY"].Value)
            SER_Reserve            = [double]($rs.Fields["PREFO"].Value + $rs.Fields["FO"].Value - $rs.Fields["NOFO"].Value)
            LoanCount              = 0
            IsNonMember            = $false
            HasSameAddressMultiple = $false
            BadLoanCount = 0; TotOWEBR = 0; TotOWEINT = 0
            RecentTxn              = 0
            HasRelatedPartyLoan    = $false
            HasNewJoinLoan         = $false
            HasRecDiscrepancy      = $false
            HasAuditOverdue        = $false
            HasExceedPayDate       = $false
            HasDirectorOverLimit   = $false
            HasDuplicateLoan       = $false
            HasDormantActivation   = $false
            HasRoundAmountTxn      = $false
            HasNewAccountBurst     = $false
            HasNegativeLoanBalance = $false
            HasLoanBeforeApproval  = $false
            LastTxnBeforePeriod    = $null
        }
        # HasNegativeLoanBalance: 用 PREBO 判斷貸款結餘為負
        $prebo = [double]($rs.Fields["PREBO"].Value)
        if ($prebo -lt 0) {
            $memberMap[$a].HasNegativeLoanBalance = $true
        }
        $memberCount++
        $rs.MoveNext()
    }
    $rs.Close(); $db.Close()
    Write-Host "  共 $memberCount 位會員" -ForegroundColor DarkGray

    # ── Step 3: 從 LGR 計算各項餘額 ─────────────────────────────────────────────────────
    Write-Host "[3/8] 從 LGR 計算各項餘額..." -ForegroundColor Yellow

    $db = $dbe.OpenDatabase($CubPath, $false, $true, $connectStr)
    $rs = $db.OpenRecordset("SELECT ACCNO, Sum(IIf(DC='C',1,-1)*MNY) AS BLN FROM LGR WHERE (Left(ACNO,3)='310' OR ACNO='216') AND DCHK='Y' GROUP BY ACCNO")
    while (-not $rs.EOF) {
        $a = $rs.Fields["ACCNO"].Value
        if ($memberMap.ContainsKey($a)) { $memberMap[$a].LGR_Share = [double]$rs.Fields["BLN"].Value }
        $rs.MoveNext()
    }
    $rs.Close()

    $rs = $db.OpenRecordset("SELECT ACCNO, Sum(IIf(DC='C',-1,1)*MNY) AS BLN FROM LGR WHERE Left(ACNO,3)='131' AND DCHK='Y' GROUP BY ACCNO")
    while (-not $rs.EOF) {
        $a = $rs.Fields["ACCNO"].Value
        if ($memberMap.ContainsKey($a)) { $memberMap[$a].LGR_Loan = [double]$rs.Fields["BLN"].Value }
        $rs.MoveNext()
    }
    $rs.Close()

    $rs = $db.OpenRecordset("SELECT ACCNO, Sum(IIf(DC='C',1,-1)*MNY) AS BLN FROM LGR WHERE ACNO='226' AND DCHK='Y' GROUP BY ACCNO")
    while (-not $rs.EOF) {
        $a = $rs.Fields["ACCNO"].Value
        if ($memberMap.ContainsKey($a)) { $memberMap[$a].LGR_Reserve = [double]$rs.Fields["BLN"].Value }
        $rs.MoveNext()
    }
    $rs.Close()
    $db.Close()

    foreach ($a in $memberMap.Keys) {
        if (-not $memberMap[$a].ContainsKey("LGR_Share")) { $memberMap[$a].LGR_Share = 0 }
        if (-not $memberMap[$a].ContainsKey("LGR_Loan")) { $memberMap[$a].LGR_Loan = 0 }
        if (-not $memberMap[$a].ContainsKey("LGR_Reserve")) { $memberMap[$a].LGR_Reserve = 0 }
    }

    # ── Step 4: 逾期貸款 + 貸款筆數 ───────────────────────────────────────────────────
    Write-Host "[4/8] 檢查逾期貸款及貸款筆數..." -ForegroundColor Yellow
    $db = $dbe.OpenDatabase($CubPath, $false, $true, $connectStr)
    $rs = $db.OpenRecordset("SELECT ACCNO, Count(*) AS BadCnt, Sum(IIf(IsNull(OWEBR),0,OWEBR)) AS TotOWEBR, Sum(IIf(IsNull(OWEINT),0,OWEINT)) AS TotOWEINT FROM BOROW WHERE IIf(IsNull(OWEBR),0,OWEBR) > 0 OR IIf(IsNull(OWEINT),0,OWEINT) > 0 GROUP BY ACCNO")
    while (-not $rs.EOF) {
        $a = $rs.Fields["ACCNO"].Value
        if ($memberMap.ContainsKey($a)) {
            $memberMap[$a].BadLoanCount = [int]$rs.Fields["BadCnt"].Value
            $memberMap[$a].TotOWEBR = [double]$rs.Fields["TotOWEBR"].Value
            $memberMap[$a].TotOWEINT = [double]$rs.Fields["TotOWEINT"].Value
        }
        $rs.MoveNext()
    }
    $rs.Close()
    $rs = $db.OpenRecordset("SELECT ACCNO, Count(*) AS LoanCnt, Sum(IIf(IsNull(ALLN),0,ALLN)) AS TotALLN FROM BOROW GROUP BY ACCNO")
    while (-not $rs.EOF) {
        $a = $rs.Fields["ACCNO"].Value
        if ($memberMap.ContainsKey($a)) {
            $memberMap[$a].LoanCount = [int]$rs.Fields["LoanCnt"].Value
        }
        $rs.MoveNext()
    }
    $rs.Close()
    $db.Close()
    Write-Host "  完成" -ForegroundColor DarkGray

    # ── Step 5: 近期交易 ─────────────────────────────────────────────────────────
    Write-Host "[5/8] 檢查近期交易..." -ForegroundColor Yellow
    $db = $dbe.OpenDatabase($CubPath, $false, $true, $connectStr)
    $rs = $db.OpenRecordset("SELECT ACCNO, Count(*) AS TxnCnt FROM LGR WHERE DAT >= '$bDate' AND DAT <= '$eDate' AND DCHK='Y' GROUP BY ACCNO")
    while (-not $rs.EOF) {
        $a = $rs.Fields["ACCNO"].Value
        if ($memberMap.ContainsKey($a)) { $memberMap[$a].RecentTxn = [int]$rs.Fields["TxnCnt"].Value }
        $rs.MoveNext()
    }
    $rs.Close()
    $db.Close()
    Write-Host "  完成" -ForegroundColor DarkGray

    # ── Step 6: 其他異常檢查 ───────────────────────────────────────────────────
    Write-Host "[6/8] 其他異常檢查..." -ForegroundColor Yellow

    function Convert-ROCDate($s) {
        if (-not $s) { return $null }
        $s = "$s" -replace '/', ''
        if ($s.Length -lt 7) { return $null }
        try { return [datetime]::new([int]$s.Substring(0,3)+1911, [int]$s.Substring(3,2), [int]$s.Substring(5,2)) }
        catch { return $null }
    }

    # 7D: 同地址多戶貸款（>=3 戶）
    try {
        $db = $dbe.OpenDatabase($CubPath, $false, $true, $connectStr)
        $rs = $db.OpenRecordset("SELECT ADDR FROM SER WHERE (OUTDAT IS NULL OR OUTDAT='') AND IIf(IsNull(ADDR),'',ADDR)<>'' GROUP BY ADDR HAVING Count(*) >= 3")
        $badAddr = @{}
        while (-not $rs.EOF) {
            $badAddr[$rs.Fields("ADDR").Value] = $true
            $rs.MoveNext()
        }
        $rs.Close()
        $db.Close()
        foreach ($a in $memberMap.Keys) {
            if ($badAddr.ContainsKey($memberMap[$a].ADDR)) { $memberMap[$a].HasSameAddressMultiple = $true }
        }
        Write-Host "  7D 同地址>=3: 完成" -ForegroundColor DarkGray
    }
    catch { Write-Host "  7D 同地址: 無法" -ForegroundColor DarkGray }

    # 7E: 非社員貸款（BOROW 中有但 SER 無此人）
    try {
        $db = $dbe.OpenDatabase($CubPath, $false, $true, $connectStr)
        $rs = $db.OpenRecordset("SELECT DISTINCT B.ACCNO FROM BOROW B WHERE B.ACCNO NOT IN (SELECT S.ACCNO FROM SER S)")
        $nonMember = @{}
        while (-not $rs.EOF) {
            $nonMember[$rs.Fields("ACCNO").Value] = $true
            $rs.MoveNext()
        }
        $rs.Close()
        $db.Close()
        foreach ($a in $memberMap.Keys) {
            if ($nonMember.ContainsKey($a)) { $memberMap[$a].IsNonMember = $true }
        }
        Write-Host "  7E 非社員貸款: 完成" -ForegroundColor DarkGray
    }
    catch { Write-Host "  7E 非社員貸款: 無法" -ForegroundColor DarkGray }

    # 7G: 關係人放款(同GRPNO)
    try {
        $db = $dbe.OpenDatabase($CubPath, $false, $true, $connectStr)
        $rs = $db.OpenRecordset("SELECT DISTINCT S.GRPNO FROM (SER S INNER JOIN BOROW B ON S.ACCNO = B.ACCNO) WHERE (S.OUTDAT IS NULL OR S.OUTDAT='') AND IIf(IsNull(S.GRPNO),'',S.GRPNO)<>'' AND IIf(IsNull(B.ALLN),0,B.ALLN)>0")
        $grpLoan = @{}
        while (-not $rs.EOF) { $grpLoan[$rs.Fields("GRPNO").Value] = $true; $rs.MoveNext() }
        $rs.Close(); $db.Close()
        foreach ($a in $memberMap.Keys) {
            $g = $memberMap[$a].GRPNO
            if ($g -and $grpLoan.ContainsKey($g)) { $memberMap[$a].HasRelatedPartyLoan = $true }
        }
        Write-Host "  7G 關係人放款: 完成" -ForegroundColor DarkGray
    }
    catch { Write-Host "  7G 關係人放款: 無法" -ForegroundColor DarkGray }

    # 7I: 新入社立即貸款
    foreach ($a in $memberMap.Keys) {
        $m = $memberMap[$a]
        $indat = $m.INDAT
        if ($indat -ge $bDate -and $indat -le $eDate -and $indat -ne '' -and $m.LoanCount -gt 0) {
            $m.HasNewJoinLoan = $true
        }
    }
    Write-Host "  7I 新入社立即貸款: 完成" -ForegroundColor DarkGray

    # ══════════════════════════════════════════════════════════════════════
    # 新增強異常檢查 7J~7T
    # ══════════════════════════════════════════════════════════════════════

    $db = $dbe.OpenDatabase($CubPath, $false, $true, $connectStr)

    # 撈取 BOROW 明細資料（供 7O 滯納天數檢查 + 7Q 重複交易使用）
    try {
        $rs = $db.OpenRecordset("SELECT ACCNO, BRNO, SEPNO, MNNUM, PUNRAT, DLYRAT, COUNCILDAT, NOPUN, DLY, ORGMNY, ALLN, LASDAT, DAT FROM BOROW WHERE IIf(IsNull(ALLN),0,ALLN) > 0")
        $borowDetail = @{}
        while (-not $rs.EOF) {
            $a = $rs.Fields("ACCNO").Value
            if (-not $memberMap.ContainsKey($a)) { $rs.MoveNext(); continue }

            $brno = [string]$rs.Fields("BRNO").Value
            $dly = [int](Get-MemberValue $rs.Fields("DLY").Value 0)

            # 7O: 滯納天數超過系統設定上限
            if ($dly -gt $滯納天數上限) {
                $memberMap[$a].HasExceedPayDate = $true
            }

            # HasLoanBeforeApproval: 先放後審
            $datVal = [string]$rs.Fields("DAT").Value
            $councilVal = [string]$rs.Fields("COUNCILDAT").Value
            if ([string]::IsNullOrWhiteSpace($councilVal)) {
                $memberMap[$a].HasLoanBeforeApproval = $true
            }
            elseif ($datVal -and $councilVal) {
                $datDt = Convert-ROCDate $datVal
                $councilDt = Convert-ROCDate $councilVal
                if ($datDt -and $councilDt -and $datDt -lt $councilDt) {
                    $memberMap[$a].HasLoanBeforeApproval = $true
                }
            }

            if (-not $borowDetail.ContainsKey($a)) { $borowDetail[$a] = @() }
            $borowDetail[$a] += @{
                BRNO = $brno
            }
            $rs.MoveNext()
        }
        $rs.Close()
        Write-Host "  7O 超過約定還款日期: 完成" -ForegroundColor DarkGray
    }
    catch { Write-Host "  7O 超過約定還款日期: 無法" -ForegroundColor DarkGray }

    # 7M: 借據與審查紀錄不符（REC_前端 vs IOU_前端）
    #     從 BOROW 檢查是否有 REC_前端/IOU_前端欄位
    try {
        $rs = $db.OpenRecordset("SELECT ACCNO FROM BOROW WHERE IIf(IsNull(REC_前端),'',REC_前端) <> IIf(IsNull(IOU_前端),'',IOU_前端) AND IIf(IsNull(ALLN),0,ALLN) > 0")
        while (-not $rs.EOF) {
            $a = $rs.Fields("ACCNO").Value
            if ($memberMap.ContainsKey($a)) { $memberMap[$a].HasRecDiscrepancy = $true }
            $rs.MoveNext()
        }
        $rs.Close()
        Write-Host "  7M 借據與審查不符: 完成" -ForegroundColor DarkGray
    }
    catch {
        # 若欄位不存在則略過
        Write-Host "  7M 借據與審查不符: 略過（無此欄位）" -ForegroundColor DarkGray
    }

    # 7N: 審查後逾期未入帳（COUNCILDAT 後超過 60 天仍未撥款）
    function Resolve-Date($s) {
        if (-not $s) { return $null }
        $m = [regex]::Match($s, '^(\d{3})/(\d{1,2})/(\d{1,2})$')
        if ($m.Success) { return [datetime]::new([int]$m.Groups[1].Value + 1911, [int]$m.Groups[2].Value, [int]$m.Groups[3].Value) }
        $m = [regex]::Match($s, '^(\d{3})(\d{2})(\d{2})$')
        if ($m.Success) { return [datetime]::new([int]$m.Groups[1].Value + 1911, [int]$m.Groups[2].Value, [int]$m.Groups[3].Value) }
        $dt = $null
        [datetime]::TryParseExact($s, @("yyyy/MM/dd","yyyy/M/d","yyyyMMdd"), $null, 'None', [ref]$dt) | Out-Null
        return $dt
    }

    try {
        $rs = $db.OpenRecordset("SELECT ACCNO, DAT, COUNCILDAT FROM BOROW WHERE IIf(IsNull(COUNCILDAT),'',COUNCILDAT) <> '' AND IIf(IsNull(ALLN),0,ALLN) > 0")
        $auditOverdue = @{}
        while (-not $rs.EOF) {
            $a = $rs.Fields("ACCNO").Value
            $datVal = [string]$rs.Fields("DAT").Value
            $councilVal = [string]$rs.Fields("COUNCILDAT").Value
            if ($datVal -and $councilVal) {
                $datDt = Resolve-Date $datVal
                $councilDt = Resolve-Date $councilVal
                if (-not $datDt -or -not $councilDt) {
                    Write-Host "  ⚠ 日期解析失敗: ACCNO=$a, DAT=$datVal, COUNCILDAT=$councilVal" -ForegroundColor Yellow
                    $rs.MoveNext(); continue
                }
                # 審查後超過 2 個月仍未撥款
                if (($datDt - $councilDt).TotalDays -gt 60) {
                    $auditOverdue[$a] = $true
                }
            }
            $rs.MoveNext()
        }
        $rs.Close()
        foreach ($a in $auditOverdue.Keys) {
            if ($memberMap.ContainsKey($a)) { $memberMap[$a].HasAuditOverdue = $true }
        }
        Write-Host "  7N 審查逾期未入帳: 完成" -ForegroundColor DarkGray
    }
    catch { Write-Host "  7N 審查逾期未入帳: 無法" -ForegroundColor DarkGray }

    $db.Close()

    # 7P: 董監事貸款超限（TYPE='1' 且貸款總額 > 2200萬）
    try {
        foreach ($a in $memberMap.Keys) {
            $m = $memberMap[$a]
            if ($m.TYPE -eq '1' -and $m.LoanCount -gt 0 -and $m.LGR_Loan -gt 22000000) {
                $m.HasDirectorOverLimit = $true
            }
        }
        Write-Host "  7P 董監事貸款超限: 完成" -ForegroundColor DarkGray
    }
    catch { Write-Host "  7P 董監事貸款超限: 無法" -ForegroundColor DarkGray }

    # 7Q: 重複交易檢測（同 BRNO 有多筆借據）
    try {
        # 計算每組 BRNO 的總筆數（跳過空 BRNO）
        $brnoCount = @{}
        foreach ($a in $borowDetail.Keys) {
            foreach ($b in $borowDetail[$a]) {
                $br = $b.BRNO
                if ([string]::IsNullOrWhiteSpace($br)) { continue }
                if (-not $brnoCount.ContainsKey($br)) { $brnoCount[$br] = 0 }
                $brnoCount[$br]++
            }
        }
        foreach ($a in $memberMap.Keys) {
            if ($memberMap[$a].HasDuplicateLoan) { continue }
            if ($borowDetail.ContainsKey($a)) {
                foreach ($b in $borowDetail[$a]) {
                    $br = $b.BRNO
                    if (-not [string]::IsNullOrWhiteSpace($br) -and $brnoCount.ContainsKey($br) -and $brnoCount[$br] -gt 1) {
                        $memberMap[$a].HasDuplicateLoan = $true
                        break
                    }
                }
            }
        }
        Write-Host "  7Q 重複交易: 完成" -ForegroundColor DarkGray
    }
    catch { Write-Host "  7Q 重複交易: 無法" -ForegroundColor DarkGray }

    # ══════════════════════════════════════════════════════════════════════
    # 新增 7S~7U：專職挪用高風險指標
    # ══════════════════════════════════════════════════════════════════════

    $db = $dbe.OpenDatabase($CubPath, $false, $true, $connectStr)

    # 7S: 休眠戶激活（>365天無交易後突然在查核期活動）
    try {
        Write-Host "  7S 休眠戶激活分析..." -NoNewline
        $bDt = Convert-ROCDate $bDate
        $rs = $db.OpenRecordset("SELECT ACCNO, Max(DAT) AS LastTxn FROM LGR WHERE DAT < '$bDate' AND DCHK='Y' GROUP BY ACCNO")
        while (-not $rs.EOF) {
            $a = $rs.Fields("ACCNO").Value
            if ($memberMap.ContainsKey($a)) {
                $lt = Convert-ROCDate ([string]$rs.Fields("LastTxn").Value)
                $memberMap[$a].LastTxnBeforePeriod = $lt
                if ($lt -and $bDt -and ($bDt - $lt).TotalDays -gt 365 -and $memberMap[$a].RecentTxn -gt 0) {
                    $memberMap[$a].HasDormantActivation = $true
                }
            }
            $rs.MoveNext()
        }
        $rs.Close()
        Write-Host "完成" -ForegroundColor DarkGray
    }
    catch { Write-Host "無法: $_" -ForegroundColor DarkGray }

    # 7T: 整數金額比例（查核期內交易 >=50% 為整千或整萬）
    try {
        Write-Host "  7T 整數金額比例分析..." -NoNewline
        $rs = $db.OpenRecordset("SELECT ACCNO, Count(*) AS Ttl, Sum(IIf(CLng(MNY) Mod 1000 = 0 Or CLng(MNY) Mod 10000 = 0, 1, 0)) AS Rnd FROM LGR WHERE DAT >= '$bDate' AND DAT <= '$eDate' AND DCHK='Y' GROUP BY ACCNO")
        while (-not $rs.EOF) {
            $a = $rs.Fields("ACCNO").Value
            if ($memberMap.ContainsKey($a)) {
                $ttl = [int]$rs.Fields("Ttl").Value
                $rnd = [int]$rs.Fields("Rnd").Value
                if ($ttl -gt 0 -and ($rnd / $ttl) -ge 0.5) {
                    $memberMap[$a].HasRoundAmountTxn = $true
                }
            }
            $rs.MoveNext()
        }
        $rs.Close()
        Write-Host "完成" -ForegroundColor DarkGray
    }
    catch { Write-Host "無法: $_" -ForegroundColor DarkGray }

    # 7U: 新帳戶短期爆發（查核期內首次交易，3天內≥3筆）
    try {
        Write-Host "  7U 新帳戶短期爆發分析..." -NoNewline
        $rs = $db.OpenRecordset("SELECT ACCNO, Min(DAT) AS FirstTxn, Max(DAT) AS LastTxn, Count(*) AS Cnt FROM LGR WHERE DAT >= '$bDate' AND DAT <= '$eDate' AND DCHK='Y' GROUP BY ACCNO")
        while (-not $rs.EOF) {
            $a = $rs.Fields("ACCNO").Value
            if ($memberMap.ContainsKey($a)) {
                # 期前已有交易者非新帳戶，不標
                if ($memberMap[$a].LastTxnBeforePeriod -ne $null) { $rs.MoveNext(); continue }
                $cnt = [int]$rs.Fields("Cnt").Value
                $ft = Convert-ROCDate ([string]$rs.Fields("FirstTxn").Value)
                $lt = Convert-ROCDate ([string]$rs.Fields("LastTxn").Value)
                if ($cnt -ge 3 -and $ft -and $lt -and ($lt -$ft).TotalDays -le 3) {
                    $memberMap[$a].HasNewAccountBurst = $true
                }
            }
            $rs.MoveNext()
        }
        $rs.Close()
        Write-Host "完成" -ForegroundColor DarkGray
    }
    catch { Write-Host "無法: $_" -ForegroundColor DarkGray }

    $db.Close()

    # ── 補預設值 ──────────────────────────────────────────────────────
    foreach ($a in $memberMap.Keys) {
        if (-not $memberMap[$a].ContainsKey("LoanCount")) { $memberMap[$a].LoanCount = 0 }
        if (-not $memberMap[$a].ContainsKey("RecentTxn")) { $memberMap[$a].RecentTxn = 0 }
    }
    Write-Host "  其他異常檢查完成" -ForegroundColor DarkGray

    # ── Step 7: 計算異常分數 ─────────────────────────────────────────────────────────
    Write-Host "[7/8] 計算異常分數..." -ForegroundColor Yellow

    $results = @()
    foreach ($a in $memberMap.Keys) {
        $m = $memberMap[$a]
        $s = Get-AnomalyScore -Member $m
        if ($s.Score -eq 0) { continue }

        $results += [PSCustomObject]@{
            AccNo             = $a
            Name1             = $m.ACCNM
            Name2             = $m.NAM
            ADDR              = $m.ADDR
            TEL               = $m.TEL
            Score             = $s.Score
            Flags_cn          = $s.Flags_cn
            LGR_Share         = $m.LGR_Share
            SER_Share         = $m.SER_Share
            DiffShare         = $s.DiffShare
            LGR_Loan          = $m.LGR_Loan
            SER_Loan          = $m.SER_Loan
            DiffLoan          = $s.DiffLoan
            LGR_Reserve       = $m.LGR_Reserve
            SER_Reserve       = $m.SER_Reserve
            DiffReserve       = $s.DiffReserve
            BadLoans          = $m.BadLoanCount
            BadPrincipal      = $m.TotOWEBR
            BadInterest       = $m.TotOWEINT
            RecentTxn         = $m.RecentTxn
            LoanCount         = $m.LoanCount
            NonMember         = if ($m.IsNonMember) { "Y" } else { "" }
            SameAddrMul       = if ($m.HasSameAddressMultiple) { "Y" } else { "" }
            RelatedParty      = if ($m.HasRelatedPartyLoan) { "Y" } else { "" }
            NewJoinLoan       = if ($m.HasNewJoinLoan) { "Y" } else { "" }
            RecDiscrepancy    = if ($m.HasRecDiscrepancy) { "Y" } else { "" }
            AuditOverdue      = if ($m.HasAuditOverdue) { "Y" } else { "" }
            ExceedPayDate     = if ($m.HasExceedPayDate) { "Y" } else { "" }
            DirectorOverLimit = if ($m.HasDirectorOverLimit) { "Y" } else { "" }
            DuplicateLoan     = if ($m.HasDuplicateLoan) { "Y" } else { "" }
            DormantActivation = if ($m.HasDormantActivation) { "Y" } else { "" }
            RoundAmountTxn    = if ($m.HasRoundAmountTxn) { "Y" } else { "" }
            NewAccountBurst   = if ($m.HasNewAccountBurst) { "Y" } else { "" }
            NegativeLoanBalance = if ($m.HasNegativeLoanBalance) { "Y" } else { "" }
            LoanBeforeApproval  = if ($m.HasLoanBeforeApproval) { "Y" } else { "" }
        }
    }

    $results = $results | Sort-Object Score -Descending

    $outPath = Join-Path $PSScriptRoot (Get-ResultCsvName)
    $enc = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
    $results | Export-Csv -Path $outPath -Encoding $enc -NoTypeInformation

    Write-Host "`n════════════════════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  異常偵測統計" -ForegroundColor White
    Write-Host "  總會員: $memberCount" -ForegroundColor DarkGray
    Write-Host "  篩選出: $($results.Count)" -ForegroundColor Yellow
    Write-Host "  輸出CSV: $outPath" -ForegroundColor Green
    Write-Host "════════════════════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan

    Write-Host "`n=== 異常排名 TOP 20 ===" -ForegroundColor White
    Write-Host ("{0,-4} {1,-6} {2,-10} {3,4}  {4}" -f "排名", "會員", "戶名", "分數", "異常說明") -ForegroundColor DarkGray
    Write-Host ("-" * 80) -ForegroundColor DarkGray
    $rank = 1
    foreach ($r in $results | Select-Object -First 20) {
        $desc = $r.Flags_cn
        if ($desc.Length -gt 55) { $desc = $desc.Substring(0, 52) + "..." }
        Write-Host ("{0,3}. {1,-6} {2,-10} {3,4}  {4}" -f $rank, $r.AccNo, $r.Name1, $r.Score, $desc)
        $rank++
    }

    Write-Host "`n=== 建議追查 ===" -ForegroundColor White
    $high = @($results | Where-Object { (Get-ScoreCategory -Score $_.Score) -eq 'High' })
    $mid = @($results | Where-Object { (Get-ScoreCategory -Score $_.Score) -eq 'Mid' })
    $low = @($results | Where-Object { (Get-ScoreCategory -Score $_.Score) -eq 'Low' })
    Write-Host "  [立即追查] 分數>=10: $($high.Count) 人（高度挪用嫌疑）" -ForegroundColor Red
    Write-Host "  [優先追查] 分數5-9 : $($mid.Count) 人（中度挪用嫌疑）" -ForegroundColor Yellow
    Write-Host "  [建議追查] 分數1-4 : $($low.Count) 人（低度挪用嫌疑）" -ForegroundColor DarkGray
    Write-Host "  [免寄]    篩選無異常: $($memberCount - $results.Count) 人" -ForegroundColor DarkGray

    Write-Host "`n=== 開啟CSV 可搭配 Excel 開啟 ===" -ForegroundColor Green

    # ── Step 8: 寫入 CUB.MDB 的 M_對帳明細 ──────────────────────────────────────────
    Write-Host "`n[8/8] 寫入 CUB.MDB 的 M_對帳明細..." -ForegroundColor Yellow
    try {
        $dbCu = $dbe.OpenDatabase($CubPath, $false, $false, $connectStr)
        $rsCu = $dbCu.OpenRecordset("SELECT SRNO FROM PARA")
        $cuNo = $rsCu.Fields("SRNO").Value; $rsCu.Close()

        $hasTable = @($dbCu.TableDefs | Where-Object { $_.Name -eq "M_對帳明細" }).Count -gt 0
        if (-not $hasTable) {
            Write-Host "  M_對帳明細 不存在，建立中..." -ForegroundColor DarkGray
            $td = $dbCu.CreateTableDef("M_對帳明細")
            $td.Fields.Append($td.CreateField("基準日", 10, 7))
            $td.Fields.Append($td.CreateField("社號", 10, 6))
            $td.Fields.Append($td.CreateField("帳號", 10, 6))
            $td.Fields.Append($td.CreateField("姓名", 10, 10))
            $td.Fields.Append($td.CreateField("寄發", 1))
            $td.Fields.Append($td.CreateField("股金", 4))
            $td.Fields.Append($td.CreateField("s_YN", 1))
            $td.Fields.Append($td.CreateField("貸款", 4))
            $td.Fields.Append($td.CreateField("l_YN", 1))
            $td.Fields.Append($td.CreateField("備轉金", 4))
            $td.Fields.Append($td.CreateField("ps_YN", 1))
            $td.Fields.Append($td.CreateField("PS6", 4))
            $td.Fields.Append($td.CreateField("6_YN", 1))
            $td.Fields.Append($td.CreateField("Memo", 10, 50))
            $td.Fields.Append($td.CreateField("電話", 10, 34))
            $td.Fields.Append($td.CreateField("通訊處", 10, 60))
            $td.Fields.Append($td.CreateField("不寄發", 1))
            $td.Fields.Append($td.CreateField("不寄發原因", 10, 60))
            $td.Fields.Append($td.CreateField("原因提供者", 10, 20))
            $dbCu.TableDefs.Append($td)
        }

        $dbCu.BeginTrans()
        try {
            $dbCu.Execute("DELETE FROM [M_對帳明細] WHERE 基準日='$eDate' AND 社號='$cuNo'")

            $written = 0
            $rs = $dbCu.OpenRecordset("SELECT * FROM [M_對帳明細] WHERE False")
            foreach ($r in $results) {
                $rs.AddNew()
                $rs.Fields("基準日").Value = [string]$eDate
                $rs.Fields("社號").Value = [string]$cuNo
                $rs.Fields("帳號").Value = [string]$r.AccNo
                $rs.Fields("姓名").Value = if ($r.Name1) { [string]$r.Name1 } else { "" }
                $rs.Fields("寄發").Value = $true
                $rs.Fields("股金").Value = [double]$r.LGR_Share
                $rs.Fields("貸款").Value = [double]$r.LGR_Loan
                $rs.Fields("備轉金").Value = [double]$r.LGR_Reserve
                $rs.Fields("PS6").Value = [double]0
                $memo = if ($r.Flags_cn) { [string]$r.Flags_cn } else { "" }
                if ($memo.Length -gt 50) { $memo = $memo.Substring(0, 47) + "..." }
                $rs.Fields("Memo").Value = $memo
                $tel = if ($r.TEL) { [string]$r.TEL } else { "" }
                if ($tel.Length -gt 30) { $tel = $tel.Substring(0, 30) }
                $rs.Fields("電話").Value = $tel
                $rs.Update()
                $written++
            }
            $rs.Close()
            $dbCu.CommitTrans()
            Write-Host "  已寫入 $written 筆至 CUB.MDB 的 M_對帳明細" -ForegroundColor Green
            Write-Host "  開啟 Access → 對帳單作業 → 資料已自動載入。列印" -ForegroundColor Green
        }
        catch {
            $dbCu.Rollback()
            Write-Host "  寫入 CUB.MDB 的 M_對帳明細 失敗，已復原: $_" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "  寫入 CUB.MDB 的 M_對帳明細 失敗: $_" -ForegroundColor Red
    }

}
catch {
    Write-Host "`n錯誤: $_" -ForegroundColor Red
}
finally {
    if ($dbCu) { try { $dbCu.Close() } catch {} }
    if ($dbe) { [Runtime.InteropServices.Marshal]::ReleaseComObject($dbe) | Out-Null }
}

Write-Host "`n按任意鍵結束..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

