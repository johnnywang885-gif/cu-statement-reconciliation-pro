# generate_statements.ps1 — Word 套印對帳單（範本克隆版）
# 用法:
#   .\generate_statements.ps1                                    (自動找最新 CSV)
#   .\generate_statements.ps1 -CsvPath "CUB_異常社員_Top5Pct_xxx.csv"
#   .\generate_statements.ps1 -IndividualPdf                     (另產出個別 PDF)
#   .\generate_statements.ps1 -CubPassword "thifincub"           (指定 CUB.MDB 密碼)
# 前置: 須先執行篩選腳本產出 CSV
# 需求: Microsoft Word 已安裝（僅用於轉 PDF）、Access Database Engine 2016 (32-bit)

param(
    [string]$CsvPath = "",
    [string]$CubPath = "",
    [string]$CubPassword = "",
    [string]$CoopName = "",
    [string]$ReportDate = "",
    [switch]$IndividualPdf
)

Set-Location (Split-Path -Parent $PSCommandPath)

# ── 尋找 CSV ──────────────────────────────────────────────────────
if ([string]::IsNullOrEmpty($CsvPath)) {
    $patterns = @(
        "CUB_異常社員_Top*Pct_*.csv",
        "CUB_異常社員_篩選結果.csv",
        "CUB_異常社員_對帳單排序_*.csv"
    )
    foreach ($p in $patterns) {
        $f = Get-ChildItem $p -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($f) { $CsvPath = $f.FullName; break }
    }
}

if ([string]::IsNullOrEmpty($CsvPath) -or -not (Test-Path $CsvPath)) {
    Write-Host "找不到 CSV 檔，請先執行篩選腳本或用 -CsvPath 指定" -ForegroundColor Red
    exit 1
}

Write-Host "讀取 CSV: $CsvPath" -ForegroundColor Yellow
$csvFile = Import-Csv $CsvPath -Encoding utf8
$validRows = @($csvFile | Where-Object { $_.AccNo -and $_.AccNo -ne "" })

if ($validRows.Count -eq 0) {
    Write-Host "CSV 無有效資料" -ForegroundColor Red
    exit 1
}
Write-Host "  共 $($validRows.Count) 筆社員" -ForegroundColor DarkGray

# ── 從 CUB.MDB 讀取合作社名與日期 ──────────────────────────────────
$coopName = ""
$reportDate = ""

if ($CubPath -eq "") {
    $CubPath = Join-Path $PSScriptRoot "CUB.MDB"
}
$CubPath = [System.IO.Path]::GetFullPath($CubPath)

$daoAvailable = $false
try { $null = New-Object -ComObject DAO.DBEngine.120; $daoAvailable = $true } catch {}

if ($daoAvailable -and (Test-Path $CubPath)) {
    try {
        $connectStr = if ($CubPassword) { ";PWD=$CubPassword" } else { "" }
        $dbe = New-Object -ComObject DAO.DBEngine.120
        $db = $dbe.OpenDatabase($CubPath, $false, $true, $connectStr)
        $rs = $db.OpenRecordset("SELECT Item, Para FROM k_para")
        $kPara = @{}
        while (-not $rs.EOF) { $kPara[$rs.Fields["Item"].Value] = $rs.Fields["Para"].Value; $rs.MoveNext() }
        $rs.Close()
        if ($kPara.ContainsKey("eDate") -and $kPara["eDate"]) { $reportDate = [string]$kPara["eDate"] }
        $rsName = $db.OpenRecordset("SELECT ACCNM FROM SER WHERE ACCNO='000000'")
        if ($rsName.EOF) {} else { $val = $rsName.Fields["ACCNM"].Value; if ($val) { $coopName = [string]$val.Trim() } }
        $rsName.Close()
        $db.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($dbe)
        $dbe = $null
    } catch {
        Write-Host "  DAO 直接讀取失敗: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

if ([string]::IsNullOrEmpty($coopName) -or [string]::IsNullOrEmpty($reportDate)) {
    $ps32 = "$env:SystemRoot\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    $infoScript = Join-Path $PSScriptRoot "read_cub_info.ps1"
    if ((Test-Path $ps32) -and (Test-Path $infoScript)) {
        Write-Host "  透過32-bit 讀取 CUB.MDB..." -ForegroundColor DarkGray
        $infoArgs = @('-ExecutionPolicy', 'Bypass', '-File', $infoScript, '-CubPath', $CubPath)
        if ($CubPassword) { $infoArgs += @('-CubPassword', $CubPassword) }
        $jsonRaw = & $ps32 @infoArgs 2>$null
        if ($jsonRaw) {
            $json = $jsonRaw | ConvertFrom-Json
            if ($json.coopName) { $coopName = $json.coopName }
            if ($json.rawDate) {
                $rawDate = $json.rawDate
                if ($rawDate -match '^\d{7}$') {
                    $y = $rawDate.Substring(0,3); $m = $rawDate.Substring(3,2); $d = $rawDate.Substring(5,2)
                    $month = [int]$m; $day = [int]$d
                    if ($month -ge 1 -and $month -le 12 -and $day -ge 1 -and $day -le 31) { $reportDate = "$y/$m/$d" }
                } elseif ($rawDate -match '^\d{8}$') {
                    $y = [int]$rawDate.Substring(0,4) - 1911; $m = $rawDate.Substring(4,2); $d = $rawDate.Substring(6,2)
                    $month = [int]$m; $day = [int]$d
                    if ($month -ge 1 -and $month -le 12 -and $day -ge 1 -and $day -le 31) { $reportDate = "$y/$m/$d" }
                } else { $reportDate = $rawDate }
            }
        }
    }
}

if (-not [string]::IsNullOrEmpty($reportDate) -and $reportDate -notmatch '/') {
    $rawDate = $reportDate
    if ($rawDate -match '^\d{7}$') {
        $y = $rawDate.Substring(0,3); $m = $rawDate.Substring(3,2); $d = $rawDate.Substring(5,2)
        $month = [int]$m; $day = [int]$d
        if ($month -ge 1 -and $month -le 12 -and $day -ge 1 -and $day -le 31) { $reportDate = "$y/$m/$d" }
    } elseif ($rawDate -match '^\d{8}$') {
        $y = [int]$rawDate.Substring(0,4) - 1911; $m = $rawDate.Substring(4,2); $d = $rawDate.Substring(6,2)
        $month = [int]$m; $day = [int]$d
        if ($month -ge 1 -and $month -le 12 -and $day -ge 1 -and $day -le 31) { $reportDate = "$y/$m/$d" }
    }
}

if (-not [string]::IsNullOrEmpty($CoopName)) { $coopName = $CoopName }

# 非互動情境下 Read-Host 會造成「看似卡住」，改為自動 fallback
if ([string]::IsNullOrEmpty($reportDate) -and [string]::IsNullOrEmpty($ReportDate)) {
    $now = Get-Date
    $rocY = $now.Year - 1911
    $reportDate = "$rocY/$($now.ToString('MM/dd'))"
    Write-Host "  警告：無法從 CUB.MDB 讀取日期，改用今天 $reportDate" -ForegroundColor Yellow
} elseif (-not [string]::IsNullOrEmpty($reportDate) -and [string]::IsNullOrEmpty($ReportDate)) {
    $ReportDate = $reportDate
} elseif ([string]::IsNullOrEmpty($ReportDate)) {
    $ReportDate = $reportDate
}
if ([string]::IsNullOrEmpty($ReportDate)) {
    $ReportDate = Read-Host "無法從 CUB.MDB 讀取資料時點，請輸入日期（如 115/8/24）"
    if ([string]::IsNullOrEmpty($ReportDate)) { Write-Host "  未輸入日期，終止" -ForegroundColor Red; exit 1 }
}
Write-Host "  資料時點: $ReportDate" -ForegroundColor DarkGray

if ([string]::IsNullOrEmpty($coopName)) {
    Write-Host "  警告：無法從 CUB.MDB 讀取合作社名，改用預設" -ForegroundColor Yellow
    $coopName = "南投縣十方儲蓄互助社"
}
Write-Host "  合作社名: $coopName" -ForegroundColor DarkGray

# ── 輔助函式 ──────────────────────────────────────────────────────
function Get-FullCoopName {
    param([string]$ShortName)
    $s = $ShortName.Trim()
    # 移除已有的前後綴再重組，確保一律為 南投縣{短名}儲蓄互助社
    $s = $s -replace "^南投縣", "" -replace "儲蓄互助社$", ""
    $s = $s.Trim()
    if ([string]::IsNullOrEmpty($s)) { $s = $ShortName.Trim() }
    return "南投縣${s}儲蓄互助社"
}
function Sanitize-Filename {
    param([string]$Name)
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $result = $Name
    foreach ($c in $invalid) { $result = $result.Replace([string]$c, '_') }
    return $result
}

# ── Preflight：清理殘留 Word / 鎖定檔 ─────────────────────────────
function Clear-StaleWordLocks {
    try {
        Get-ChildItem -Path $PSScriptRoot -Filter "~`$*.docx" -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        Get-ChildItem -Path (Join-Path $PSScriptRoot "對帳單") -Filter "~`$*" -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        $asdDir = Join-Path $env:APPDATA "Microsoft\Word"
        if (Test-Path $asdDir) {
            Get-ChildItem -Path $asdDir -Filter "~WRL*.tmp" -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
            Get-ChildItem -Path $asdDir -Filter "*.asd" -Force -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddHours(-1) } | Remove-Item -Force -ErrorAction SilentlyContinue
        }
    } catch {}
}
function Stop-StaleWord {
    try {
        $procs = Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" -ErrorAction SilentlyContinue
        foreach ($p in $procs) {
            $cmd = $p.CommandLine
            $isAutomation = $cmd -match "/Automation"
            $hasVisibleTitle = $false
            try {
                $procObj = Get-Process -Id $p.ProcessId -ErrorAction SilentlyContinue
                if ($procObj -and $procObj.MainWindowTitle) { $hasVisibleTitle = $true }
            } catch {}
            if ($isAutomation -or -not $hasVisibleTitle) {
                Write-Host "  清理殘留 Word 行程 PID $($p.ProcessId)" -ForegroundColor DarkGray
                try { Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue } catch {}
            }
        }
        Start-Sleep -Seconds 1
    } catch {}
}
Clear-StaleWordLocks
Stop-StaleWord

# ── 範本克隆產生 DOCX（Open XML）───────────────────────────────
$wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
$w14Ns = "http://schemas.microsoft.com/office/word/2010/wordml"

function New-ParaId {
    $chars = "0123456789ABCDEF"
    $id = ""
    for ($i=0; $i -lt 8; $i++) { $id += $chars[(Get-Random -Maximum 16)] }
    return $id
}

function Set-ParaText {
    param($pNode, [string]$newText, $docXml, $nsMgr)
    # 保留 pPr，重置所有 w:r / w:hyperlink 等，改為單一 w:r
    $pPr = $pNode.SelectSingleNode("w:pPr", $nsMgr)
    $firstR = $pNode.SelectSingleNode("w:r", $nsMgr)
    $origRPr = $null
    if ($firstR) { $origRPr = $firstR.SelectSingleNode("w:rPr", $nsMgr) }
    $toRemove = @()
    foreach ($ch in $pNode.ChildNodes) {
        if ($ch.LocalName -ne "pPr") { $toRemove += $ch }
    }
    foreach ($c in $toRemove) { $null = $pNode.RemoveChild($c) }
    $newR = $docXml.CreateElement("w:r", $wNs)
    if ($origRPr) { $null = $newR.AppendChild($origRPr.CloneNode($true)) }
    $newT = $docXml.CreateElement("w:t", $wNs)
    if ($newText -match "^\s" -or $newText -match "\s$") {
        $null = $newT.SetAttribute("xml:space", "preserve")
    }
    $null = $newT.AppendChild($docXml.CreateTextNode($newText))
    $null = $newR.AppendChild($newT)
    $null = $pNode.AppendChild($newR)
}

function Set-TableCellText {
    param($tblNode, [int]$rowIdx, [int]$colIdx, [string]$newText, $docXml, $nsMgr)
    # rowIdx/colIdx 為 0-based
    $trs = $tblNode.SelectNodes("w:tr", $nsMgr)
    if ($rowIdx -ge $trs.Count) { return }
    $tr = $trs[$rowIdx]
    $tcs = $tr.SelectNodes("w:tc", $nsMgr)
    if ($colIdx -ge $tcs.Count) { return }
    $tc = $tcs[$colIdx]
    $p = $tc.SelectSingleNode("w:p", $nsMgr)
    if (-not $p) { return }
    Set-ParaText -pNode $p -newText $newText -docXml $docXml -nsMgr $nsMgr
    # 對齊：金額欄靠右，其餘置中（依範本已設，此處僅校正金額）
    if ($colIdx -eq 1 -and $rowIdx -gt 0) {
        $pPr = $p.SelectSingleNode("w:pPr", $nsMgr)
        if (-not $pPr) {
            $pPr = $docXml.CreateElement("w:pPr", $wNs)
            $null = $p.InsertBefore($pPr, $p.FirstChild)
        }
        $jc = $pPr.SelectSingleNode("w:jc", $nsMgr)
        if (-not $jc) { $jc = $docXml.CreateElement("w:jc", $wNs); $null = $pPr.AppendChild($jc) }
        $null = $jc.SetAttribute("val", $wNs, "right")
    }
}

function Add-KeepAndBreak {
    param($pNode, [bool]$isFirstOfBlock, [bool]$isLastOfBlock, $docXml, $nsMgr)
    $pPr = $pNode.SelectSingleNode("w:pPr", $nsMgr)
    if (-not $pPr) {
        $pPr = $docXml.CreateElement("w:pPr", $wNs)
        $pNode.InsertBefore($pPr, $pNode.FirstChild) | Out-Null
    }
    # keepLines 全加
    if (-not $pPr.SelectSingleNode("w:keepLines", $nsMgr)) {
        $kl = $docXml.CreateElement("w:keepLines", $wNs)
        $null = $pPr.AppendChild($kl)
    }
    # keepNext 除了塊內最後一段
    if (-not $isLastOfBlock) {
        if (-not $pPr.SelectSingleNode("w:keepNext", $nsMgr)) {
            $kn = $docXml.CreateElement("w:keepNext", $wNs)
            $null = $pPr.AppendChild($kn)
        }
    }
    # pageBreakBefore 僅塊首（第二塊起）
    if ($isFirstOfBlock) {
        if (-not $pPr.SelectSingleNode("w:pageBreakBefore", $nsMgr)) {
            $pb = $docXml.CreateElement("w:pageBreakBefore", $wNs)
            $null = $pPr.AppendChild($pb)
        }
    }
    # 更新 paraId 避免重複
    try {
        $null = $pNode.SetAttribute("paraId", $w14Ns, (New-ParaId))
        $null = $pNode.SetAttribute("textId", $w14Ns, "77777777")
    } catch {}
}

$templatePath = Join-Path $PSScriptRoot "對帳單列印範本.docx"
if (-not (Test-Path $templatePath)) {
    Write-Host "找不到範本: $templatePath" -ForegroundColor Red
    exit 1
}
$outDir = Join-Path $PSScriptRoot "對帳單"
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
$docxPath = Join-Path $outDir "對帳單_合併.docx"
$pdfPath  = Join-Path $outDir "對帳單_合併.pdf"

Write-Host "產生 DOCX（範本克隆）..." -ForegroundColor Yellow
$tmpRoot = Join-Path $env:TEMP ("opencode_stmt_{0}" -f [Guid]::NewGuid().ToString("N"))
$expandPath = Join-Path $tmpRoot "tpl"
New-Item -ItemType Directory -Path $expandPath -Force | Out-Null
Expand-Archive -LiteralPath $templatePath -DestinationPath $expandPath -Force

$docXmlPath = Join-Path $expandPath "word/document.xml"
[xml]$docXml = Get-Content -LiteralPath $docXmlPath -Raw -Encoding UTF8
$nsMgr = New-Object System.Xml.XmlNamespaceManager($docXml.NameTable)
$nsMgr.AddNamespace("w", $wNs)
$nsMgr.AddNamespace("w14", $w14Ns)

$body = $docXml.DocumentElement.SelectSingleNode("w:body", $nsMgr)
$sectPr = $body.SelectSingleNode("w:sectPr", $nsMgr)
if ($sectPr) { $null = $body.RemoveChild($sectPr) }

# 範本塊：27 p + 1 tbl = 28 節點（依範本實測）
$memberNodes = @()
foreach ($ch in $body.ChildNodes) { $memberNodes += $ch }
$body.RemoveAll()

Write-Host "  範本塊: $($memberNodes.Count) 節點 (含 tbl)" -ForegroundColor DarkGray

# longText 範本（P25, body index 24）
$longTextTemplate = "　　本對帳單乃由中華民國儲蓄互助協會督導組寄發，係為稽核儲蓄互助社之帳務及保障您的權益所做的例行查核對帳，您在 {0} 帳戶至 {1} 止各項餘額如下表。為維護您權益，請按您的存摺金額填入後，將此單以傳真方式或郵寄擲回。謝謝您的合作！"

$rowIdx = 0
foreach ($row in $validRows) {
    $rowIdx++
    if ($rowIdx % 10 -eq 1 -or $rowIdx -eq $validRows.Count) {
        Write-Host "    產生 $rowIdx/$($validRows.Count) ..." -ForegroundColor DarkGray
    }
    $addrRaw = if ($row.ADDR) { $row.ADDR.Trim() } else { "" }
    $postalCode = ""; $address = $addrRaw
    if ($addrRaw -match '^(\d{3,5})(.+)') { $postalCode = $Matches[1]; $address = $Matches[2].Trim() }
    $addrLine = if ($postalCode) { "$postalCode  $address" } else { $address }
    $name1 = if ($row.Name1) { $row.Name1.Trim() } else { "" }
    $nameLine = "$name1  君啟"
    $fullCoop = Get-FullCoopName $coopName
    $longText = $longTextTemplate -f $fullCoop, $ReportDate
    $accNo = if ($row.AccNo) { $row.AccNo } else { "" }
    $share = if ($row.LGR_Share) { [long]$row.LGR_Share } else { 0 }
    $loan  = if ($row.LGR_Loan) { [long]$row.LGR_Loan } else { 0 }
    $reserve = if ($row.LGR_Reserve) { [long]$row.LGR_Reserve } else { 0 }
    $shareStr = "{0:N0}" -f $share
    $loanStr  = "{0:N0}" -f $loan
    $reserveStr = "{0:N0}" -f $reserve
    $footerText = "帳號: $accNo                    社員:__________________(簽名或蓋章)             SN:    $rowIdx   "

    $isFirstBlock = ($rowIdx -eq 1)
    # 複製 28 節點
    $clones = @()
    for ($i=0; $i -lt $memberNodes.Count; $i++) {
        $orig = $memberNodes[$i]
        $clone = $orig.CloneNode($true)
        $clones += $clone
    }
    # 依索引替換（對應範本 body 0-based）
    # 9: 寄件者郵遞區號 404, 15: 收件者地址, 16: 姓名, 24: 長文, 27: 頁尾, 25: tbl 金額
    Set-ParaText -pNode $clones[9] -newText "404" -docXml $docXml -nsMgr $nsMgr
    Set-ParaText -pNode $clones[15] -newText $addrLine -docXml $docXml -nsMgr $nsMgr
    Set-ParaText -pNode $clones[16] -newText $nameLine -docXml $docXml -nsMgr $nsMgr
    Set-ParaText -pNode $clones[24] -newText $longText -docXml $docXml -nsMgr $nsMgr
    Set-ParaText -pNode $clones[27] -newText $footerText -docXml $docXml -nsMgr $nsMgr
    # tbl 金額：tbl 為 clones[25]
    $tblNode = $clones[25]
    Set-TableCellText -tblNode $tblNode -rowIdx 1 -colIdx 1 -newText $shareStr -docXml $docXml -nsMgr $nsMgr
    Set-TableCellText -tblNode $tblNode -rowIdx 2 -colIdx 1 -newText $loanStr -docXml $docXml -nsMgr $nsMgr
    Set-TableCellText -tblNode $tblNode -rowIdx 3 -colIdx 1 -newText $reserveStr -docXml $docXml -nsMgr $nsMgr
    # tbl cantSplit
    $trs = $tblNode.SelectNodes("w:tr", $nsMgr)
    foreach ($tr in $trs) {
        $trPr = $tr.SelectSingleNode("w:trPr", $nsMgr)
        if (-not $trPr) { $trPr = $docXml.CreateElement("w:trPr", $wNs); $tr.PrependChild($trPr) | Out-Null }
        if (-not $trPr.SelectSingleNode("w:cantSplit", $nsMgr)) {
            $cs = $docXml.CreateElement("w:cantSplit", $wNs)
            $trPr.AppendChild($cs) | Out-Null
        }
        # trHeight 已為 20，無需改
    }
    # keep / pageBreak：對塊內所有 p 加 keep，確保整塊不跨頁
    $pIndicesInBlock = @(0..24) + @(26,27) # 所有 p 的 clone 索引（跳過 tbl 25）
    for ($pi=0; $pi -lt $pIndicesInBlock.Count; $pi++) {
        $idx = $pIndicesInBlock[$pi]
        $isLast = ($pi -eq $pIndicesInBlock.Count -1)
        $addBreak = ($pi -eq 0 -and -not $isFirstBlock)
        Add-KeepAndBreak -pNode $clones[$idx] -isFirstOfBlock $addBreak -isLastOfBlock $isLast -docXml $docXml -nsMgr $nsMgr
    }

    foreach ($c in $clones) { $null = $body.AppendChild($c) }
}

# 加回 sectPr
if ($sectPr) { $null = $body.AppendChild($sectPr) }
$docXml.Save($docXmlPath)

# 重新打包為 DOCX
if (Test-Path $docxPath) { Remove-Item $docxPath -Force }
Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
[System.IO.Compression.ZipFile]::CreateFromDirectory($expandPath, $docxPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
# 修正屬性
try { (Get-Item $docxPath -Force).Attributes = 'Normal' } catch {}
Write-Host "Word 已儲存: $docxPath" -ForegroundColor Green

# 清理暫存
try { Remove-Item $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue } catch {}
# 確保檔案釋放且 DAO COM 已回收
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
Start-Sleep -Seconds 1

# ── 轉 PDF（Word COM 僅此處使用）─────────────────────────────
$word = $null
$doc = $null
$scriptSuccess = $false
try {
    Write-Host "轉 PDF..." -ForegroundColor Yellow
    # 先轉到本機暫存再搬回雲端硬碟，避免雲端同步鎖定
    $tmpPdf = Join-Path $env:TEMP ("stmt_{0}.pdf" -f [Guid]::NewGuid().ToString("N"))
    if (Test-Path $pdfPath) {
        try { Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue; Start-Sleep -Seconds 1 } catch {}
        try { Remove-Item $pdfPath -Force -ErrorAction SilentlyContinue } catch { Start-Sleep -Seconds 2; try { Remove-Item $pdfPath -Force -ErrorAction SilentlyContinue } catch {} }
    }
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try { $word.AutomationSecurity = 3 } catch {}
    try { $word.ScreenUpdating = $false } catch {}
    $doc = $word.Documents.Open($docxPath, $false, $false)
    try {
        $doc.SaveAs2($tmpPdf, 17)
    } catch {
        Write-Host "  SaveAs2 失敗，改用 ExportAsFixedFormat..." -ForegroundColor Yellow
        $doc.ExportAsFixedFormat($tmpPdf, 17)
    }
    $doc.Close(0)
    $doc = $null
    # 搬回正式位置（雲端硬碟可能鎖定，改用重試 + 改名備援）
    $finalPdf = $pdfPath
    $retry=0
    $moved=$false
    while ($retry -lt 5) {
        try {
            if (Test-Path $finalPdf) {
                try { Remove-Item $finalPdf -Force -ErrorAction Stop } catch {
                    # 若被雲端同步鎖定，改存為附檔名
                    $finalPdf = Join-Path $outDir ("對帳單_合併_{0}.pdf" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
                    Write-Host "  原 PDF 被佔用，改存為: $finalPdf" -ForegroundColor Yellow
                    break
                }
            }
            Move-Item -LiteralPath $tmpPdf -Destination $finalPdf -Force -ErrorAction Stop
            $moved=$true
            break
        } catch {
            $retry++
            Start-Sleep -Seconds 2
            if ($retry -ge 5) {
                $finalPdf = Join-Path $outDir ("對帳單_合併_{0}.pdf" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
                try { Move-Item -LiteralPath $tmpPdf -Destination $finalPdf -Force -ErrorAction Stop; $moved=$true; break } catch { throw }
            }
        }
    }
    if (-not $moved -and (Test-Path $tmpPdf)) {
        # 最後備援：直接複製
        try { Copy-Item -LiteralPath $tmpPdf -Destination $finalPdf -Force; $moved=$true } catch {}
    }
    if ($moved) { $pdfPath = $finalPdf }
    Write-Host "PDF 已儲存: $pdfPath" -ForegroundColor Green
    $scriptSuccess = $true
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  完成！共 $($validRows.Count) 筆社員" -ForegroundColor Cyan
    Write-Host "  合併 PDF: $pdfPath" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    # ── 個別 PDF（可選）────────────────────────────────────────
    if ($IndividualPdf) {
        Write-Host "產生個別 PDF..." -ForegroundColor Yellow
        # 個別改為同樣範本克隆單筆，逐筆轉 PDF
        $sn = 0
        foreach ($row in $validRows) {
            $sn++
            $tmpSingle = Join-Path $env:TEMP ("opencode_single_{0}_{1}" -f $sn, [Guid]::NewGuid().ToString("N"))
            New-Item -ItemType Directory -Path $tmpSingle -Force | Out-Null
            Expand-Archive -LiteralPath $templatePath -DestinationPath $tmpSingle -Force
            $sDocXmlPath = Join-Path $tmpSingle "word/document.xml"
            [xml]$sDocXml = Get-Content -LiteralPath $sDocXmlPath -Raw -Encoding UTF8
            $sNsMgr = New-Object System.Xml.XmlNamespaceManager($sDocXml.NameTable)
            $sNsMgr.AddNamespace("w", $wNs); $sNsMgr.AddNamespace("w14", $w14Ns)
            $sBody = $sDocXml.DocumentElement.SelectSingleNode("w:body", $sNsMgr)
            $sSect = $sBody.SelectSingleNode("w:sectPr", $sNsMgr)
            if ($sSect) { $null = $sBody.RemoveChild($sSect) }
            $sNodes = @($sBody.ChildNodes | ForEach-Object { $_ })
            $sBody.RemoveAll()
            # 單筆資料同上替換（複用邏輯）
            $addrRaw2 = if ($row.ADDR) { $row.ADDR.Trim() } else { "" }
            $postal2=""; $addr2=$addrRaw2
            if ($addrRaw2 -match '^(\d{3,5})(.+)') { $postal2=$Matches[1]; $addr2=$Matches[2].Trim() }
            $addrLine2 = if ($postal2) { "$postal2  $addr2" } else { $addr2 }
            $name2 = if ($row.Name1) { $row.Name1.Trim() } else { "" }
            $nameLine2 = "$name2  君啟"
            $fullCoop2 = Get-FullCoopName $coopName
            $long2 = $longTextTemplate -f $fullCoop2, $ReportDate
            $acc2 = if ($row.AccNo) { $row.AccNo } else { "" }
            $sShare = if ($row.LGR_Share) { [long]$row.LGR_Share } else { 0 }
            $sLoan  = if ($row.LGR_Loan) { [long]$row.LGR_Loan } else { 0 }
            $sRes   = if ($row.LGR_Reserve) { [long]$row.LGR_Reserve } else { 0 }
            $clones2 = @()
            foreach ($orig in $sNodes) { $clones2 += $orig.CloneNode($true) }
            Set-ParaText -pNode $clones2[9] -newText "404" -docXml $sDocXml -nsMgr $sNsMgr
            Set-ParaText -pNode $clones2[15] -newText $addrLine2 -docXml $sDocXml -nsMgr $sNsMgr
            Set-ParaText -pNode $clones2[16] -newText $nameLine2 -docXml $sDocXml -nsMgr $sNsMgr
            Set-ParaText -pNode $clones2[24] -newText $long2 -docXml $sDocXml -nsMgr $sNsMgr
            $footer2 = "帳號: $acc2                    社員:__________________(簽名或蓋章)             SN:    $sn   "
            Set-ParaText -pNode $clones2[27] -newText $footer2 -docXml $sDocXml -nsMgr $sNsMgr
            $tbl2 = $clones2[25]
            Set-TableCellText -tblNode $tbl2 -rowIdx 1 -colIdx 1 -newText ("{0:N0}" -f $sShare) -docXml $sDocXml -nsMgr $sNsMgr
            Set-TableCellText -tblNode $tbl2 -rowIdx 2 -colIdx 1 -newText ("{0:N0}" -f $sLoan) -docXml $sDocXml -nsMgr $sNsMgr
            Set-TableCellText -tblNode $tbl2 -rowIdx 3 -colIdx 1 -newText ("{0:N0}" -f $sRes) -docXml $sDocXml -nsMgr $sNsMgr
            foreach ($c in $clones2) { $null = $sBody.AppendChild($c) }
            if ($sSect) { $null = $sBody.AppendChild($sSect) }
            $sDocXml.Save($sDocXmlPath)
            $docxSingle = Join-Path $tmpSingle "single.docx"
            # 打包
            $tmpSingleExpanded = $tmpSingle
            # CreateFromDirectory 需要目錄內含 [Content_Types].xml 等
            $singleOut = Join-Path $outDir ("SN{0}_{1}_{2}.docx" -f $sn, $acc2, (Sanitize-Filename $name2))
            if (Test-Path $singleOut) { Remove-Item $singleOut -Force }
            [System.IO.Compression.ZipFile]::CreateFromDirectory($tmpSingleExpanded, $singleOut, [System.IO.Compression.CompressionLevel]::Optimal, $false)
            # 轉 PDF
            $singleDoc = $word.Documents.Open($singleOut, $false, $true)
            $singlePdf = [System.IO.Path]::ChangeExtension($singleOut, ".pdf")
            $singleDoc.SaveAs2($singlePdf, 17)
            $singleDoc.Close(0)
            try { Remove-Item $tmpSingle -Recurse -Force -ErrorAction SilentlyContinue } catch {}
            if ($sn % 20 -eq 0) { [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
        }
        Write-Host "  個別 PDF 已產出至 $outDir" -ForegroundColor Green
    }
} catch {
    Write-Host "錯誤: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
} finally {
    if ($doc) { try { $doc.Close(0) } catch {} }
    if ($word) { try { $word.Quit() } catch {}; [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    if ($dbe) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($dbe) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($scriptSuccess) { Write-Host "Word 已關閉" -ForegroundColor DarkGray }
}

