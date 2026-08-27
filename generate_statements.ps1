# generate_statements.ps1 — Word 套印對帳單，一次性產出 PDF
# 用法:
#   .\generate_statements.ps1                                    (自動找最新 CSV)
#   .\generate_statements.ps1 -CsvPath "CUB_異常社員_Top5Pct_xxx.csv"
#   .\generate_statements.ps1 -IndividualPdf                     (另產出個別 PDF)
#   .\generate_statements.ps1 -CubPassword "thifincub"           (指定 CUB.MDB 密碼)
# 前置: 須先執行篩選腳本產出 CSV
# 需求: Microsoft Word 已安裝、Access Database Engine 2016 (32-bit)

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

if ([string]::IsNullOrEmpty($ReportDate)) {
    $ReportDate = Read-Host "無法從 CUB.MDB 讀取資料時點，請輸入日期（如 115/8/24）"
}
Write-Host "  資料時點: $ReportDate" -ForegroundColor DarkGray

if ([string]::IsNullOrEmpty($coopName) -and [string]::IsNullOrEmpty($CoopName)) {
    $coopName = Read-Host "無法從 CUB.MDB 讀取合作社名，請輸入合作社名稱（如 南投縣十方儲蓄互助社）"
}
if (-not [string]::IsNullOrEmpty($CoopName)) { $coopName = $CoopName }
Write-Host "  合作社名: $coopName" -ForegroundColor DarkGray

# ── 輔助函式 ──────────────────────────────────────────────────────
function Sanitize-Filename {
    param([string]$Name)
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $result = $Name
    foreach ($c in $invalid) { $result = $result.Replace([string]$c, '_') }
    return $result
}

function Format-Range {
    param($Range, [int]$FontSize=14, [string]$FontName="標楷體", [int]$Bold=0,
          [int]$Align=0, [int]$LineRule=4, [double]$LineVal=12,
          [int]$BorderBot=1)
    $Range.Font.Name = $FontName
    $Range.Font.Size = $FontSize
    $Range.Font.Bold = $Bold
    $Range.ParagraphFormat.Alignment = $Align
    $Range.ParagraphFormat.LineSpacingRule = $LineRule
    $Range.ParagraphFormat.LineSpacing = $LineVal
    $Range.ParagraphFormat.SpaceAfter = 0
    $Range.ParagraphFormat.SpaceBefore = 0
    if ($BorderBot) {
        $Range.Borders.Item(3).LineStyle = 1
        $Range.Borders.Item(3).LineWidth = 6
        $Range.Borders.Item(3).ColorIndex = 1
    }
}

# ── 產生 Word → PDF ──────────────────────────────────────────────
$word = $null
$doc = $null
$scriptSuccess = $false

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    $doc = $word.Documents.Add()
    $sel = $word.Selection

    $doc.PageSetup.TopMargin    = 28.35
    $doc.PageSetup.BottomMargin = 8.5
    $doc.PageSetup.LeftMargin   = 28.35
    $doc.PageSetup.RightMargin  = 28.35

    $rowIdx = 0
    foreach ($row in $validRows) {
        $rowIdx++

        $addrRaw  = if ($row.ADDR) { $row.ADDR.Trim() } else { "" }
        $postalCode = ""; $address = $addrRaw
        if ($addrRaw -match '^(\d{3,5})(.+)') {
            $postalCode = $Matches[1]; $address = $Matches[2].Trim()
        }
        $name1 = if ($row.Name1) { $row.Name1.Trim() } else { "" }

        # ── P1-P8: 空白區域 ──────────────────────────────────────
        for ($i = 0; $i -lt 8; $i++) {
            $sel.TypeParagraph()
            $pFmt = $sel.Paragraphs.Last.Range.ParagraphFormat
            $pFmt.LineSpacingRule = 4
            $pFmt.LineSpacing = 0.7
            $pFmt.SpaceAfter = 0
            $pFmt.SpaceBefore = 0
        }

        # ── P9: 第一條摺線 ───────────────────────────────────────
        $sel.TypeParagraph()
        $p9 = $sel.Paragraphs.Last.Range
        Format-Range $p9 -LineRule 5 -LineVal 1

        # ── P10-P14: 協會資訊（底線, exact 12pt）────────────────
        $assocLines = @("04", "台中市北區北平路一段33號", "中華民國儲蓄互助協會",
                        "電話：04-22917272~8603", "傳真：04-22936903")
        foreach ($line in $assocLines) {
            $sel.TypeText($line)
            $sel.TypeParagraph()
            Format-Range $sel.Paragraphs.Last.Range -LineRule 4 -LineVal 12 -BorderBot 1
        }

        # ── P15: 空行（auto 48, 底線）───────────────────────────
        $sel.TypeParagraph()
        Format-Range $sel.Paragraphs.Last.Range -LineRule 5 -LineVal 2.4 -BorderBot 1

        # ── P16: 收件人地址（標楷體 14pt, 置中, auto 48, 底線）──
        $addrLine = if ($postalCode) { "$postalCode  $address" } else { $address }
        $sel.TypeText($addrLine)
        $sel.TypeParagraph()
        Format-Range $sel.Paragraphs.Last.Range -Align 1 -LineRule 5 -LineVal 2.4 -BorderBot 1

        # ── P17: 姓名+君啟（標楷體 14pt 粗體, 置中, auto 48, 底線）
        $sel.TypeText("$name1  君啟")
        $sel.TypeParagraph()
        Format-Range $sel.Paragraphs.Last.Range -Align 1 -Bold 1 -LineRule 5 -LineVal 2.4 -BorderBot 1

        # ── P18: 空行（置中, 粗體, 標楷體, auto 48, 底線）──────
        $sel.TypeParagraph()
        Format-Range $sel.Paragraphs.Last.Range -Align 1 -Bold 1 -LineRule 5 -LineVal 2.4 -BorderBot 1

        # ── P19: 空行（exact 14pt, 粗體, 底線）──────────────────
        $sel.TypeParagraph()
        Format-Range $sel.Paragraphs.Last.Range -Bold 1 -LineRule 4 -LineVal 14 -BorderBot 1

        # ── P20: 空行（exact 14pt, 粗體, 底線）──────────────────
        $sel.TypeParagraph()
        Format-Range $sel.Paragraphs.Last.Range -Bold 1 -LineRule 4 -LineVal 14 -BorderBot 1

        # ── P21: 第二條摺線（exact 14pt, 粗體）─────────────────
        $sel.TypeParagraph()
        Format-Range $sel.Paragraphs.Last.Range -Bold 1 -LineRule 4 -LineVal 14 -BorderBot 0

        # ── P22: 空行（exact 14pt, 粗體, 底線）──────────────────
        $sel.TypeParagraph()
        Format-Range $sel.Paragraphs.Last.Range -Bold 1 -LineRule 4 -LineVal 14 -BorderBot 1

        # ── P23: 親愛的社員您好: ─────────────────────────────────
        $sel.TypeText("親愛的社員您好:")
        $sel.TypeParagraph()
        Format-Range $sel.Paragraphs.Last.Range -Bold 1 -LineRule 4 -LineVal 14 -BorderBot 1

        # ── P24: 信文正文 ─────────────────────────────────────────
        $sel.TypeText("　　本對帳單乃由中華民國儲蓄互助協會督導組寄發，係為稽核儲蓄互助社之帳務及保障您的權益所做的例行查核對帳，您在 $coopName 帳戶至 $reportDate 止各項餘額如下表。為維護您權益，請按您的存摺金額填入後，將此單以傳真方式或郵寄擲回。謝謝您的合作！")
        $sel.TypeParagraph()
        Format-Range $sel.Paragraphs.Last.Range -Bold 1 -LineRule 4 -LineVal 14 -BorderBot 1

        # ── 表格 ──────────────────────────────────────────────────
        $share   = if ($row.LGR_Share)   { [long]$row.LGR_Share   } else { 0 }
        $loan    = if ($row.LGR_Loan)    { [long]$row.LGR_Loan    } else { 0 }
        $Reserve = if ($row.LGR_Reserve) { [long]$row.LGR_Reserve } else { 0 }

        $tableData = @(
            @("股　金",    $share,   "□正確　□錯誤：實有金額　　　　　"),
            @("貸　款",    $loan,    "□正確　□錯誤：實有金額　　　　　"),
            @("備　轉　金", $Reserve, "□正確　□錯誤：實有金額　　　　　"),
            @("備　註",    "",       "")
        )

        $numRows = $tableData.Count + 1
        $sel.Tables.Add($sel.Range, $numRows, 3) | Out-Null
        $tbl = $sel.Tables(1)
        $tbl.Borders.Enable = 1
        $tbl.Borders.OutsideLineStyle = 1
        $tbl.Borders.InsideLineStyle   = 1
        $tbl.Columns(1).Width = 116
        $tbl.Columns(2).Width = 95
        $tbl.Columns(3).Width = 323

        # 表頭
        $tbl.Cell(1,1).Range.Text = "帳目"
        $tbl.Cell(1,2).Range.Text = "餘額"
        $tbl.Cell(1,3).Range.Text = "核 對 勘 誤 回 函 說 明"
        for ($c = 1; $c -le 3; $c++) {
            $cr = $tbl.Cell(1,$c).Range
            $cr.Font.Name = "標楷體"; $cr.Font.Bold = 1
            $cr.ParagraphFormat.LineSpacingRule = 5; $cr.ParagraphFormat.LineSpacing = 2.4
            $cr.ParagraphFormat.Alignment = 1
            $cr.Collapse(0) | Out-Null
        }

        # 資料列
        for ($i = 0; $i -lt $tableData.Count; $i++) {
            $r = $i + 2
            $val = $tableData[$i][1]
            $valStr = if ($val -is [string]) { $val } else { "{0:N0}" -f $val }
            $tbl.Cell($r,1).Range.Text = $tableData[$i][0]
            $tbl.Cell($r,2).Range.Text = $valStr
            $tbl.Cell($r,3).Range.Text = $tableData[$i][2]
            for ($c = 1; $c -le 3; $c++) {
                $cr = $tbl.Cell($r,$c).Range
                $cr.Font.Name = "標楷體"; $cr.Font.Bold = 1
                $cr.ParagraphFormat.LineSpacingRule = 5; $cr.ParagraphFormat.LineSpacing = 2.4
                if ($c -eq 2) { $cr.ParagraphFormat.Alignment = 2 }
                $cr.Collapse(0) | Out-Null
            }
        }

        $sel.EndOf(15) | Out-Null
        $sel.MoveDown() | Out-Null

        # ── P25: 表格後空行（底線, space=21, auto 48）───────────
        $sel.TypeParagraph()
        Format-Range $sel.Paragraphs.Last.Range -Bold 1 -LineRule 5 -LineVal 2.4 -BorderBot 1

        # ── P26: 頁尾 ────────────────────────────────────────────
        $accNo = if ($row.AccNo) { $row.AccNo } else { "" }
        $sel.TypeText("帳號: $accNo                    社員:__________________(簽名或蓋章)             SN:    $rowIdx   ")
        $sel.TypeParagraph()
        Format-Range $sel.Paragraphs.Last.Range -Bold 1 -LineRule 5 -LineVal 2.4 -BorderBot 1
    }

    # ── 存檔 ──────────────────────────────────────────────────────
    $outDir = Join-Path $PSScriptRoot "對帳單"
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

    $docxPath = Join-Path $outDir "對帳單_合併.docx"
    $pdfPath  = Join-Path $outDir "對帳單_合併.pdf"

    $doc.SaveAs2($docxPath, 16)
    $f = Get-Item $docxPath -Force
    $f.Attributes = 'Normal'
    Write-Host "Word 已儲存: $docxPath" -ForegroundColor Green

    # ── 後處理：透過 Word COM 設定 PageBreakBefore ──────────────────
    Write-Host "後處理：加入分頁符號..." -ForegroundColor Yellow

    $memberCount = 0
    $docEnd = $doc.Content.End
    for ($t = 2; $t -le $doc.Tables.Count; $t++) {
        $tblEnd = $doc.Tables($t).Range.End
        $rangeEnd = [Math]::Min($tblEnd + 5000, $docEnd)
        if ($tblEnd -ge $docEnd) { continue }
        $afterTbl = $doc.Range($tblEnd, $rangeEnd)
        if ($afterTbl.Paragraphs.Count -ge 3) {
            $p1Range = $afterTbl.Paragraphs(3).Range
            $p1Range.ParagraphFormat.PageBreakBefore = $true
            $memberCount++
        }
    }
    Write-Host "  已為 $memberCount 筆加入 PageBreakBefore" -ForegroundColor DarkGray

    # ── 收尾：讓最後一人的頁尾段落成為文件最後一段，消除多餘空頁 ──
    # Word 文件強制結尾必須有至少一個段落；若最後一段是空白，會被
    # 擠到最後一頁之後（多一個空白頁）。解法：刪除最後一人「帳號」
    # 頁尾之後的所有段落，使頁尾本身成為文件最後一段（與範本一致）。
    $lastIdx = $doc.Paragraphs.Count
    $footerIdx = $lastIdx
    $docEnd = $doc.Content.End
    for ($i = $lastIdx; $i -ge [Math]::Max(2, $lastIdx - 30); $i--) {
        $t = $doc.Paragraphs($i).Range.Text.TrimEnd("`r", "`n", [char]7)
        if ($t.Trim() -ne '') { $footerIdx = $i; break }
    }
    $footerEnd = $doc.Paragraphs($footerIdx).Range.End
    if ($footerEnd -lt $docEnd) {
        $trailing = $doc.Range($footerEnd, $docEnd)
        $null = $trailing.Delete()
    }

    # ── 儲存含分頁的 DOCX ─────────────────────────────────────────
    $doc.Save()
    Write-Host "DOCX 已更新（含分頁）: $docxPath" -ForegroundColor Green

    # ── 轉 PDF ─────────────────────────────────────────────────────
    $doc.SaveAs2($pdfPath, 17)
    Write-Host "PDF 已儲存: $pdfPath" -ForegroundColor Green

    # ── 個別 PDF（可選）────────────────────────────────────────────
    if ($IndividualPdf) {
        Write-Host "產生個別 PDF..." -ForegroundColor Yellow
        $sn = 0
        foreach ($row in $validRows) {
            $sn++
            $singleDoc = $word.Documents.Add()
            $singleSel = $word.Selection
            $singleDoc.PageSetup.TopMargin    = 28.35
            $singleDoc.PageSetup.BottomMargin = 8.5
            $singleDoc.PageSetup.LeftMargin   = 28.35
            $singleDoc.PageSetup.RightMargin  = 28.35

            $addrRaw2 = if ($row.ADDR) { $row.ADDR.Trim() } else { "" }
            $postalCode2 = ""; $address2 = $addrRaw2
            if ($addrRaw2 -match '^(\d{3,5})(.+)') { $postalCode2 = $Matches[1]; $address2 = $Matches[2].Trim() }
            $name2 = if ($row.Name1) { $row.Name1.Trim() } else { "" }

            for ($i = 0; $i -lt 8; $i++) { $singleSel.TypeParagraph() }

            $singleSel.TypeParagraph()
            Format-Range $singleSel.Paragraphs.Last.Range -LineRule 5 -LineVal 1

            foreach ($line in $assocLines) {
                $singleSel.TypeText($line)
                $singleSel.TypeParagraph()
                Format-Range $singleSel.Paragraphs.Last.Range -LineRule 4 -LineVal 12 -BorderBot 1
            }

            $singleSel.TypeParagraph()
            Format-Range $singleSel.Paragraphs.Last.Range -LineRule 5 -LineVal 2.4 -BorderBot 1

            $addrLine2 = if ($postalCode2) { "$postalCode2  $address2" } else { $address2 }
            $singleSel.TypeText($addrLine2)
            $singleSel.TypeParagraph()
            Format-Range $singleSel.Paragraphs.Last.Range -Align 1 -LineRule 5 -LineVal 2.4 -BorderBot 1

            $singleSel.TypeText("$name2  君啟")
            $singleSel.TypeParagraph()
            Format-Range $singleSel.Paragraphs.Last.Range -Align 1 -Bold 1 -LineRule 5 -LineVal 2.4 -BorderBot 1

            $singleSel.TypeParagraph()
            Format-Range $singleSel.Paragraphs.Last.Range -Align 1 -Bold 1 -LineRule 5 -LineVal 2.4 -BorderBot 1

            $singleSel.TypeParagraph()
            Format-Range $singleSel.Paragraphs.Last.Range -Bold 1 -LineRule 4 -LineVal 14 -BorderBot 1

            $singleSel.TypeParagraph()
            Format-Range $singleSel.Paragraphs.Last.Range -Bold 1 -LineRule 4 -LineVal 14 -BorderBot 1

            $singleSel.TypeParagraph()
            Format-Range $singleSel.Paragraphs.Last.Range -Bold 1 -LineRule 4 -LineVal 14 -BorderBot 0

            $singleSel.TypeParagraph()
            Format-Range $singleSel.Paragraphs.Last.Range -Bold 1 -LineRule 4 -LineVal 14 -BorderBot 1

            $singleSel.TypeText("親愛的社員您好:")
            $singleSel.TypeParagraph()
            Format-Range $singleSel.Paragraphs.Last.Range -Bold 1 -LineRule 4 -LineVal 14 -BorderBot 1

            $singleSel.TypeText("　　本對帳單乃由中華民國儲蓄互助協會督導組寄發，係為稽核儲蓄互助社之帳務及保障您的權益所做的例行查核對帳，您在 $coopName 帳戶至 $reportDate 止各項餘額如下表。為維護您權益，請按您的存摺金額填入後，將此單以傳真方式或郵寄擲回。謝謝您的合作！")
            $singleSel.TypeParagraph()
            Format-Range $singleSel.Paragraphs.Last.Range -Bold 1 -LineRule 4 -LineVal 14 -BorderBot 1

            $s2 = if ($row.LGR_Share)   { [long]$row.LGR_Share   } else { 0 }
            $l2 = if ($row.LGR_Loan)    { [long]$row.LGR_Loan    } else { 0 }
            $r2 = if ($row.LGR_Reserve) { [long]$row.LGR_Reserve } else { 0 }
            $td2 = @(
                @("股　金",    $s2, "□正確　□錯誤：實有金額　　　　　"),
                @("貸　款",    $l2, "□正確　□錯誤：實有金額　　　　　"),
                @("備　轉　金", $r2, "□正確　□錯誤：實有金額　　　　　"),
                @("備　註",    "",  "")
            )
            $nr2 = $td2.Count + 1
            $singleSel.Tables.Add($singleSel.Range, $nr2, 3) | Out-Null
            $st2 = $singleSel.Tables(1)
            $st2.Borders.Enable = 1
            $st2.Borders.OutsideLineStyle = 1
            $st2.Borders.InsideLineStyle   = 1
            $st2.Columns(1).Width = 116
            $st2.Columns(2).Width = 95
            $st2.Columns(3).Width = 323
            $st2.Cell(1,1).Range.Text = "帳目"
            $st2.Cell(1,2).Range.Text = "餘額"
            $st2.Cell(1,3).Range.Text = "核 對 勘 誤 回 函 說 明"
            for ($c = 1; $c -le 3; $c++) {
                $cr = $st2.Cell(1,$c).Range
                $cr.Font.Name = "標楷體"; $cr.Font.Bold = 1
                $cr.ParagraphFormat.LineSpacingRule = 5; $cr.ParagraphFormat.LineSpacing = 2.4
                $cr.ParagraphFormat.Alignment = 1; $cr.Collapse(0) | Out-Null
            }
            for ($i = 0; $i -lt $td2.Count; $i++) {
                $ri = $i + 2
                $vs = if ($td2[$i][1] -is [string]) { $td2[$i][1] } else { "{0:N0}" -f $td2[$i][1] }
                $st2.Cell($ri,1).Range.Text = $td2[$i][0]
                $st2.Cell($ri,2).Range.Text = $vs
                $st2.Cell($ri,3).Range.Text = $td2[$i][2]
                for ($c = 1; $c -le 3; $c++) {
                    $cr = $st2.Cell($ri,$c).Range
                    $cr.Font.Name = "標楷體"; $cr.Font.Bold = 1
                    $cr.ParagraphFormat.LineSpacingRule = 5; $cr.ParagraphFormat.LineSpacing = 2.4
                    if ($c -eq 2) { $cr.ParagraphFormat.Alignment = 2 }
                    $cr.Collapse(0) | Out-Null
                }
            }
            $singleSel.EndOf(15) | Out-Null
            $singleSel.MoveDown() | Out-Null

            $singleSel.TypeParagraph()
            Format-Range $singleSel.Paragraphs.Last.Range -Bold 1 -LineRule 5 -LineVal 2.4 -BorderBot 1

            $singleSel.TypeText("帳號: $($row.AccNo)                    社員:__________________(簽名或蓋章)             SN:    $sn   ")
            $singleSel.TypeParagraph()
            Format-Range $singleSel.Paragraphs.Last.Range -Bold 1 -LineRule 5 -LineVal 2.4 -BorderBot 1

            $acc = if ($row.AccNo) { $row.AccNo } else { "000000" }
            $safeName = Sanitize-Filename $name2
            $singlePdf = Join-Path $outDir "SN${sn}_${acc}_${safeName}.pdf"
            $singleDoc.SaveAs2($singlePdf, 17)
            $singleDoc.Close(0)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($singleDoc)

            if ($sn % 20 -eq 0) {
                [System.GC]::Collect()
                [System.GC]::WaitForPendingFinalizers()
            }
        }
        Write-Host "  個別 PDF 已產出至 $outDir" -ForegroundColor Green
    }

    $doc.Close(0)
    $doc = $null
    $scriptSuccess = $true

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  完成！共 $($validRows.Count) 筆社員" -ForegroundColor Cyan
    Write-Host "  合併 PDF: $pdfPath" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
}
catch {
    Write-Host "錯誤: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
}
finally {
    if ($doc)    { try { $doc.Close(0) } catch {} }
    if ($word)   {
        try { $word.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    if ($dbe) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($dbe) }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    if ($scriptSuccess) { Write-Host "Word 已關閉" -ForegroundColor DarkGray }
}


