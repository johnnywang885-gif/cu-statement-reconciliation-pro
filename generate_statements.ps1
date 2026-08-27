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

# 嘗試在當前 pwsh 直接用 DAO（32-bit 環境）
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

# DAO 不可用時（64-bit pwsh），呼叫32-bit read_cub_info.ps1
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

# 格式化日期（從 DAO 直接讀取時的 rawDate 處理）
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

# 若日期仍為空，詢問使用者
if ([string]::IsNullOrEmpty($ReportDate)) {
    $ReportDate = Read-Host "無法從 CUB.MDB 讀取資料時點，請輸入日期（如 115/8/24）"
}
Write-Host "  資料時點: $ReportDate" -ForegroundColor DarkGray

# 若合作社名仍為空，詢問使用者
if ([string]::IsNullOrEmpty($coopName) -and [string]::IsNullOrEmpty($CoopName)) {
    $coopName = Read-Host "無法從 CUB.MDB 讀取合作社名，請輸入合作社名稱（如 南投縣十方儲蓄互助社）"
}
if (-not [string]::IsNullOrEmpty($CoopName)) { $coopName = $CoopName }
Write-Host "  合作社名: $coopName" -ForegroundColor DarkGray

# ── 輔助函式 ──────────────────────────────────────────────────────
function Set-CellText {
    param($Table, [int]$Row, [int]$Col, [string]$Text, [int]$Bold = 0, [int]$Align = 0)
    $cell = $Table.Cell($Row, $Col)
    $r = $cell.Range
    $r.Text = $Text
    $r.Font.Name = "標楷體"
    $r.Font.Bold = 1
    if ($Align) { $r.ParagraphFormat.Alignment = $Align }
    $r.ParagraphFormat.LineSpacingRule = 5  # wdLineSpaceMultiple
    $r.ParagraphFormat.LineSpacing = 2.4    # 0.2 lines
    $r.Collapse(0) | Out-Null
}

function Write-FoldLine {
    param($Sel)
    $Sel.TypeParagraph()
    $pRange = $Sel.Paragraphs.Last.Range
    $pRange.Borders.Item(3).LineStyle = 1   # wdBorderBottom = 3, wdBorderSingle = 1
    $pRange.Borders.Item(3).LineWidth = 6   # 1.5pt
    $pRange.Borders.Item(3).ColorIndex = 1  # wdBlack
}

function Write-BlankArea {
    param($Sel, $Doc)
    # 範本：8 個空行（上方 1/3，三折第一折）
    for ($i = 0; $i -lt 8; $i++) { $Sel.TypeParagraph() }
}

function Write-AddressBlock {
    param($Sel, [string]$PostalCode, [string]$Address, [string]$Name)
    # 地址行（標楷體 14pt, 置中, 底線）
    $addrLine = if ($PostalCode) { "$PostalCode  $Address" } else { $Address }
    $Sel.TypeText($addrLine)
    $addrRange = $Sel.Paragraphs.Last.Range
    $addrRange.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
    $addrRange.Font.Name = "標楷體"
    $addrRange.Font.Size = 14
    $addrRange.Borders.Item(3).LineStyle = 1
    $addrRange.Borders.Item(3).LineWidth = 6
    $addrRange.Borders.Item(3).ColorIndex = 1
    $addrRange.ParagraphFormat.LineSpacingRule = 5  # wdLineSpaceMultiple
    $addrRange.ParagraphFormat.LineSpacing = 2.4    # 0.2 lines
    $Sel.TypeParagraph()
    # 姓名+君啟（標楷體 14pt 粗體, 置中, 底線）
    $Sel.TypeText("$Name  君啟")
    $nameRange = $Sel.Paragraphs.Last.Range
    $nameRange.ParagraphFormat.Alignment = 1
    $nameRange.Font.Name = "標楷體"
    $nameRange.Font.Size = 14
    $nameRange.Font.Bold = 1
    $nameRange.Borders.Item(3).LineStyle = 1
    $nameRange.Borders.Item(3).LineWidth = 6
    $nameRange.Borders.Item(3).ColorIndex = 1
    $nameRange.ParagraphFormat.LineSpacingRule = 5
    $nameRange.ParagraphFormat.LineSpacing = 2.4
    $Sel.TypeParagraph()
    # 空行（置中, 粗體, 底線）
    $Sel.Range.ParagraphFormat.Alignment = 1
    $Sel.TypeParagraph()
    $emptyRange = $Sel.Paragraphs.Last.Range
    $emptyRange.ParagraphFormat.Alignment = 1
    $emptyRange.Font.Name = "標楷體"
    $emptyRange.Font.Size = 14
    $emptyRange.Font.Bold = 1
    $emptyRange.Borders.Item(3).LineStyle = 1
    $emptyRange.Borders.Item(3).LineWidth = 6
    $emptyRange.Borders.Item(3).ColorIndex = 1
    $emptyRange.ParagraphFormat.LineSpacingRule = 5
    $emptyRange.ParagraphFormat.LineSpacing = 2.4
}

function Write-LetterBody {
    param($Sel, [string]$CoopName, [string]$Date)
    # 設定段落格式（標楷體 14pt 粗體, exact 14pt 行高, 底線）
    $Sel.TypeText("親愛的社員您好:")
    $p1 = $Sel.Paragraphs.Last.Range
    $p1.Font.Name = "標楷體"; $p1.Font.Size = 14; $p1.Font.Bold = 1
    $p1.Borders.Item(3).LineStyle = 1; $p1.Borders.Item(3).LineWidth = 6; $p1.Borders.Item(3).ColorIndex = 1
    $p1.ParagraphFormat.LineSpacingRule = 3; $p1.ParagraphFormat.LineSpacing = 14
    $Sel.TypeParagraph()
    $Sel.TypeText("　　本對帳單乃由中華民國儲蓄互助協會督導組寄發，係為稽核儲蓄互助社之帳務及保障您的權益所做的例行查核對帳，您在 $CoopName 帳戶至 $Date 止各項餘額如下表。為維護您權益，請按您的存摺金額填入後，將此單以傳真方式或郵寄擲回。謝謝您的合作！")
    $p2 = $Sel.Paragraphs.Last.Range
    $p2.Font.Name = "標楷體"; $p2.Font.Size = 14; $p2.Font.Bold = 1
    $p2.Borders.Item(3).LineStyle = 1; $p2.Borders.Item(3).LineWidth = 6; $p2.Borders.Item(3).ColorIndex = 1
    $p2.ParagraphFormat.LineSpacingRule = 3; $p2.ParagraphFormat.LineSpacing = 14
    $Sel.TypeParagraph()
}

function Write-StatementTable {
    param($Doc, $Sel, [long]$Share, [long]$Loan, [long]$Reserve)
    $tableData = @(
        @("股　金",    $Share,   "□正確　□錯誤：實有金額　　　　　"),
        @("貸　款",    $Loan,    "□正確　□錯誤：實有金額　　　　　"),
        @("備　轉　金", $Reserve, "□正確　□錯誤：實有金額　　　　　"),
        @("備　註",    "",       "")
    )

    $numRows = $tableData.Count + 1
    $Sel.Tables.Add($Sel.Range, $numRows, 3) | Out-Null
    $tbl = $Sel.Tables(1)
    $tbl.Borders.Enable = 1
    $tbl.Borders.OutsideLineStyle = 1
    $tbl.Borders.InsideLineStyle   = 1
    $tbl.Columns(1).Width = 116
    $tbl.Columns(2).Width = 95
    $tbl.Columns(3).Width = 323

    # 表頭（標楷體粗體）
    Set-CellText $tbl 1 1 "帳目"                   -Bold 1 -Align 1
    Set-CellText $tbl 1 2 "餘額"                   -Bold 1 -Align 1
    Set-CellText $tbl 1 3 "核 對 勘 誤 回 函 說 明" -Bold 1 -Align 1

    # 資料列（標楷體粗體）
    for ($i = 0; $i -lt $tableData.Count; $i++) {
        $r = $i + 2
        $val = $tableData[$i][1]
        $valStr = if ($val -is [string]) { $val } else { "{0:N0}" -f $val }
        Set-CellText $tbl $r 1 $tableData[$i][0] -Bold 1
        Set-CellText $tbl $r 2 $valStr -Bold 1 -Align 2
        Set-CellText $tbl $r 3 $tableData[$i][2] -Bold 1
    }

    $Sel.EndOf(15) | Out-Null
    $Sel.MoveDown() | Out-Null
    $Sel.TypeParagraph()
    # 表格後空行（底線 sz=6, space=21, 標楷體粗體）
    $afterTbl = $Sel.Paragraphs.Last.Range
    $afterTbl.Font.Name = "標楷體"; $afterTbl.Font.Bold = 1
    $afterTbl.Borders.Item(3).LineStyle = 1; $afterTbl.Borders.Item(3).LineWidth = 6; $afterTbl.Borders.Item(3).ColorIndex = 1
    $afterTbl.ParagraphFormat.LineSpacingRule = 5; $afterTbl.ParagraphFormat.LineSpacing = 2.4
}

function Write-AssociationInfo {
    param($Sel)
    $lines = @("404", "台中市北區北平路一段33號", "中華民國儲蓄互助協會",
               "電話：04-22917272~8603", "傳真：04-22936903")
    foreach ($line in $lines) {
        $Sel.TypeText($line)
        $Sel.TypeParagraph()
        $pRange = $Sel.Paragraphs.Last.Range
        # 底線（與範本一致：single, sz=6）
        $pRange.Borders.Item(3).LineStyle = 1
        $pRange.Borders.Item(3).LineWidth = 6
        $pRange.Borders.Item(3).ColorIndex = 1
        # 行高 exact 240 twips = 12pt
        $pRange.ParagraphFormat.LineSpacingRule = 3  # wdLineSpaceExactly
        $pRange.ParagraphFormat.LineSpacing = 12
    }
    # 空行（協會資訊與收件人之間的間隔）
    $Sel.TypeParagraph()
    $pRange = $Sel.Paragraphs.Last.Range
    $pRange.Borders.Item(3).LineStyle = 1
    $pRange.Borders.Item(3).LineWidth = 6
    $pRange.Borders.Item(3).ColorIndex = 1
    $pRange.ParagraphFormat.LineSpacingRule = 3
    $pRange.ParagraphFormat.LineSpacing = 12
}

function Sanitize-Filename {
    param([string]$Name)
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $result = $Name
    foreach ($c in $invalid) { $result = $result.Replace([string]$c, '_') }
    return $result
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

    # A4、邊界（配合範本）
    $doc.PageSetup.PaperSize = 1
    $doc.PageSetup.TopMargin    = 28
    $doc.PageSetup.BottomMargin = 8
    $doc.PageSetup.LeftMargin   = 28
    $doc.PageSetup.RightMargin  = 28

    $rowIdx = 0
    foreach ($row in $validRows) {
        $rowIdx++

        if ($rowIdx -gt 1) {
            $doc.Words.Last.InsertBreak(1) | Out-Null
        }

        # ── 解析郵遞區號與地址 ─────────────────────────────────────
        $addrRaw  = if ($row.ADDR) { $row.ADDR.Trim() } else { "" }
        $postalCode = ""
        $address    = $addrRaw
        if ($addrRaw -match '^(\d{3,5})(.+)') {
            $postalCode = $Matches[1]
            $address    = $Matches[2].Trim()
        }
        $name1 = if ($row.Name1) { $row.Name1.Trim() } else { "" }

        # ── 空白区域（上方 1/3，三折第一折）─────────────────────
        Write-BlankArea $sel $doc

        # ── 第一條摺線 ────────────────────────────────────────────
        Write-FoldLine $sel

        # ── 協會資訊 ─────────────────────────────────────────────
        Write-AssociationInfo $sel

        # ── 收件人地址（置中）────────────────────────────────────
        Write-AddressBlock $sel $postalCode $address $name1

        # ── 第二條摺線 ────────────────────────────────────────────
        Write-FoldLine $sel

        # ── 摺線與信文間的空行（標楷體 14pt 粗體, exact 280, 底線）──
        for ($b = 0; $b -lt 2; $b++) {
            $sel.TypeParagraph()
            $bp = $sel.Paragraphs.Last.Range
            $bp.Font.Name = "標楷體"; $bp.Font.Size = 14; $bp.Font.Bold = 1
            $bp.Borders.Item(3).LineStyle = 1; $bp.Borders.Item(3).LineWidth = 6; $bp.Borders.Item(3).ColorIndex = 1
            $bp.ParagraphFormat.LineSpacingRule = 3; $bp.ParagraphFormat.LineSpacing = 14
        }

        # ── 信文 ──────────────────────────────────────────────────
        Write-LetterBody $sel $coopName $reportDate

        # ── 表格 ──────────────────────────────────────────────────
        $share   = if ($row.LGR_Share)   { [long]$row.LGR_Share   } else { 0 }
        $loan    = if ($row.LGR_Loan)    { [long]$row.LGR_Loan    } else { 0 }
        $Reserve = if ($row.LGR_Reserve) { [long]$row.LGR_Reserve } else { 0 }

        Write-StatementTable $doc $sel $share $loan $Reserve

        # ── 頁尾 ──────────────────────────────────────────────────
        $accNo = if ($row.AccNo) { $row.AccNo } else { "" }
        $sel.TypeText("帳號: $accNo                    社員:__________________(簽名或蓋章)             SN:    $rowIdx   ")
        $footerRange = $sel.Paragraphs.Last.Range
        $footerRange.Font.Name = "標楷體"
        $footerRange.Font.Bold = 1
        $footerRange.Borders.Item(3).LineStyle = 1
        $footerRange.Borders.Item(3).LineWidth = 6
        $footerRange.Borders.Item(3).ColorIndex = 1
        $footerRange.ParagraphFormat.LineSpacingRule = 5
        $footerRange.ParagraphFormat.LineSpacing = 2.4
    }

    # ── 存檔 ──────────────────────────────────────────────────────
    $outDir = Join-Path $PSScriptRoot "對帳單"
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

    $docxPath = Join-Path $outDir "對帳單_合併.docx"
    $pdfPath  = Join-Path $outDir "對帳單_合併.pdf"

    # 直接匯出 PDF
    $doc.SaveAs2($pdfPath, 17)
    Write-Host "PDF 已儲存: $pdfPath" -ForegroundColor Green

    # 另存 DOCX
    $doc.SaveAs2($docxPath, 16)
    $f = Get-Item $docxPath -Force
    $f.Attributes = 'Normal'
    Write-Host "Word 已儲存: $docxPath" -ForegroundColor Green

    # ── 個別 PDF（可選）────────────────────────────────────────────
    if ($IndividualPdf) {
        Write-Host "產生個別 PDF..." -ForegroundColor Yellow
        $sn = 0
        foreach ($row in $validRows) {
            $sn++
            $singleDoc = $word.Documents.Add()
            $singleSel = $word.Selection
            $singleDoc.PageSetup.PaperSize = 1
            $singleDoc.PageSetup.TopMargin    = 28
            $singleDoc.PageSetup.BottomMargin = 8
            $singleDoc.PageSetup.LeftMargin   = 28
            $singleDoc.PageSetup.RightMargin  = 28

            $addrRaw2 = if ($row.ADDR) { $row.ADDR.Trim() } else { "" }
            $postalCode2 = ""; $address2 = $addrRaw2
            if ($addrRaw2 -match '^(\d{3,5})(.+)') { $postalCode2 = $Matches[1]; $address2 = $Matches[2].Trim() }
            $name2 = if ($row.Name1) { $row.Name1.Trim() } else { "" }

            Write-BlankArea $singleSel $singleDoc
            Write-FoldLine $singleSel
            Write-AssociationInfo $singleSel
            Write-AddressBlock $singleSel $postalCode2 $address2 $name2
            Write-FoldLine $singleSel
            Write-LetterBody $singleSel $coopName $reportDate

            $s2 = if ($row.LGR_Share)   { [long]$row.LGR_Share   } else { 0 }
            $l2 = if ($row.LGR_Loan)    { [long]$row.LGR_Loan    } else { 0 }
            $r2 = if ($row.LGR_Reserve) { [long]$row.LGR_Reserve } else { 0 }
            Write-StatementTable $singleDoc $singleSel $s2 $l2 $r2

            $singleSel.TypeText("帳號: $($row.AccNo)                    社員:__________________(簽名或蓋章)             SN:    $sn   ")
            $singleFooter = $singleSel.Paragraphs.Last.Range
            $singleFooter.Font.Name = "標楷體"
            $singleFooter.Font.Bold = 1
            $singleFooter.Borders.Item(3).LineStyle = 1
            $singleFooter.Borders.Item(3).LineWidth = 6
            $singleFooter.Borders.Item(3).ColorIndex = 1

            $acc = if ($row.AccNo) { $row.AccNo } else { "000000" }
            $safeName = Sanitize-Filename $name2
            $singlePdf = Join-Path $outDir "SN${sn}_${acc}_${safeName}.pdf"
            $singleDoc.SaveAs2($singlePdf, 17)
            $singleDoc.Close(0)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($singleDoc)

            # 每 20 筆釋放記憶體
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
