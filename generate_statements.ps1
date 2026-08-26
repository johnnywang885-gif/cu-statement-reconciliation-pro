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

# 32-bit 重啟時從環境變數讀取參數
if ($env:GS_CsvPath)      { $CsvPath = $env:GS_CsvPath }
if ($env:GS_CubPath)      { $CubPath = $env:GS_CubPath }
if ($env:GS_CubPassword)  { $CubPassword = $env:GS_CubPassword }
if ($env:GS_CoopName)     { $CoopName = $env:GS_CoopName }
if ($env:GS_ReportDate)   { $ReportDate = $env:GS_ReportDate }
if ($env:GS_IndividualPdf) { $IndividualPdf = $true }
# 清除環境變數避免殘留
if ($env:GS_CsvPath)      { Remove-Item Env:GS_CsvPath -ErrorAction SilentlyContinue }
if ($env:GS_CubPath)      { Remove-Item Env:GS_CubPath -ErrorAction SilentlyContinue }
if ($env:GS_CubPassword)  { Remove-Item Env:GS_CubPassword -ErrorAction SilentlyContinue }
if ($env:GS_CoopName)     { Remove-Item Env:GS_CoopName -ErrorAction SilentlyContinue }
if ($env:GS_ReportDate)   { Remove-Item Env:GS_ReportDate -ErrorAction SilentlyContinue }
if ($env:GS_IndividualPdf) { Remove-Item Env:GS_IndividualPdf -ErrorAction SilentlyContinue }

Set-Location (Split-Path -Parent $MyInvocation.MyCommand.Path)

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

# DAO 32-bit 初始化
$daoAvailable = $false
try { $null = New-Object -ComObject DAO.DBEngine.120; $daoAvailable = $true } catch {}

if (-not $daoAvailable -and -not $env:DAO_RESTARTED) {
    $ps32 = "$env:SystemRoot\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    if (Test-Path $ps32) {
        Write-Host "  偵測到 32-bit 需求，自動切換..." -ForegroundColor Yellow
        $env:DAO_RESTARTED = '1'
        # 用環境變數傳遞參數（避免 PS 5.1 splatting 問題）
        if ($CsvPath) { $env:GS_CsvPath = $CsvPath }
        if ($CubPath) { $env:GS_CubPath = $CubPath }
        if ($CubPassword) { $env:GS_CubPassword = $CubPassword }
        if ($CoopName) { $env:GS_CoopName = $CoopName }
        if ($ReportDate) { $env:GS_ReportDate = $ReportDate }
        if ($IndividualPdf) { $env:GS_IndividualPdf = '1' }
        & $ps32 -ExecutionPolicy Bypass -File $PSCommandPath
        exit $LASTEXITCODE
    }
}

if ($CubPath -eq "") {
    $CubPath = Join-Path $PSScriptRoot "CUB.MDB"
}
$CubPath = [System.IO.Path]::GetFullPath($CubPath)

if ($daoAvailable -and (Test-Path $CubPath)) {
    try {
        $connectStr = if ($CubPassword) { ";PWD=$CubPassword" } else { "" }
        $dbe = New-Object -ComObject DAO.DBEngine.120
        $db = $dbe.OpenDatabase($CubPath, $false, $true, $connectStr)
        $rs = $db.OpenRecordset("SELECT Item, Para FROM k_para")
        $kPara = @{}
        while (-not $rs.EOF) {
            $kPara[$rs.Fields["Item"].Value] = $rs.Fields["Para"].Value
            $rs.MoveNext()
        }
        $rs.Close(); $db.Close()

        # 日期（eDate）— 格式化為民國年/月/日
        if ($ReportDate -eq "" -and $kPara.ContainsKey("eDate") -and $kPara["eDate"]) {
            $rawDate = $kPara["eDate"]
            # 處理 "1150615" → "115/06/15"、"115/6/15" → "115/06/15"、已是正確格式則保留
            if ($rawDate -match '^\d{7}$') {
                # 7位數如 1150615 → 115/06/15
                $y = $rawDate.Substring(0,3)
                $m = $rawDate.Substring(3,2)
                $d = $rawDate.Substring(5,2)
                $ReportDate = "$y/$m/$d"
            } elseif ($rawDate -match '^\d{8}$') {
                # 8位數如 20260615 → 115/06/15 (需轉換西元→民國)
                $y = [int]$rawDate.Substring(0,4) - 1911
                $m = $rawDate.Substring(4,2)
                $d = $rawDate.Substring(6,2)
                $ReportDate = "$y/$m/$d"
            } else {
                $ReportDate = $rawDate
            }
        }

        # 合作社名（嘗試多個可能的 key）
        foreach ($key in @("CoopName", "SOCNAME", "SocName", "coopname")) {
            if ($kPara.ContainsKey($key) -and $kPara[$key]) {
                $coopName = $kPara[$key]
                break
            }
        }

        # 若 k_para 沒有，嘗試 PARA 表
        if ([string]::IsNullOrEmpty($coopName)) {
            try {
                $db2 = $dbe.OpenDatabase($CubPath, $false, $true, $connectStr)
                $rs2 = $db2.OpenRecordset("SELECT * FROM PARA")
                if (-not $rs2.EOF) {
                    foreach ($fld in $rs2.Fields) {
                        if ($fld.Name -match "NAME|NAME2|SOCNAME|COOP" -and $fld.Value) {
                            $coopName = $fld.Value
                            break
                        }
                    }
                }
                $rs2.Close(); $db2.Close()
            } catch {}
        }
    } catch {
        Write-Host "  讀取 CUB.MDB 時發生錯誤: $($_.Exception.Message)" -ForegroundColor Yellow
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
    if ($Bold) { $r.Bold = $Bold }
    if ($Align) { $r.ParagraphFormat.Alignment = $Align }
    $r.Collapse(0) | Out-Null
}

function Write-FoldLine {
    param($Sel)
    # 新增一個段落，用底部邊框作為摺線
    $Sel.TypeParagraph()
    $pRange = $Sel.Paragraphs.Last.Range
    $pRange.Borders.Item(3).LineStyle = 1   # wdBorderBottom = 3, wdBorderSingle = 1
    $pRange.Borders.Item(3).LineWidth = 6   # 1.5pt
    $pRange.Borders.Item(3).ColorIndex = 1  # wdBlack
}

function Write-BlankArea {
    param($Sel, $Doc)
    # 用空白段落 + 固定行距模擬上方1/3空白（~22行 x 16pt ≈ 352pt ≈ A4 1/3）
    for ($i = 0; $i -lt 22; $i++) { $Sel.TypeParagraph() }
}

function Write-AddressBlock {
    param($Sel, [string]$PostalCode, [string]$Address, [string]$Name)
    # 第一行：郵遞區號 + 地址（置中）
    $addrLine = if ($PostalCode) { "$PostalCode  $Address" } else { $Address }
    $Sel.TypeText($addrLine)
    $Sel.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
    $Sel.TypeParagraph()
    # 第二行：姓名 + 君 啟（置中）
    $Sel.TypeText("$Name 君  啟")
    $Sel.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
    $Sel.TypeParagraph()
    # 重設為靠左
    $Sel.Range.ParagraphFormat.Alignment = 0  # wdAlignParagraphLeft
    $Sel.TypeParagraph()
    $Sel.TypeParagraph()
}

function Write-LetterBody {
    param($Sel, [string]$CoopName, [string]$Date)
    $Sel.TypeText("親愛的社員您好:")
    $Sel.TypeParagraph()
    $Sel.TypeText("　　本對帳單乃由中華民國儲蓄互助協會督導組寄發，係為稽核儲蓄互助社之帳務及保障您的權益所做的例行查核對帳，您在 $CoopName 帳戶至 $Date 止各項餘額如下表。為維護您權益，請按您的存摺金額填入後，將此單以傳真方式或郵寄擲回。謝謝您的合作！")
    $Sel.TypeParagraph()
}

function Write-StatementTable {
    param($Doc, $Sel, [long]$Share, [long]$Loan, [long]$reserve)
    $tableData = @(
        @("股　金",    $share,   "□正確　□錯誤：實有金額　　　　　"),
        @("貸　款",    $loan,    "□正確　□錯誤：實有金額　　　　　"),
        @("備　轉　金", $reserve, "□正確　□錯誤：實有金額　　　　　"),
        @("備　註",    "",       "")
    )

    $numRows = $tableData.Count + 1
    $Sel.Tables.Add($Sel.Range, $numRows, 3) | Out-Null
    $tbl = $Sel.Tables(1)
    $tbl.Borders.Enable = 1
    $tbl.Borders.OutsideLineStyle = 1
    $tbl.Borders.InsideLineStyle   = 1
    $tbl.Columns(1).Width = 110
    $tbl.Columns(2).Width = 90
    $tbl.Columns(3).Width = 268

    # 表頭
    Set-CellText $tbl 1 1 "帳目"                   -Bold 1 -Align 1
    Set-CellText $tbl 1 2 "餘額"                   -Bold 1 -Align 1
    Set-CellText $tbl 1 3 "核 對 勘 誤 回 函 說 明" -Bold 1 -Align 1

    # 資料列
    for ($i = 0; $i -lt $tableData.Count; $i++) {
        $r = $i + 2
        $val = $tableData[$i][1]
        $valStr = if ($val -is [string]) { $val } else { "{0:N0}" -f $val }
        Set-CellText $tbl $r 1 $tableData[$i][0]
        Set-CellText $tbl $r 2 $valStr -Align 2
        Set-CellText $tbl $r 3 $tableData[$i][2]
    }

    $Sel.EndOf(15) | Out-Null
    $Sel.MoveDown() | Out-Null
    $Sel.TypeParagraph()
}

function Write-AssociationInfo {
    param($Sel)
    $Sel.TypeText("404")
    $Sel.TypeParagraph()
    $Sel.TypeText("台中市北區北平路一段33號")
    $Sel.TypeParagraph()
    $Sel.TypeText("中華民國儲蓄互助協會")
    $Sel.TypeParagraph()
    $Sel.TypeText("電話：           04-22917272~8603")
    $Sel.TypeParagraph()
    $Sel.TypeText("傳真：           04-22936903")
    $Sel.TypeParagraph()
    $Sel.TypeParagraph()
}

# ── 產生 Word → PDF ──────────────────────────────────────────────
$word = $null
$doc = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0

    # 設定印表機（避免 PaperSize 錯誤）
    try {
        $word.ActivePrinter = "Microsoft Print to PDF"
        Write-Host "  印表機: Microsoft Print to PDF" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "  未找到 Microsoft Print to PDF，嘗試使用預設印表機" -ForegroundColor Yellow
    }

    $doc = $word.Documents.Add()
    $sel = $word.Selection

    # A4、適當邊界
    $doc.PageSetup.PaperSize = 1
    $doc.PageSetup.TopMargin    = 40
    $doc.PageSetup.BottomMargin = 30
    $doc.PageSetup.LeftMargin   = 45
    $doc.PageSetup.RightMargin  = 45

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

        # ── 信文 ──────────────────────────────────────────────────
        Write-LetterBody $sel $coopName $reportDate

        # ── 表格 ──────────────────────────────────────────────────
        $share   = if ($row.LGR_Share)   { [long]$row.LGR_Share   } else { 0 }
        $loan    = if ($row.LGR_Loan)    { [long]$row.LGR_Loan    } else { 0 }
        $reserve = if ($row.LGR_Reserve) { [long]$row.LGR_Reserve } else { 0 }

        Write-StatementTable $doc $sel $share $loan $reserve

        # ── 頁尾 ──────────────────────────────────────────────────
        $accNo = if ($row.AccNo) { $row.AccNo } else { "" }
        $sel.TypeText("帳號:                     社員:                                                                                SN:    $rowIdx")
        $sel.TypeParagraph()
        $sel.TypeText("           $accNo                                                      (簽名或蓋章)")
    }

    # ── 存檔 ──────────────────────────────────────────────────────
    $outDir = Join-Path $PSScriptRoot "對帳單"
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

    $docxPath = Join-Path $outDir "對帳單_合併.docx"
    $pdfPath  = Join-Path $outDir "對帳單_合併.pdf"

    $doc.SaveAs2($docxPath)
    Write-Host "Word 已儲存: $docxPath" -ForegroundColor Green

    # 另存 PDF（wdExportFormatPDF = 17）
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
            $singleDoc.PageSetup.PaperSize = 1
            $singleDoc.PageSetup.TopMargin    = 40
            $singleDoc.PageSetup.BottomMargin = 30
            $singleDoc.PageSetup.LeftMargin   = 45
            $singleDoc.PageSetup.RightMargin  = 45

            $addrRaw2 = if ($row.ADDR) { $row.ADDR.Trim() } else { "" }
            $postalCode2 = ""; $address2 = $addrRaw2
            if ($addrRaw2 -match '^(\d{3,5})(.+)') { $postalCode2 = $Matches[1]; $address2 = $Matches[2].Trim() }
            $name2 = if ($row.Name1) { $row.Name1.Trim() } else { "" }

            # 空白区域（上方 1/3）
            Write-BlankArea $singleSel $singleDoc

            # 第一條摺線
            Write-FoldLine $singleSel

            # 協會資訊
            Write-AssociationInfo $singleSel

            # 收件人地址（置中）
            Write-AddressBlock $singleSel $postalCode2 $address2 $name2

            # 第二條摺線
            Write-FoldLine $singleSel

            # 信文
            Write-LetterBody $singleSel $coopName $reportDate

            $s2 = if ($row.LGR_Share)   { [long]$row.LGR_Share   } else { 0 }
            $l2 = if ($row.LGR_Loan)    { [long]$row.LGR_Loan    } else { 0 }
            $r2 = if ($row.LGR_Reserve) { [long]$row.LGR_Reserve } else { 0 }
            Write-StatementTable $singleDoc $singleSel $s2 $l2 $r2

            $singleSel.TypeText("帳號:                     社員:                                                                                SN:    $sn")
            $singleSel.TypeParagraph()
            $singleSel.TypeText("           $($row.AccNo)                                                      (簽名或蓋章)")

            $acc = if ($row.AccNo) { $row.AccNo } else { "000000" }
            $singlePdf = Join-Path $outDir "SN${sn}_${acc}_${name2}.pdf"
            $singleDoc.SaveAs2($singlePdf, 17)
            $singleDoc.Close(0)
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($singleDoc)
        }
        Write-Host "  個別 PDF 已產出至 $outDir" -ForegroundColor Green
    }

    $doc.Close(0)
    $doc = $null

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
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Host "Word 已關閉" -ForegroundColor DarkGray
}
