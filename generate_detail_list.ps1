# generate_detail_list.ps1 — 寄發對帳單明細表（依 Top5Pct CSV 與 PDF 範本）
# 用法:
#   .\generate_detail_list.ps1 -CsvPath "CUB_異常社員_Top5Pct_xxx.csv" -CubPassword "thifincub"
#   .\generate_detail_list.ps1  # 自動找最新 Top5Pct

param(
    [string]$CsvPath = "",
    [string]$CubPath = "",
    [string]$CubPassword = "",
    [string]$OutPath = "",
    [string]$ExcludedCsv = ""
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
        $f = Get-ChildItem $p -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($f) { $CsvPath = $f.FullName; break }
    }
}
if ([string]::IsNullOrEmpty($CsvPath) -or -not (Test-Path $CsvPath)) {
    Write-Host "找不到 CSV 檔" -ForegroundColor Red; exit 1
}
Write-Host "讀取 CSV: $CsvPath" -ForegroundColor Yellow
$csvFile = Import-Csv $CsvPath -Encoding utf8
$validRows = @($csvFile | Where-Object { $_.AccNo -and $_.AccNo -ne "" } | Sort-Object AccNo)
if ($validRows.Count -eq 0) { Write-Host "無有效資料" -ForegroundColor Red; exit 1 }
Write-Host "  共 $($validRows.Count) 筆" -ForegroundColor DarkGray

# ── 讀取 CUB 日期 ────────────────────────────────────────────────
if ($CubPath -eq "") { $CubPath = Join-Path $PSScriptRoot "CUB.MDB" }
$CubPath = [System.IO.Path]::GetFullPath($CubPath)
$reportDate = ""  # 基準日期 (eDate)
$sendDate = ""    # 寄發日期 (當天 ROC)

# 嘗試讀 eDate
$daoAvailable = $false
try { $null = New-Object -ComObject DAO.DBEngine.120; $daoAvailable = $true } catch {}
if ($daoAvailable -and (Test-Path $CubPath)) {
    try {
        $cs = if ($CubPassword) { ";PWD=$CubPassword" } else { "" }
        $dbe = New-Object -ComObject DAO.DBEngine.120
        $db = $dbe.OpenDatabase($CubPath, $false, $true, $cs)
        $rs = $db.OpenRecordset("SELECT Para FROM k_para WHERE Item='eDate'")
        if (-not $rs.EOF) { $reportDate = [string]$rs.Fields["Para"].Value }
        $rs.Close(); $db.Close()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($dbe)
    } catch { Write-Host "  DAO 讀取 eDate 失敗: $($_.Exception.Message)" -ForegroundColor Yellow }
}
if ([string]::IsNullOrEmpty($reportDate)) {
    $ps32 = "$env:SystemRoot\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    $infoScript = Join-Path $PSScriptRoot "read_cub_info.ps1"
    if ((Test-Path $ps32) -and (Test-Path $infoScript)) {
        $infoArgs = @('-ExecutionPolicy','Bypass','-File',$infoScript,'-CubPath',$CubPath)
        if ($CubPassword) { $infoArgs += @('-CubPassword',$CubPassword) }
        $jsonRaw = & $ps32 @infoArgs 2>$null
        if ($jsonRaw) {
            $j = $jsonRaw | ConvertFrom-Json
            if ($j.rawDate) { $reportDate = $j.rawDate }
        }
    }
}
# 轉 ROC 格式
function ConvertTo-RocDate($raw) {
    if ($raw -match '^\d{7}$') { return "$($raw.Substring(0,3))/$($raw.Substring(3,2))/$($raw.Substring(5,2))" }
    elseif ($raw -match '^\d{8}$') { $y=[int]$raw.Substring(0,4)-1911; return "$y/$($raw.Substring(4,2))/$($raw.Substring(6,2))" }
    else { return $raw }
}
if ($reportDate -match '^\d{7,8}$') { $reportDate = ConvertTo-RocDate $reportDate }
if ([string]::IsNullOrEmpty($reportDate)) { $reportDate = "115/06/15" }

# 寄發日期：當天 ROC
$now = Get-Date
$rocY = $now.Year - 1911
$sendDate = "{0}/{1:00}/{2:00}" -f $rocY, $now.Month, $now.Day
# 若需與 PDF 範本一致 115/07/07，可改固定：
# $sendDate = "115/07/07"
Write-Host "  基準日期: $reportDate  寄發日期: $sendDate" -ForegroundColor DarkGray

if ([string]::IsNullOrEmpty($OutPath)) {
    $OutPath = Join-Path (Join-Path $PSScriptRoot "對帳單") "對帳單_寄發明細.pdf"
}
$outDir = Split-Path $OutPath -Parent
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

# ── Preflight 清理 Word ──────────────────────────────────────────
function Clear-StaleWordLocks {
    try {
        Get-ChildItem -Path $PSScriptRoot -Filter "~`$*.docx" -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
        Get-ChildItem -Path $outDir -Filter "~`$*" -Force -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
    } catch {}
}
function Stop-StaleWord {
    try {
        $procs = Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'" -ErrorAction SilentlyContinue
        foreach ($p in $procs) {
            $isAutomation = $p.CommandLine -match "/Automation"
            $hasTitle = $false
            try { $po = Get-Process -Id $p.ProcessId -ErrorAction SilentlyContinue; if ($po.MainWindowTitle) { $hasTitle=$true } } catch {}
            if ($isAutomation -or -not $hasTitle) { try { Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue } catch {} }
        }
        Start-Sleep -Seconds 1
    } catch {}
}
Clear-StaleWordLocks; Stop-StaleWord

# ── 產生 Word → PDF ──────────────────────────────────────────────
$wNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
$rowsPerPage = 45

# 建立臨時 DOCX 用 Open XML（複製範本結構簡化）
# 為求快速與穩定，直接用 Word COM 建立明細表（40 筆/頁，分頁）
$word = $null; $doc = $null; $scriptSuccess=$false
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    try { $word.AutomationSecurity = 3 } catch {}
    try { $word.ScreenUpdating = $false } catch {}

    $doc = $word.Documents.Add()
    $sel = $word.Selection

    # 版面：A4 窄邊界（縮小上下邊界以容納更多列）
    $doc.PageSetup.TopMargin = 28
    $doc.PageSetup.BottomMargin = 28
    $doc.PageSetup.LeftMargin = 36
    $doc.PageSetup.RightMargin = 36
    $doc.PageSetup.HeaderDistance = 14

    $totalPages = [Math]::Ceiling($validRows.Count / $rowsPerPage)
    Write-Host "  產生 $totalPages 頁，每頁 $rowsPerPage 筆..." -ForegroundColor Yellow

    for ($pageIdx=0; $pageIdx -lt $totalPages; $pageIdx++) {
        $start = $pageIdx * $rowsPerPage
        $end = [Math]::Min($start + $rowsPerPage -1, $validRows.Count-1)
        $pageRows = $validRows[$start..$end]
        $pageCount = $pageRows.Count
        $cumulative = $end + 1

        # ── 標題 ──────────────────────────────────────────────
        if ($pageIdx -gt 0) {
            # 分頁
            $sel.InsertBreak(7) # wdPageBreak
        }
        # 合作社名（縮小上下間距）
        $p = $sel.Paragraphs.Last.Range
        $p.ParagraphFormat.Alignment = 1
        $p.ParagraphFormat.SpaceAfter = 0
        $p.ParagraphFormat.SpaceBefore = 0
        $p.ParagraphFormat.LineSpacingRule = 4; $p.ParagraphFormat.LineSpacing = 12
        $p.Font.Name = "標楷體"; $p.Font.Size = 14; $p.Font.Bold = $true
        $sel.TypeText("南投縣 十方 儲蓄互助社")
        $sel.TypeParagraph()
        # 明細標題（緊貼）
        $p = $sel.Paragraphs.Last.Range
        $p.ParagraphFormat.Alignment = 1
        $p.ParagraphFormat.SpaceAfter = 2
        $p.ParagraphFormat.SpaceBefore = 0
        $p.ParagraphFormat.LineSpacingRule = 4; $p.ParagraphFormat.LineSpacing = 12
        $p.Font.Name = "標楷體"; $p.Font.Size = 14; $p.Font.Bold = $true
        $sel.TypeText("對帳單明細")
        $sel.TypeParagraph()
        # 頁次（靠右，緊貼表格）
        $p = $sel.Paragraphs.Last.Range
        $p.ParagraphFormat.Alignment = 2
        $p.ParagraphFormat.SpaceAfter = 2
        $p.ParagraphFormat.SpaceBefore = 0
        $p.ParagraphFormat.LineSpacingRule = 4; $p.ParagraphFormat.LineSpacing = 10
        $p.Font.Name = "標楷體"; $p.Font.Size = 9
        $sel.TypeText("頁次： $($pageIdx+1) / $totalPages")
        $sel.TypeParagraph()
        # 表格前再壓縮一段
        $p = $sel.Paragraphs.Last.Range
        $p.ParagraphFormat.SpaceAfter = 0
        $p.ParagraphFormat.SpaceBefore = 0
        $p.ParagraphFormat.LineSpacingRule = 4; $p.ParagraphFormat.LineSpacing = 2

        # ── 表格 ──────────────────────────────────────────────
        # 7 欄：帳號, 姓名, 寄發日期, 回收日期, 正確, 基準日期, 更正事項
        $cols = 7
        # 表格列數：1 表頭 + $pageCount 資料列 + 2 統計列 (小計/合計)
        $tblRows = 1 + $pageCount + 2
        $sel.Tables.Add($sel.Range, $tblRows, $cols) | Out-Null
        $tbl = $doc.Tables($doc.Tables.Count)
        $tbl.Borders.Enable = 1
        $tbl.Borders.OutsideLineStyle = 1
        $tbl.Borders.InsideLineStyle = 1
        $tbl.AllowAutoFit = $false
        # 欄寬 (總寬約 540pt)
        $widths = @(60, 80, 75, 75, 45, 75, 130)
        for ($c=1; $c -le $cols; $c++) { $tbl.Columns($c).Width = $widths[$c-1] }
        # 列高固定（縮小以容納 45 列）
        for ($r=1; $r -le $tblRows; $r++) {
            $tbl.Rows($r).Height = 12
            $tbl.Rows($r).HeightRule = 2 # wdRowHeightExactly
        }

        # 表頭
        $headers = @("帳號","姓名","寄發日期","回收日期","正確","基準日期","更正事項")
        for ($c=1; $c -le $cols; $c++) {
            $cell = $tbl.Cell(1,$c).Range
            $cell.Text = $headers[$c-1]
            $cell.Font.Name="標楷體"; $cell.Font.Size=9; $cell.Font.Bold=$true
            $cell.ParagraphFormat.Alignment=1
            $cell.ParagraphFormat.SpaceAfter=0
            $cell.ParagraphFormat.SpaceBefore=0
            $cell.ParagraphFormat.LineSpacingRule=4; $cell.ParagraphFormat.LineSpacing=11
        }
        # 資料列
        for ($r=0; $r -lt $pageCount; $r++) {
            $row = $pageRows[$r]
            $acc = $row.AccNo
            $name = $row.Name1.Trim()
            $vals = @($acc, $name, $sendDate, "", "", $reportDate, "")
            for ($c=1; $c -le $cols; $c++) {
                $cell = $tbl.Cell(2+$r,$c).Range
                $cell.Text = $vals[$c-1]
                $cell.Font.Name="標楷體"; $cell.Font.Size=9
                $cell.ParagraphFormat.Alignment=1
                $cell.ParagraphFormat.SpaceAfter=0
                $cell.ParagraphFormat.SpaceBefore=0
                $cell.ParagraphFormat.LineSpacingRule=4; $cell.ParagraphFormat.LineSpacing=11
                # 帳號靠左，姓名靠左，其餘置中
                if ($c -eq 1 -or $c -eq 2) { $cell.ParagraphFormat.Alignment=0 }
            }
        }
        # 小計 / 合計 列
        $rowSmall = 2 + $pageCount
        $rowTotal = 3 + $pageCount
        function Set-DetailCell {
            param($CellRange, [int]$Align=1, [int]$Bold=0)
            $CellRange.Font.Name="標楷體"; $CellRange.Font.Size=9; $CellRange.Font.Bold=$Bold
            $CellRange.ParagraphFormat.Alignment=$Align
            $CellRange.ParagraphFormat.SpaceAfter=0
            $CellRange.ParagraphFormat.SpaceBefore=0
            $CellRange.ParagraphFormat.LineSpacingRule=4; $CellRange.ParagraphFormat.LineSpacing=11
        }
        # 小計列：僅顯示小計數量
        Set-DetailCell -CellRange $tbl.Cell($rowSmall,1).Range -Align 2 -Bold 1
        $tbl.Cell($rowSmall,1).Range.Text = "小計："
        Set-DetailCell -CellRange $tbl.Cell($rowSmall,2).Range
        $tbl.Cell($rowSmall,2).Range.Text = "$pageCount"
        # 合計列
        Set-DetailCell -CellRange $tbl.Cell($rowTotal,1).Range -Align 2 -Bold 1
        $tbl.Cell($rowTotal,1).Range.Text = "合計："
        Set-DetailCell -CellRange $tbl.Cell($rowTotal,2).Range
        $tbl.Cell($rowTotal,2).Range.Text = "$cumulative"
        # 其餘儲存格留空
        for ($c=3; $c -le $cols; $c++) {
            Set-DetailCell -CellRange $tbl.Cell($rowSmall,$c).Range
            Set-DetailCell -CellRange $tbl.Cell($rowTotal,$c).Range
        }

        # 移到表格後
        $sel.EndKey(6) | Out-Null # wdStory
        $sel.TypeParagraph()
    }

    # ── 不寄發對帳單對象列表（若有提供 ExcludedCsv）──────────────
    if (-not [string]::IsNullOrEmpty($ExcludedCsv) -and (Test-Path $ExcludedCsv)) {
        $excludedRows = @(Import-Csv $ExcludedCsv -Encoding utf8 | Where-Object { $_.AccNo -and $_.AccNo -ne "" })
        if ($excludedRows.Count -gt 0) {
            Write-Host "  追加不寄發對帳單對象列表 ($($excludedRows.Count) 筆)..." -ForegroundColor DarkGray
            # 分頁
            $sel.InsertBreak(7) # wdPageBreak
            # 標題
            $p = $sel.Paragraphs.Last.Range
            $p.ParagraphFormat.Alignment = 1
            $p.ParagraphFormat.SpaceAfter = 6
            $p.ParagraphFormat.SpaceBefore = 0
            $p.Font.Name = "標楷體"; $p.Font.Size = 14; $p.Font.Bold = $true
            $sel.TypeText("不寄發對帳單對象列表")
            $sel.TypeParagraph()
            # 列表表格：帳號 | 姓名 | 原因
            $exTblRows = 1 + $excludedRows.Count
            $sel.Tables.Add($sel.Range, $exTblRows, 3) | Out-Null
            $exTbl = $doc.Tables($doc.Tables.Count)
            $exTbl.Borders.Enable = 1
            $exTbl.Borders.OutsideLineStyle = 1
            $exTbl.Borders.InsideLineStyle = 1
            $exTbl.AllowAutoFit = $false
            $exWidths = @(100, 200, 240)
            for ($c=1; $c -le 3; $c++) { $exTbl.Columns($c).Width = $exWidths[$c-1] }
            for ($r=1; $r -le $exTblRows; $r++) {
                $exTbl.Rows($r).Height = 14
                $exTbl.Rows($r).HeightRule = 2
            }
            $exHeaders = @("帳號","姓名","原因")
            for ($c=1; $c -le 3; $c++) {
                $cell = $exTbl.Cell(1,$c).Range
                $cell.Text = $exHeaders[$c-1]
                $cell.Font.Name="標楷體"; $cell.Font.Size=10; $cell.Font.Bold=$true
                $cell.ParagraphFormat.Alignment=1
                $cell.ParagraphFormat.SpaceAfter=0; $cell.ParagraphFormat.SpaceBefore=0
                $cell.ParagraphFormat.LineSpacingRule=4; $cell.ParagraphFormat.LineSpacing=11
            }
            for ($r=0; $r -lt $excludedRows.Count; $r++) {
                $er = $excludedRows[$r]
                $eName = if ($er.Name1) { $er.Name1.Trim() } else { "" }
                $eReason = if ($er.原因) { $er.原因.Trim() } else { "放棄寄發" }
                $evals = @($er.AccNo, $eName, $eReason)
                for ($c=1; $c -le 3; $c++) {
                    $cell = $exTbl.Cell(2+$r,$c).Range
                    $cell.Text = $evals[$c-1]
                    $cell.Font.Name="標楷體"; $cell.Font.Size=10
                    $cell.ParagraphFormat.Alignment=1
                    $cell.ParagraphFormat.SpaceAfter=0; $cell.ParagraphFormat.SpaceBefore=0
                    $cell.ParagraphFormat.LineSpacingRule=4; $cell.ParagraphFormat.LineSpacing=11
                    if ($c -eq 1 -or $c -eq 2) { $cell.ParagraphFormat.Alignment=0 }
                }
            }
            $sel.EndKey(6) | Out-Null
            $sel.TypeParagraph()
        }
    }

    # ── 儲存 ──────────────────────────────────────────────────
    $docxPath = [System.IO.Path]::ChangeExtension($OutPath, ".docx")
    # 先存 DOCX 到本機暫存
    $tmpDocx = Join-Path $env:TEMP ("detail_{0}.docx" -f [Guid]::NewGuid().ToString("N"))
    $doc.SaveAs2($tmpDocx, 16)
    $doc.Close(0); $doc=$null
    # 搬回正式位置（含重試，避免雲端鎖定）
    $finalDocx = $docxPath
    try {
        if (Test-Path $finalDocx) { try { Remove-Item $finalDocx -Force -ErrorAction Stop } catch { Start-Sleep -Seconds 1; Remove-Item $finalDocx -Force -ErrorAction SilentlyContinue } }
        Move-Item -LiteralPath $tmpDocx -Destination $finalDocx -Force -ErrorAction Stop
    } catch {
        $finalDocx = Join-Path $outDir ("對帳單_寄發明細_{0}.docx" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
        try { Move-Item -LiteralPath $tmpDocx -Destination $finalDocx -Force -ErrorAction Stop } catch { Copy-Item -LiteralPath $tmpDocx -Destination $finalDocx -Force }
        Write-Host "  原 DOCX 被佔用，改存為: $finalDocx" -ForegroundColor Yellow
    }
    $docxPath = $finalDocx
    Write-Host "DOCX 已儲存: $docxPath" -ForegroundColor Green

    # 轉 PDF（重新開啟 DOCX）
    $tmpPdf = Join-Path $env:TEMP ("detail_{0}.pdf" -f [Guid]::NewGuid().ToString("N"))
    $doc = $word.Documents.Open($docxPath, $false, $true)
    try { $doc.SaveAs2($tmpPdf, 17) } catch { $doc.ExportAsFixedFormat($tmpPdf, 17) }
    $doc.Close(0); $doc=$null
    # 搬回
    if (Test-Path $OutPath) { try { Remove-Item $OutPath -Force } catch { Start-Sleep -Seconds 1; try { Remove-Item $OutPath -Force } catch {} } }
    $finalPdf = $OutPath
    try { Move-Item -LiteralPath $tmpPdf -Destination $finalPdf -Force -ErrorAction Stop }
    catch {
        $finalPdf = Join-Path $outDir ("對帳單_寄發明細_{0}.pdf" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
        Move-Item -LiteralPath $tmpPdf -Destination $finalPdf -Force
        Write-Host "  原 PDF 被佔用，改存為: $finalPdf" -ForegroundColor Yellow
    }
    Write-Host "PDF 已儲存: $finalPdf" -ForegroundColor Green
    $scriptSuccess = $true
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  完成！共 $($validRows.Count) 筆，分 $totalPages 頁" -ForegroundColor Cyan
    Write-Host "  明細 PDF: $finalPdf" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
} catch {
    Write-Host "錯誤: $_" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
} finally {
    if ($doc) { try { $doc.Close(0) } catch {} }
    if ($word) { try { $word.Quit() } catch {}; [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($scriptSuccess) { Write-Host "Word 已關閉" -ForegroundColor DarkGray }
}


