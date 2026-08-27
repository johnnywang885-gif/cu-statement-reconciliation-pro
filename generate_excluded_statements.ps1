# generate_excluded_statements.ps1 — 剔除不寄發對象後產出對帳單 PDF＋明細＋不寄發列表
# 用法:
#   .\generate_excluded_statements.ps1 -CubPassword "thifincub"
#   .\generate_excluded_statements.ps1 -CsvPath "CUB_異常社員_Top5Pct_xxx.csv" -CubPassword "thifincub"
# 前置: 已執行篩選腳本產出 Top5Pct CSV；存在不寄發核對表 docx
# 流程: 比對候選名單 vs 不寄發核對表 → 剔除不寄發 → 產出剔除後對帳單 PDF
#       + 寄發明細(內含不寄發列表) + 同步寫入 CUB.MDB M_對帳明細.不寄發

param(
    [string]$CsvPath = "",
    [string]$NoMailPath = "",
    [string]$CubPath = "",
    [string]$CubPassword = "",
    [string]$CoopName = ""
)

Set-Location (Split-Path -Parent $PSCommandPath)

# ── 預設路徑 ──────────────────────────────────────────────────────
$outDir = Join-Path $PSScriptRoot "對帳單"
if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
if ([string]::IsNullOrEmpty($NoMailPath)) {
    $NoMailPath = Join-Path $outDir "114年社員放棄寄發對帳單核對表-115.docx"
}
if ([string]::IsNullOrEmpty($CubPath)) { $CubPath = Join-Path $PSScriptRoot "CUB.MDB" }
$CubPath = [System.IO.Path]::GetFullPath($CubPath)

# ── 尋找候選 CSV ──────────────────────────────────────────────────
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
    Write-Host "找不到候選 CSV 檔" -ForegroundColor Red; exit 1
}
if (-not (Test-Path $NoMailPath)) {
    Write-Host "找不到不寄發核對表: $NoMailPath" -ForegroundColor Red; exit 1
}
if (-not (Test-Path $CubPath)) {
    Write-Host "找不到 CUB.MDB: $CubPath" -ForegroundColor Red; exit 1
}

# ── 解析不寄發核對表 docx ─────────────────────────────────────────
function Normalize-Account([string]$a) {
    $a = $a.Trim()
    $a = $a -replace '^N', ''      # 去 N 前綴
    $a = $a -replace '^0+', ''     # 去前導零
    return $a
}

Write-Host "解析不寄發核對表: $NoMailPath" -ForegroundColor Yellow
$tmp = Join-Path $env:TEMP ("nomail_{0}" -f [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
try { Expand-Archive -LiteralPath $NoMailPath -DestinationPath $tmp -Force } catch { Write-Host "  無法解壓核對表: $_" -ForegroundColor Red; exit 1 }
$noMailXmlPath = Join-Path $tmp "word/document.xml"
if (-not (Test-Path $noMailXmlPath)) { Write-Host "  核對表內無 document.xml" -ForegroundColor Red; exit 1 }
[xml]$noMailXml = Get-Content -LiteralPath $noMailXmlPath -Raw -Encoding UTF8
$nsMgr = New-Object System.Xml.XmlNamespaceManager($noMailXml.NameTable)
$nsMgr.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
$tbls = $noMailXml.SelectNodes("//w:tbl", $nsMgr)
if ($tbls.Count -eq 0) { Write-Host "  核對表內無表格" -ForegroundColor Red; exit 1 }
$tbl = $tbls[0]
$noMailSet = New-Object 'System.Collections.Generic.HashSet[string]'
$noMailList = @()
foreach ($tr in $tbl.SelectNodes("w:tr", $nsMgr)) {
    $tcs = $tr.SelectNodes("w:tc", $nsMgr)
    for ($i = 0; $i -lt $tcs.Count; $i += 2) {
        if (($i + 1) -ge $tcs.Count) { continue }
        $acc = (($tcs[$i].SelectNodes(".//w:t", $nsMgr)) | ForEach-Object { $_.InnerText }) -join ""
        $nm  = (($tcs[$i + 1].SelectNodes(".//w:t", $nsMgr)) | ForEach-Object { $_.InnerText }) -join ""
        if ($acc.Trim() -ne '') {
            $n = Normalize-Account $acc
            if ($noMailSet.Add($n)) {
                $noMailList += [pscustomobject]@{ AccNo = $acc.Trim(); Name1 = $nm.Trim() }
            }
        }
    }
}
Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "  不寄發帳號: $($noMailList.Count) 筆" -ForegroundColor DarkGray

# ── 讀候選 CSV 並比對 ─────────────────────────────────────────────
Write-Host "讀取候選名單: $CsvPath" -ForegroundColor Yellow
$csvFile = Import-Csv $CsvPath -Encoding utf8
$candidates = @($csvFile | Where-Object { $_.AccNo -and $_.AccNo -ne "" } | Sort-Object AccNo)
Write-Host "  候選: $($candidates.Count) 筆" -ForegroundColor DarkGray

$excluded = @()
$kept = @()
foreach ($c in $candidates) {
    $n = Normalize-Account $c.AccNo
    if ($noMailSet.Contains($n)) { $excluded += $c } else { $kept += $c }
}
Write-Host "  剔除不寄發: $($excluded.Count) 筆；保留寄發: $($kept.Count) 筆" -ForegroundColor Yellow

# ── 產出中間 CSV ──────────────────────────────────────────────────
$keptCsvPath = Join-Path $outDir "寄發對象_剔除後.csv"
$excludedCsvPath = Join-Path $outDir "不寄發對象.csv"
$kept | Export-Csv -LiteralPath $keptCsvPath -Encoding utf8BOM -NoTypeInformation
$excluded | ForEach-Object {
    [pscustomobject]@{ AccNo = $_.AccNo; Name1 = ($_.Name1).Trim(); 原因 = "放棄寄發" }
} | Export-Csv -LiteralPath $excludedCsvPath -Encoding utf8BOM -NoTypeInformation
Write-Host "  已寫入: $keptCsvPath" -ForegroundColor DarkGray
Write-Host "  已寫入: $excludedCsvPath" -ForegroundColor DarkGray

# ── 呼叫既有腳本 ──────────────────────────────────────────────────
$stmtOut = Join-Path $outDir "對帳單_合併_剔除後.pdf"
$detailOut = Join-Path $outDir "對帳單_寄發明細_剔除後.pdf"

$baseArgs = @('-ExecutionPolicy', 'Bypass', '-NoProfile')

Write-Host "`n[1/2] 產生剔除後對帳單 PDF ($($kept.Count) 筆)..." -ForegroundColor Cyan
$stmtArgs = $baseArgs + @('-File', (Join-Path $PSScriptRoot 'generate_statements.ps1'),
    '-CsvPath', $keptCsvPath, '-OutPath', $stmtOut)
if (-not [string]::IsNullOrEmpty($CubPassword)) { $stmtArgs += @('-CubPassword', $CubPassword) }
if (-not [string]::IsNullOrEmpty($CubPath)) { $stmtArgs += @('-CubPath', $CubPath) }
if (-not [string]::IsNullOrEmpty($CoopName)) { $stmtArgs += @('-CoopName', $CoopName) }
& pwsh @stmtArgs
if ($LASTEXITCODE -ne 0) { Write-Host "  generate_statements 失敗" -ForegroundColor Red }

Write-Host "`n[2/2] 產生剔除後寄發明細 (含不寄發列表)..." -ForegroundColor Cyan
$detArgs = $baseArgs + @('-File', (Join-Path $PSScriptRoot 'generate_detail_list.ps1'),
    '-CsvPath', $keptCsvPath, '-OutPath', $detailOut, '-ExcludedCsv', $excludedCsvPath)
if (-not [string]::IsNullOrEmpty($CubPassword)) { $detArgs += @('-CubPassword', $CubPassword) }
if (-not [string]::IsNullOrEmpty($CubPath)) { $detArgs += @('-CubPath', $CubPath) }
& pwsh @detArgs
if ($LASTEXITCODE -ne 0) { Write-Host "  generate_detail_list 失敗" -ForegroundColor Red }

# ── 同步寫入 CUB.MDB M_對帳明細.不寄發 ───────────────────────────
Write-Host "`n同步寫入 CUB.MDB M_對帳明細.不寄發..." -ForegroundColor Cyan
if ($excluded.Count -gt 0) {
    $dbScript = @'
param([string]$CubPath,[string]$CubPassword,[string]$AccList)
$ErrorActionPreference = "Stop"
try {
    $AccNos = @($AccList -split ',') | Where-Object { $_ -ne '' }
    $dbe = New-Object -ComObject DAO.DBEngine.120
    $db = $dbe.OpenDatabase($CubPath, $false, $false, ";PWD=$CubPassword")
    foreach($a in $AccNos){
        $qdef = $db.CreateQueryDef("", "UPDATE [M_對帳明細] SET 不寄發=True, 不寄發原因='放棄寄發' WHERE 帳號=[p_acc]")
        $qdef.Parameters("p_acc").Value = $a
        $qdef.Execute()
        $affected = $qdef.RecordsAffected
        if($affected -gt 0){ Write-Output ("OK " + $a + " (影響 " + $affected + " 列)") }
        else { Write-Output ("SKIP " + $a + " (不在 M_對帳明細)") }
    }
    $db.Close()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($dbe)
} catch { Write-Output ("ERR " + $_.Exception.Message) }
'@
    $dbScriptPath = Join-Path $env:TEMP ("nomail_db_{0}.ps1" -f [Guid]::NewGuid().ToString("N"))
    Set-Content -LiteralPath $dbScriptPath -Value $dbScript -Encoding utf8BOM
    $accList = ($excluded | ForEach-Object { $_.AccNo.Trim() }) -join ','
    $ps32 = "$env:SystemRoot\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    $dbOut = & $ps32 -ExecutionPolicy Bypass -File $dbScriptPath -CubPath $CubPath -CubPassword $CubPassword -AccList $accList 2>&1
    $dbOut | ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
    Remove-Item $dbScriptPath -Force -ErrorAction SilentlyContinue
} else {
    Write-Host "  無剔除對象，不需同步" -ForegroundColor DarkGray
}

# ── 摘要 ──────────────────────────────────────────────────────────
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  候選 $($candidates.Count) → 剔除 $($excluded.Count) → 寄發 $($kept.Count)" -ForegroundColor Cyan
Write-Host "  剔除後對帳單: $stmtOut" -ForegroundColor Green
Write-Host "  剔除後明細(含不寄發列表): $detailOut" -ForegroundColor Green
Write-Host "  不寄發清單: $excludedCsvPath" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan

