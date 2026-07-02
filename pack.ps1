# pack.ps1
# 設定主機輸出編碼為 UTF-8
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 定義工作目錄
$sourceDir = $PSScriptRoot
$zipName = "對帳單工具_v1_發布版.zip"
$zipPath = Join-Path $sourceDir $zipName

# 如果已存在舊的 zip，先刪除
if (Test-Path $zipPath) {
    Remove-Item $zipPath -Force
}

# 建立暫存資料夾用來打包
$tempDirName = "statement_tool_pack_temp_" + (Get-Random)
$tempDir = Join-Path $env:TEMP $tempDirName
if (Test-Path $tempDir) {
    Remove-Item $tempDir -Recurse -Force
}
New-Item -ItemType Directory -Path $tempDir | Out-Null

# 建立 lib 資料夾在暫存區
$tempLibDir = Join-Path $tempDir "lib"
New-Item -ItemType Directory -Path $tempLibDir | Out-Null

# 複製 lib\AnomalyScore.psm1
$libSrc = Join-Path $sourceDir "lib\AnomalyScore.psm1"
if (Test-Path $libSrc) {
    Copy-Item $libSrc -Destination $tempLibDir
}

# 複製對帳單.bat
$batSrc = Join-Path $sourceDir "對帳單.bat"
if (Test-Path $batSrc) {
    Copy-Item $batSrc -Destination $tempDir
}

# 複製使用說明.txt
$readmeSrc = Join-Path $sourceDir "使用說明.txt"
if (Test-Path $readmeSrc) {
    Copy-Item $readmeSrc -Destination $tempDir
}

# 複製所有 .ps1，排除 pack.ps1 自身及開發用診斷腳本
$excludeList = @("pack.ps1", "check_lgr.ps1", "check_lgr2.ps1", "inspect_cub_schema.ps1")
Get-ChildItem -Path $sourceDir -Filter "*.ps1" | ForEach-Object {
    if ($_.Name -notin $excludeList) {
        Copy-Item $_.FullName -Destination $tempDir
    }
}

# 壓縮暫存資料夾內容
# 注意：Compress-Archive 在路徑包含萬用字元時有時在舊版 PowerShell 會有行為差異
# 為了最安全，將 WorkDir 切換到 tempDir 後執行壓縮
$oldLocation = Get-Location
Set-Location $tempDir
try {
    # 壓縮當前目錄下的所有項目
    Compress-Archive -Path * -DestinationPath $zipPath -Force
}
finally {
    Set-Location $oldLocation
}

# 刪除暫存資料夾
if (Test-Path $tempDir) {
    Remove-Item $tempDir -Recurse -Force
}

Write-Host "==================================================" -ForegroundColor Green
Write-Host "打包完成！" -ForegroundColor Green
Write-Host "產出的 ZIP 檔案已置於：" -ForegroundColor Yellow
Write-Host $zipPath -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Green
