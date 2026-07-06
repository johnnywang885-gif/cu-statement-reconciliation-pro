# install.ps1 — 自動安裝 Access Database Engine 2016 32-bit

# 檢查是否已安裝（用登錄，避開 64-bit 無法建立 32-bit COM 的問題）
$regPaths = @(
    "HKLM:\SOFTWARE\Classes\DAO.DBEngine.120",
    "HKLM:\SOFTWARE\Classes\WOW6432Node\DAO.DBEngine.120",
    "HKLM:\SOFTWARE\Microsoft\Office\16.0\Access Connectivity Engine",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Access Connectivity Engine",
    "HKLM:\SOFTWARE\Microsoft\Office\14.0\Access Connectivity Engine",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\14.0\Access Connectivity Engine"
)
$daoInstalled = $false
foreach ($p in $regPaths) {
    if (Test-Path $p) { $daoInstalled = $true; break }
}

if ($daoInstalled) {
    Write-Host "Access Database Engine 2016 32-bit 已安裝，無需動作。" -ForegroundColor Green
    exit 0
} else {
    Write-Host "Access Database Engine 2016 32-bit 尚未安裝，開始下載..." -ForegroundColor Yellow
}

$url = "https://download.microsoft.com/download/3/5/C/35C84BD8-6F30-4B22-9AB8-07C97E6A4008/AccessDatabaseEngine.exe"
$tmpPath = "$env:TEMP\AccessDatabaseEngine.exe"

try {
    Write-Host "正在從微軟下載安裝程式..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri $url -OutFile $tmpPath -UseBasicParsing
    Write-Host "下載完成。" -ForegroundColor Green
} catch {
    Write-Host "下載失敗，將開啟瀏覽器請您手動下載..." -ForegroundColor Red
    Start-Process "https://www.microsoft.com/en-us/download/details.aspx?id=54920"
    exit 1
}

Write-Host "正在安裝（需要管理員權限，請按「是」）..." -ForegroundColor Yellow
try {
    Start-Process -FilePath $tmpPath -ArgumentList "/quiet" -Wait -Verb RunAs
    # 安裝後用登錄驗證
    $installed = $false
    foreach ($p in $regPaths) { if (Test-Path $p) { $installed = $true; break } }
    if ($installed) {
        Write-Host "Access Database Engine 2016 32-bit 安裝成功！" -ForegroundColor Green
    } else {
        throw "安裝驗證失敗"
    }
} catch {
    Write-Host "安裝失敗，請手動安裝。" -ForegroundColor Red
    Write-Host "請至微軟官網下載 32-bit 版本：" -ForegroundColor Yellow
    Write-Host "https://www.microsoft.com/en-us/download/details.aspx?id=54920" -ForegroundColor Cyan
    Write-Host "注意：必須下載 AccessDatabaseEngine.exe (非 _X64 版本)" -ForegroundColor Yellow
    exit 1
} finally {
    if (Test-Path $tmpPath) { Remove-Item $tmpPath -Force }
}