@echo off
chcp 950 >nul
setlocal
title 對帳單工具 - 環境安裝

echo ================================================
echo    對帳單工具 - 自動環境安裝
echo ================================================
echo.
echo 本工具將自動安裝所需軟體：
echo  1. PowerShell 7（若尚未安裝）
echo  2. Access Database Engine 2016 32-bit（若尚未安裝）
echo.
echo 安裝過程中會出現「使用者帳戶控制」視窗，
echo 請按「是」以繼續安裝。
echo.

where pwsh >nul 2>nul
if %errorlevel% neq 0 (
    echo [1/2] 正在安裝 PowerShell 7...
    echo.
    winget install --id Microsoft.PowerShell --silent --accept-package-agreements --accept-source-agreements 2>nul
    if %errorlevel% neq 0 (
        echo.
        echo 無法自動安裝 PowerShell 7。
        echo 請手動至以下網址下載安裝：
        echo https://github.com/PowerShell/PowerShell/releases
        echo.
        pause
        exit /b
    )
    echo PowerShell 7 安裝完成！
) else (
    echo [1/2] PowerShell 7 已安裝 - 略過
)

echo.
echo [2/2] 正在檢查 Access Database Engine...
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
echo.
pause