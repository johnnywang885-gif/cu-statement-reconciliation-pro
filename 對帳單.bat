@echo off
chcp 950 >nul
setlocal

where pwsh >nul 2>nul
if %errorlevel% neq 0 (
    echo ================================================
    echo    錯誤：您的系統尚未安裝 PowerShell 7
    echo ================================================
    echo.
    echo 此工具需要安裝 PowerShell 7 ^(pwsh.exe^)。
    echo 請至下列網址下載安裝：
    echo https://github.com/PowerShell/PowerShell/releases
    echo.
    echo 安裝完成後請重新執行此程式。
    echo.
    pause
    exit /b
)


set "SELECTED_CUB="
set "DEFAULT_PWD=thifincub"
set "CUB_PWD=%DEFAULT_PWD%"

:SELECT_FILE
for /f "tokens=*" %%i in ('pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0select_cub.ps1"') do set "NEW_CUB=%%i"

if "%NEW_CUB%"=="" (
    if "%SELECTED_CUB%"=="" (
        echo.
        echo 未選擇檔案，程式結束。
        timeout /t 3 /nobreak >nul
        exit /b
    ) else (
        echo.
        echo 維持原選擇，繼續使用檔案：%SELECTED_CUB%
        timeout /t 2 /nobreak >nul
        goto MENU
    )
)
set "SELECTED_CUB=%NEW_CUB%"
set "CUB_PWD=%DEFAULT_PWD%"
goto GET_PASSWORD

:GET_PASSWORD
if not "%CUB_PWD%"=="" goto MENU
cls
echo.
echo ================================================
echo    CUB.MDB 密碼
echo ================================================
echo.
for /f "tokens=*" %%p in ('pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0get_password.ps1"') do set "CUB_PWD=%%p"
if "%CUB_PWD%"=="" (
    echo.
    echo 密碼不可空白，請重新輸入。
    timeout /t 2 /nobreak >nul
    goto GET_PASSWORD
)
goto MENU

:MENU
cls
echo.
echo ================================================
echo    對帳單 - 儲互社資料處理系統
echo ================================================
echo.
echo   目前 CUB.MDB：%SELECTED_CUB%
echo.
echo   1. 異常偵測
echo   2. 百分比篩選
echo   3. 一鍵篩選前5%%（含結餘）
echo   4. 重選 CUB.MDB
echo   5. 產生對帳單 PDF
echo   6. 結束
echo.
echo ================================================
set /p CHOICE=請輸入選項 (1-6):

if "%CHOICE%"=="1" goto RUN_ANOMALY
if "%CHOICE%"=="2" goto FILTER_PERCENT
if "%CHOICE%"=="3" goto RUN_TOP5
if "%CHOICE%"=="4" goto SELECT_FILE
if "%CHOICE%"=="5" goto GENERATE_PDF
if "%CHOICE%"=="6" goto END
goto MENU

:RUN_ANOMALY
cls
echo.
echo 正在執行異常偵測...
echo.
pwsh -ExecutionPolicy Bypass -File "%~dp0find_anomaly_members.ps1" -CubPath "%SELECTED_CUB%" -CubPassword "%CUB_PWD%"
echo.
pause
goto MENU

:FILTER_PERCENT
cls
echo.
echo 正在依比例篩選...
echo.
set /p PCT=請輸入篩選百分比 (1-100):
if "%PCT%"=="" goto FILTER_PERCENT
set /a PCT_NUM=%PCT% 2>nul
if %PCT_NUM% leq 0 goto FILTER_PERCENT
if %PCT_NUM% gtr 100 goto FILTER_PERCENT
pwsh -ExecutionPolicy Bypass -File "%~dp0filter_by_percent.ps1" -Percent %PCT%
echo.
pause
goto MENU

:RUN_TOP5
cls
echo.
echo [1/2] 執行異常偵測...
echo.
pwsh -ExecutionPolicy Bypass -File "%~dp0find_anomaly_members.ps1" -CubPath "%SELECTED_CUB%" -CubPassword "%CUB_PWD%"
if %errorlevel% neq 0 (
    echo.
    echo 異常偵測失敗，中止。
    pause
    goto MENU
)
echo.
echo [2/2] 篩選前5%%...
echo.
pwsh -ExecutionPolicy Bypass -File "%~dp0filter_top5.ps1" -Percent 5
echo.
pause
goto MENU

:GENERATE_PDF
cls
echo.
echo 正在產生對帳單 PDF...
echo.
pwsh -ExecutionPolicy Bypass -File "%~dp0generate_statements.ps1"
echo.
pause
goto MENU

:END
cls
echo.
echo 再見！
timeout /t 2 /nobreak >nul