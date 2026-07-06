@echo off
chcp 950 >nul
setlocal

where pwsh >nul 2>nul
if %errorlevel% neq 0 (
    echo ================================================
    echo    錯誤：偵測到系統未安裝 PowerShell 7
    echo ================================================
    echo.
    echo 執行此工具需要安裝 PowerShell 7 ^(pwsh.exe^)。
    echo 請至微軟官方或 GitHub 下載安裝：
    echo https://github.com/PowerShell/PowerShell/releases
    echo.
    echo 請安裝完成後重新執行此程式。
    echo.
    pause
    exit /b
)

REM 檢查 Access Database Engine 32-bit (DAO.DBEngine.120)
pwsh -NoProfile -Command "try { $null = New-Object -ComObject DAO.DBEngine.120; exit 0 } catch { exit 1 }"
if %errorlevel% neq 0 (
    echo ================================================
    echo    錯誤：偵測不到 Access Database Engine 32-bit
    echo ================================================
    echo.
    echo 這個工具需要 Access Database Engine 2016 (32-bit)。
    echo 請至微軟官網下載並安裝：
    echo https://www.microsoft.com/en-us/download/details.aspx?id=54920
    echo.
    echo 注意：必須下載 32 位元版本（AccessDatabaseEngine.exe，非 _X64）
    echo.
    echo 或執行資料夾中的「安裝環境.bat」來自動安裝。
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
        echo 未選擇檔案，結束。
        timeout /t 3 /nobreak >nul
        exit /b
    ) else (
        echo.
        echo 取消更換，保持原檔案：%SELECTED_CUB%
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
echo    對帳單 - 異常社員篩選系統
echo ================================================
echo.
echo   目前 CUB.MDB：%SELECTED_CUB%
echo.
echo   1. 偵測異常
echo   2. 百分比篩選
echo   3. 選 CUB.MDB
echo   4. 結束
echo.
echo ================================================
set /p CHOICE=請輸入選項 (1-4):

if "%CHOICE%"=="1" goto RUN_ANOMALY
if "%CHOICE%"=="2" goto FILTER_PERCENT
if "%CHOICE%"=="3" goto SELECT_FILE
if "%CHOICE%"=="4" goto END
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
echo 正在依比例過濾...
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
:END
cls
echo.
echo 再見！
timeout /t 2 /nobreak >nul
